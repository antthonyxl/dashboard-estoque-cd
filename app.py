from __future__ import annotations

import html
import io
import re
import unicodedata
from pathlib import Path

import openpyxl
import pandas as pd
import streamlit as st


APP_DIR = Path(__file__).resolve().parent
DEFAULT_FILES = [APP_DIR / "data" / "Analise CD GERAL.xlsx", APP_DIR / "data" / "Analise CD 989.xlsx"]
CATEGORIES = ["Ajustes Lojas", "Transferências", "Devoluções", "Inventário"]


def normalize_label(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(char for char in text if not unicodedata.combining(char))
    text = re.sub(r"\s+", " ", text).strip().casefold()
    return text


def normalize_code(value: object) -> str | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)):
        return str(int(value)) if float(value).is_integer() else str(value).strip()
    text = str(value).strip()
    if not text:
        return None
    return re.sub(r"\.0$", "", text)


def first_valid(series: pd.Series) -> str:
    values = series.dropna().astype(str).str.strip()
    values = values[values.ne("")]
    return values.iloc[0] if not values.empty else "Produto sem descrição"


def read_contiguous_sheet(ws: openpyxl.worksheet.worksheet.Worksheet) -> pd.DataFrame:
    """Lê apenas a área útil e evita percorrer linhas formatadas até 1.048.576."""
    rows: list[tuple] = []
    empty_streak = 0
    for row in ws.iter_rows(values_only=True):
        is_empty = all(value is None for value in row)
        if is_empty:
            empty_streak += 1
        else:
            empty_streak = 0
            rows.append(row)
        if rows and empty_streak >= 25:
            break

    if not rows:
        return pd.DataFrame()

    headers = [str(value).strip() if value is not None else f"coluna_{index}" for index, value in enumerate(rows[0])]
    frame = pd.DataFrame(rows[1:], columns=headers).dropna(how="all")
    frame.columns = [str(column).strip() for column in frame.columns]
    return frame


@st.cache_data(show_spinner=False)
def load_workbook(file_bytes: bytes) -> dict[str, pd.DataFrame]:
    workbook = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    frames: dict[str, pd.DataFrame] = {}
    try:
        for sheet_name in workbook.sheetnames:
            frames[sheet_name.strip()] = read_contiguous_sheet(workbook[sheet_name])
    finally:
        workbook.close()
    return frames


def find_sheet(frames: dict[str, pd.DataFrame], expected: str) -> pd.DataFrame:
    normalized_expected = normalize_label(expected)
    for name, frame in frames.items():
        if normalize_label(name) == normalized_expected:
            return frame.copy()
    raise ValueError(f'A aba obrigatória "{expected}" não foi encontrada.')


def standardize_columns(frame: pd.DataFrame) -> pd.DataFrame:
    aliases = {
        "codigo produto": "codigo_produto",
        "produto": "produto",
        "quantidade em estoque": "quantidade_estoque",
        "qtd diferenca": "quantidade_mov",
        "valor custo liquido": "valor_liquido",
        "valor custo bruto": "valor_bruto",
        "empresa": "empresa",
        "c.g.o.": "cgo",
        "codigo c.g.o.": "codigo_cgo",
        "codigo empresa": "codigo_empresa",
    }
    renamed = {}
    for column in frame.columns:
        normalized = normalize_label(column)
        renamed[column] = aliases.get(normalized, normalized.replace(" ", "_"))
    return frame.rename(columns=renamed)


def numeric(series: pd.Series | None, index: pd.Index) -> pd.Series:
    if series is None:
        return pd.Series(float("nan"), index=index, dtype="float64")
    return pd.to_numeric(series, errors="coerce")


@st.cache_data(show_spinner=False)
def prepare_data(file_bytes: bytes) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Separa CD e ajustes por sinal da QUANTIDADE antes de agregar por PLU.

    CD negativo (quebra)   <-> Ajustes Lojas positivo (sobra)
    CD positivo (sobra)    <-> Ajustes Lojas negativo (quebra)

    No arquivo recebido, o valor monetário de Ajustes Lojas é sempre positivo,
    mesmo nas quantidades negativas. Por isso o valor não define o sentido.
    """
    frames = load_workbook(file_bytes)
    stock_raw = standardize_columns(find_sheet(frames, "Estoque CD"))
    required_stock = {"codigo_produto", "produto", "quantidade_estoque", "valor_liquido"}
    missing_stock = required_stock.difference(stock_raw.columns)
    if missing_stock:
        raise ValueError("Colunas ausentes em Estoque CD: " + ", ".join(sorted(missing_stock)))

    stock_raw["codigo_produto"] = stock_raw["codigo_produto"].map(normalize_code)
    stock_raw = stock_raw.dropna(subset=["codigo_produto"]).copy()
    # Preserva o CD de cada registro. O ranking geral continua agregado por PLU,
    # mas os novos rankings apresentam cada produto separadamente em cada CD.
    if "codigo_empresa" in stock_raw.columns:
        stock_raw["codigo_cd"] = stock_raw["codigo_empresa"].map(normalize_code)
    else:
        stock_raw["codigo_cd"] = pd.Series(index=stock_raw.index, dtype="object")
    if "empresa" not in stock_raw.columns:
        stock_raw["empresa"] = "CD sem descrição"
    stock_raw["empresa"] = stock_raw["empresa"].fillna("CD sem descrição").astype(str).str.strip()
    inferred_cd = stock_raw["empresa"].str.extract(r"^\s*(\d{2,4})(?=\D|$)", expand=False)
    stock_raw["codigo_cd"] = stock_raw["codigo_cd"].fillna(inferred_cd).fillna("CD não informado")
    stock_raw["qtd_original"] = numeric(stock_raw["quantidade_estoque"], stock_raw.index).fillna(0)
    stock_raw["valor_original"] = numeric(stock_raw["valor_liquido"], stock_raw.index).fillna(0)
    stock_raw["quantidade_estoque"] = stock_raw["qtd_original"].abs()
    stock_raw["valor_estoque"] = stock_raw["valor_original"].abs()
    stock_raw["tipo_estoque"] = "Sem sinal"
    stock_raw.loc[stock_raw["qtd_original"].lt(0), "tipo_estoque"] = "Quebra"
    stock_raw.loc[stock_raw["qtd_original"].gt(0), "tipo_estoque"] = "Sobra"
    divergent = (stock_raw["qtd_original"] * stock_raw["valor_original"]).lt(0)

    stock = (
        stock_raw[stock_raw["tipo_estoque"].ne("Sem sinal")]
        .groupby(["tipo_estoque", "codigo_produto"], as_index=False, sort=False)
        .agg(
            produto=("produto", first_valid),
            quantidade_estoque=("quantidade_estoque", "sum"),
            valor_estoque=("valor_estoque", "sum"),
        )
    )
    stock_by_cd = (
        stock_raw[stock_raw["tipo_estoque"].ne("Sem sinal")]
        .groupby(["codigo_cd", "tipo_estoque", "codigo_produto"], as_index=False, sort=False)
        .agg(
            nome_cd=("empresa", first_valid),
            produto=("produto", first_valid),
            quantidade_estoque=("quantidade_estoque", "sum"),
            valor_estoque=("valor_estoque", "sum"),
        )
    )
    # Os códigos presentes no CD são a referência para as quatro abas de movimentos.
    cd_codes = set(stock_raw["codigo_produto"])

    specs = {
        "Ajustes Lojas": ("Ajustes Lojas", "empresa"),
        "Transferências": ("Transferências", "cgo"),
        "Devoluções": ("Devoluções", "cgo"),
        "Inventário": ("Inventário", "cgo"),
    }
    movement_frames: list[pd.DataFrame] = []
    quality_rows: list[dict] = [{
        "Aba": "Estoque CD",
        "Linhas úteis": len(stock_raw),
        "Produtos únicos": stock_raw["codigo_produto"].nunique(),
        "Produtos do CD vinculados": stock_raw["codigo_produto"].nunique(),
        "Produtos ignorados": 0,
        "Quantidade negativa": int(stock_raw["qtd_original"].lt(0).sum()),
        "Quantidade positiva": int(stock_raw["qtd_original"].gt(0).sum()),
        "Quantidade zero/ausente": int(stock_raw["qtd_original"].eq(0).sum()),
        "Sinais divergentes qtd/valor": int(divergent.sum()),
    }]

    for category, (sheet_name, subcategory_column) in specs.items():
        raw = standardize_columns(find_sheet(frames, sheet_name))
        required = {"codigo_produto", "produto", "valor_liquido"}
        if category == "Ajustes Lojas":
            required.add("quantidade_mov")
        missing = required.difference(raw.columns)
        if missing:
            raise ValueError(f"Colunas ausentes em {sheet_name}: " + ", ".join(sorted(missing)))

        raw["codigo_produto"] = raw["codigo_produto"].map(normalize_code)
        all_codes = set(raw["codigo_produto"].dropna())
        linked = raw[raw["codigo_produto"].isin(cd_codes)].copy()
        linked["categoria"] = category
        linked["valor_movimento_assinado"] = numeric(linked.get("valor_liquido"), linked.index).fillna(0)
        linked["valor_movimento"] = linked["valor_movimento_assinado"].abs()
        linked["quantidade_movimento_assinada"] = numeric(linked.get("quantidade_mov"), linked.index)
        linked["quantidade_movimento"] = linked["quantidade_movimento_assinada"].abs()

        linked["tipo_movimento"] = "Informativa" if category != "Ajustes Lojas" else "Sem sinal"
        if category == "Ajustes Lojas":
            linked.loc[linked["quantidade_movimento_assinada"].gt(0), "tipo_movimento"] = "Sobra"
            linked.loc[linked["quantidade_movimento_assinada"].lt(0), "tipo_movimento"] = "Quebra"

        if subcategory_column in linked.columns:
            linked["subcategoria"] = linked[subcategory_column].fillna("Não informado").astype(str).str.strip()
            linked.loc[linked["subcategoria"].eq(""), "subcategoria"] = "Não informado"
        else:
            linked["subcategoria"] = "Sem detalhamento"

        movement_frames.append(linked[[
            "codigo_produto", "produto", "categoria", "subcategoria", "tipo_movimento",
            "quantidade_movimento", "quantidade_movimento_assinada",
            "valor_movimento", "valor_movimento_assinado",
        ]])
        qty = linked["quantidade_movimento_assinada"]
        quality_rows.append({
            "Aba": category,
            "Linhas úteis": len(raw),
            "Produtos únicos": len(all_codes),
            "Produtos do CD vinculados": linked["codigo_produto"].nunique(),
            "Produtos ignorados": len(all_codes.difference(cd_codes)),
            "Quantidade negativa": int(qty.lt(0).sum()) if category == "Ajustes Lojas" else None,
            "Quantidade positiva": int(qty.gt(0).sum()) if category == "Ajustes Lojas" else None,
            "Quantidade zero/ausente": int((qty.isna() | qty.eq(0)).sum()) if category == "Ajustes Lojas" else None,
            "Sinais divergentes qtd/valor": None,
        })

    movements = pd.concat(movement_frames, ignore_index=True) if movement_frames else pd.DataFrame()
    quality = pd.DataFrame(quality_rows)
    return stock, stock_by_cd, movements, quality


def brl(value: float) -> str:
    formatted = f"{float(value):,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")
    return f"R$ {formatted}"


def integer_br(value: float) -> str:
    return f"{float(value):,.0f}".replace(",", ".")


def percentage(value: float) -> str:
    return f"{value * 100:.1f}%".replace(".", ",")


def metric_card(label: str, value: str, detail: str = "") -> None:
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{html.escape(label)}</div>
            <div class="metric-value">{html.escape(value)}</div>
            <div class="metric-detail">{html.escape(detail)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def side_adjustments(movements: pd.DataFrame, nature: str) -> pd.DataFrame:
    """Seleciona exclusivamente a direção oposta do CD para conciliar."""
    counterpart = "Sobra" if nature == "Quebra" else "Quebra"
    return movements.loc[
        movements["categoria"].eq("Ajustes Lojas")
        & movements["tipo_movimento"].eq(counterpart)
    ].copy()


def build_ranking(stock: pd.DataFrame, movements: pd.DataFrame, nature: str) -> pd.DataFrame:
    """Calcula conciliação sem compensar positivos e negativos do mesmo PLU."""
    side = stock[stock["tipo_estoque"].eq(nature)].copy()
    counterpart = side_adjustments(movements, nature)
    if counterpart.empty:
        side["valor_identificado"] = 0.0
        side["quantidade_identificada"] = 0.0
        side["lojas"] = 0
    else:
        totals = (
            counterpart.groupby("codigo_produto", as_index=False)
            .agg(
                valor_identificado=("valor_movimento", "sum"),
                quantidade_identificada=("quantidade_movimento", "sum"),
                lojas=("subcategoria", "nunique"),
            )
        )
        side = side.merge(totals, on="codigo_produto", how="left")
        side[["valor_identificado", "quantidade_identificada", "lojas"]] = side[
            ["valor_identificado", "quantidade_identificada", "lojas"]
        ].fillna(0)
    # A conciliação é sempre limitada ao menor valor por produto.
    # Valores excedentes nos ajustes são exibidos separadamente, sem gerar cobertura >100%.
    side["valor_conciliado"] = side[["valor_estoque", "valor_identificado"]].min(axis=1)
    side["saldo_justificar"] = (side["valor_estoque"] - side["valor_conciliado"]).clip(lower=0)
    side["excedente_ajustes"] = (side["valor_identificado"] - side["valor_estoque"]).clip(lower=0)
    side["cobertura"] = 0.0
    nonzero = side["valor_estoque"].gt(0)
    side.loc[nonzero, "cobertura"] = (
        side.loc[nonzero, "valor_conciliado"] / side.loc[nonzero, "valor_estoque"]
    )
    side["status"] = "Parcial"
    side.loc[side["valor_identificado"].le(0), "status"] = "Sem ajustes compatíveis"
    side.loc[side["saldo_justificar"].round(2).eq(0), "status"] = "Conciliada"
    side.loc[side["excedente_ajustes"].round(2).gt(0), "status"] = "Ajustes acima do CD"
    side.loc[side["valor_estoque"].eq(0), "status"] = "CD sem valor financeiro"
    return side


def build_cd_ranking(stock_by_cd: pd.DataFrame, movements: pd.DataFrame,
                     nature: str) -> pd.DataFrame:
    """Ranking por CD + PLU, com o total das lojas como referência por PLU.

    Ajustes Lojas não traz CD de destino/origem: não há rateio entre CDs.
    Portanto o mesmo valor de loja pode constar para o mesmo PLU em vários CDs.
    """
    side = stock_by_cd[stock_by_cd["tipo_estoque"].eq(nature)].copy()
    if side.empty:
        return side.assign(
            valor_identificado=pd.Series(dtype=float),
            quantidade_identificada=pd.Series(dtype=float),
            lojas=pd.Series(dtype=int),
            valor_conciliado=pd.Series(dtype=float),
            saldo_justificar=pd.Series(dtype=float),
            cobertura=pd.Series(dtype=float),
        )

    counterpart = side_adjustments(movements, nature)
    if not counterpart.empty:
        total_lojas = counterpart.groupby("codigo_produto", as_index=False).agg(
            valor_identificado=("valor_movimento", "sum"),
            quantidade_identificada=("quantidade_movimento", "sum"),
            lojas=("subcategoria", "nunique"),
        )
        side = side.merge(total_lojas, how="left", on="codigo_produto", validate="many_to_one")
    else:
        side["valor_identificado"] = 0.0
        side["quantidade_identificada"] = 0.0
        side["lojas"] = 0

    side[["valor_identificado", "quantidade_identificada", "lojas"]] = side[
        ["valor_identificado", "quantidade_identificada", "lojas"]
    ].fillna(0)
    side["valor_conciliado"] = side[["valor_estoque", "valor_identificado"]].min(axis=1)
    side["saldo_justificar"] = (side["valor_estoque"] - side["valor_conciliado"]).clip(lower=0)
    side["cobertura"] = side["valor_conciliado"].div(
        side["valor_estoque"].where(side["valor_estoque"].ne(0))
    ).fillna(0)
    return side.sort_values(["codigo_cd", "valor_estoque", "codigo_produto"],
                            ascending=[True, False, True]).reset_index(drop=True)


def render_cd_ranking(stock_by_cd: pd.DataFrame, movements: pd.DataFrame,
                      nature: str, search: str) -> None:
    """Exibe um ranking independente para cada CD, sem misturar os dois sentidos."""
    breaking = nature == "Quebra"
    cd_label = "Quebra" if breaking else "Sobra"
    store_label = "Sobra" if breaking else "Quebra"
    st.markdown(f"### Ranking: {cd_label} no CD × {store_label} nas lojas")
    st.caption(
        f"O ranking é ordenado pelo maior valor de {cd_label.lower()} no CD. "
        f"CD com quantidade {'negativa' if breaking else 'positiva'} × "
        f"Ajustes Lojas com quantidade {'positiva' if breaking else 'negativa'}; "
        "comparação pelo Código Produto (PLU)."
    )

    ranking = build_cd_ranking(stock_by_cd, movements, nature)
    if ranking.empty:
        st.info(f"Nenhum produto com {cd_label.lower()} foi encontrado no Estoque CD.")
        return

    # CDs com ou sem registros do sentido atual ficam listados no filtro.
    centers = (
        stock_by_cd.groupby("codigo_cd", as_index=False)
        .agg(nome_cd=("nome_cd", first_valid))
        .sort_values("codigo_cd", kind="stable")
    )
    cd_names = dict(zip(centers["codigo_cd"], centers["nome_cd"]))
    options = centers["codigo_cd"].tolist()
    key_side = "quebra" if breaking else "sobra"
    controls = st.columns([1.45, 1, 1])
    with controls[0]:
        selected_cds = st.multiselect(
            "Centros de distribuição", options=options, default=options,
            format_func=lambda c: cd_names.get(c, str(c)), key=f"cd_filter_{key_side}",
        )
    with controls[1]:
        limit = st.selectbox(
            "Produtos exibidos por CD", options=[20, 50, 100, 250, 500, "Todos"],
            index=1, key=f"cd_limit_{key_side}",
        )
    with controls[2]:
        matched_only = st.toggle(
            "Somente com ajuste na loja", value=False, key=f"cd_match_{key_side}",
        )

    if not selected_cds:
        st.info("Selecione pelo menos um CD para mostrar o ranking.")
        return
    ranking = ranking[ranking["codigo_cd"].isin(selected_cds)].copy()
    if search.strip():
        token = normalize_label(search)
        ranking = ranking[
            ranking["produto"].map(normalize_label).str.contains(re.escape(token), regex=True)
            | ranking["codigo_produto"].map(normalize_label).str.contains(re.escape(token), regex=True)
        ]
    if matched_only:
        ranking = ranking[ranking["valor_identificado"].gt(0)]
    if ranking.empty:
        st.info("Não existem produtos correspondentes aos filtros selecionados.")
        return

    if len(selected_cds) > 1:
        st.info(
            "A aba Ajustes Lojas não identifica a qual CD a movimentação pertence. "
            "Por isso, o total das lojas é mostrado como referência para cada PLU e pode "
            "aparecer em mais de um CD. Não some os valores das lojas entre CDs como "
            "se fossem movimentações exclusivas."
        )
    else:
        st.caption(
            "Os ajustes das lojas são agrupados pelo PLU e não possuem vínculo explícito "
            "de loja para o CD selecionado."
        )

    summary = st.columns(4)
    for column, (title, value, detail) in zip(summary, [
        (f"Total {cd_label} no CD", brl(ranking["valor_estoque"].sum()), "CDs selecionados"),
        (f"{store_label} nas lojas (referência)", brl(ranking["valor_identificado"].sum()),
         "Por PLU × CD; não exclusivo se vários CDs"),
        ("Valor comparado por CD", brl(ranking["valor_conciliado"].sum()),
         "Mínimo por PLU e CD; não é rateio"),
        ("Saldo não coberto no CD", brl(ranking["saldo_justificar"].sum()),
         "Valor CD menos o comparado"),
    ]):
        with column:
            metric_card(title, value, detail)

    display_cols = {
        "codigo_cd": "CD", "nome_cd": "Descrição CD", "codigo_produto": "Código Produto",
        "produto": "Produto", "quantidade_estoque": f"Quantidade {cd_label} CD",
        "valor_estoque": f"Valor {cd_label} no CD",
        "quantidade_identificada": f"Quantidade {store_label} Lojas",
        "valor_identificado": f"Valor {store_label} na Loja",
        "cobertura": "Cobertura",
    }
    configs = {
        "#": st.column_config.NumberColumn(width="small", format="%d"),
        "CD": st.column_config.TextColumn(width="small"),
        "Descrição CD": st.column_config.TextColumn(width="medium"),
        "Código Produto": st.column_config.TextColumn(width="small"),
        "Produto": st.column_config.TextColumn(width="large"),
        f"Quantidade {cd_label} CD": st.column_config.NumberColumn(format="%,.2f"),
        f"Quantidade {store_label} Lojas": st.column_config.NumberColumn(format="%,.2f"),
        f"Valor {cd_label} no CD": st.column_config.NumberColumn(format="R$ %,.2f"),
        f"Valor {store_label} na Loja": st.column_config.NumberColumn(format="R$ %,.2f"),
        "Cobertura": st.column_config.NumberColumn(format="percent"),
    }
    st.write()
    # Exporta TODOS os registros filtrados, mesmo que as tabelas mostrem Top N.
    full_export = ranking.rename(columns=display_cols)
    st.download_button(
        f"Baixar ranking completo — {cd_label} CD × {store_label} Lojas (CSV)",
        data=full_export.to_csv(index=False, sep=";", decimal=",", encoding="utf-8-sig").encode("utf-8-sig"),
        file_name=f"ranking_por_cd_{key_side}_cd.csv", mime="text/csv",
        width="stretch", key=f"download_cd_{key_side}",
    )

    for cd in selected_cds:
        cd_rows = ranking[ranking["codigo_cd"].eq(cd)]
        if cd_rows.empty:
            continue
        st.markdown(f"#### {html.escape(str(cd_names.get(cd, cd)))}")
        st.caption(
            f"{integer_br(len(cd_rows))} produtos | "
            f"{cd_label} CD: {brl(cd_rows['valor_estoque'].sum())} | "
            f"{store_label} Lojas: {brl(cd_rows['valor_identificado'].sum())} "
            "(referência por PLU)"
        )
        shown = cd_rows if limit == "Todos" else cd_rows.head(limit)
        shown = shown.reset_index(drop=True).copy()
        shown.insert(0, "#", range(1, len(shown) + 1))
        shown = shown.rename(columns=display_cols)
        st.dataframe(
            shown[["#", "Código Produto", "Produto", f"Quantidade {cd_label} CD",
                   f"Valor {cd_label} no CD", f"Quantidade {store_label} Lojas",
                   f"Valor {store_label} na Loja", "Cobertura"]],
            hide_index=True, width="stretch", height=min(720, 40 + 35 * max(len(shown), 1)),
            column_config=configs,
        )


def inject_css() -> None:
    st.markdown(
        """
        <style>
        :root {
            color-scheme: dark;
            --accent: #ff4b55;
            --app-bg: #0d1117;
            --sidebar-bg: #232630;
            --panel: #191d24;
            --panel-secondary: #14181e;
            --input-bg: #11151b;
            --text: #f3f4f6;
            --text-strong: #f8fafc;
            --muted: #aeb5c0;
            --muted-soft: #8f98a5;
            --border: #2c323d;
            --button-border: #3a414d;
            --shadow: rgba(0, 0, 0, .18);
            --hover: rgba(255, 75, 85, .055);
        }
        @media (prefers-color-scheme: light) {
            :root {
                color-scheme: light;
                --app-bg: #f6f7fb;
                --sidebar-bg: #eef0f5;
                --panel: #ffffff;
                --panel-secondary: #f8f9fc;
                --input-bg: #ffffff;
                --text: #1b1f27;
                --text-strong: #111827;
                --muted: #5f6877;
                --muted-soft: #70798a;
                --border: #d8dde6;
                --button-border: #c5cbd6;
                --shadow: rgba(25, 35, 55, .08);
                --hover: rgba(255, 75, 85, .07);
            }
        }
        html, body, [data-testid="stAppViewContainer"], .stApp {
            background: var(--app-bg) !important;
            color: var(--text) !important;
        }
        [data-testid="stHeader"] {
            background: color-mix(in srgb, var(--app-bg) 92%, transparent) !important;
            color: var(--text) !important;
        }
        [data-testid="stToolbar"], [data-testid="stDecoration"] {
            color: var(--text) !important;
        }
        [data-testid="stSidebar"] {
            background: var(--sidebar-bg) !important;
            color: var(--text) !important;
            border-right: 1px solid var(--border);
        }
        [data-testid="stSidebar"] > div { padding-top: 1.1rem; }
        .stApp p, .stApp label, .stApp h1, .stApp h2, .stApp h3,
        .stApp h4, .stApp h5, .stApp h6, [data-testid="stSidebar"] p,
        [data-testid="stSidebar"] label { color: var(--text); }
        .block-container {
            max-width: none !important;
            padding: 1.35rem 2rem 3rem !important;
        }
        h1, h2, h3 { letter-spacing: -0.025em; }
        .eyebrow {
            display: block; min-height: 1.8rem; padding: .25rem 0 .45rem;
            overflow: visible; line-height: 1.4; color: var(--accent);
            font-weight: 700; text-transform: uppercase; letter-spacing: .12em; font-size: .72rem;
        }
        .hero-subtitle { color: var(--muted); margin-top: -.55rem; margin-bottom: 1.35rem; }
        .metric-card {
            min-width: 0; min-height: 132px; padding: 18px 18px 16px; overflow: hidden;
            border: 1px solid var(--border);
            border-radius: 12px; background: linear-gradient(145deg, var(--panel) 0%, var(--panel-secondary) 100%);
            box-shadow: 0 10px 26px var(--shadow);
        }
        .metric-label { color: var(--muted); font-size: .78rem; font-weight: 650; min-height: 20px; }
        .metric-value {
            color: var(--text-strong); font-size: clamp(.92rem, 1.28vw, 1.45rem); line-height: 1.15;
            letter-spacing: -.035em; font-weight: 720; font-variant-numeric: tabular-nums;
            margin: 9px 0 6px; white-space: nowrap;
        }
        .metric-detail { color: var(--muted-soft); font-size: .72rem; }
        [data-testid="stDataFrame"] {
            border: 1px solid var(--border); border-radius: 10px; overflow: hidden;
            background: var(--panel);
        }
        [data-testid="stExpander"] {
            border: 1px solid var(--border); border-radius: 10px;
            background: linear-gradient(145deg, var(--panel) 0%, var(--panel-secondary) 100%);
            overflow: hidden;
        }
        [data-testid="stExpander"] summary { padding: .72rem .9rem; }
        [data-testid="stExpander"] summary:hover { background: var(--hover); }
        [data-baseweb="tab-list"] { gap: 1.1rem; border-bottom: 1px solid var(--border); }
        [data-baseweb="tab"] { padding-left: 0; padding-right: 0; }
        [data-baseweb="tab"][aria-selected="true"] { color: var(--accent) !important; }
        [data-baseweb="input"] > div, [data-baseweb="select"] > div,
        [data-testid="stFileUploaderDropzone"], textarea, input {
            background: var(--input-bg) !important;
            color: var(--text) !important;
            border-color: var(--border) !important;
        }
        [data-baseweb="popover"], [role="listbox"] {
            background: var(--panel) !important;
            color: var(--text) !important;
        }
        .stButton > button, .stDownloadButton > button {
            color: var(--text); background: var(--panel); border-color: var(--button-border); border-radius: 8px;
        }
        .stButton > button:hover, .stDownloadButton > button:hover { border-color: var(--accent); color: var(--accent); }
        [data-testid="stFileUploaderDropzone"] { background: var(--input-bg); border-color: var(--border); }
        .detail-title { border-left: 3px solid var(--accent); padding-left: 12px; margin: 1.2rem 0 .8rem; }
        .detail-code { color: var(--muted-soft); }
        .method-note { color: var(--muted-soft); font-size: .78rem; line-height: 1.45; }
        hr { border-color: var(--border) !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def detail_table(movements: pd.DataFrame, category: str) -> pd.DataFrame:
    category_data = movements[movements["categoria"].eq(category)]
    if category_data.empty:
        return pd.DataFrame()
    grouped = (
        category_data.groupby("subcategoria", as_index=False)
        .agg(
            quantidade=("quantidade_movimento_assinada", lambda values: values.sum(min_count=1)),
            valor=("valor_movimento", "sum") if category == "Ajustes Lojas" else ("valor_movimento_assinado", "sum"),
            registros=("codigo_produto", "size"),
        )
        .sort_values("valor", ascending=False)
    )
    grouped["Quantidade"] = grouped["quantidade"].map(lambda value: "—" if pd.isna(value) else integer_br(value))
    grouped["Valor vinculado"] = grouped["valor"].map(brl)
    return grouped.rename(columns={"subcategoria": "Detalhamento", "registros": "Registros"})[
        ["Detalhamento", "Quantidade", "Valor vinculado", "Registros"]
    ]


def render_store_ranking(ranking: pd.DataFrame, movements: pd.DataFrame, nature: str,
                         search: str, top_n: int) -> None:
    adjustments = side_adjustments(movements, nature)
    counterpart = "sobras" if nature == "Quebra" else "quebras"
    if adjustments.empty:
        st.warning(f"Não há {counterpart} em Ajustes Lojas para os produtos do CD.")
        return

    product_ranking = ranking.copy()
    if search.strip():
        token = normalize_label(search)
        product_ranking = product_ranking[
            product_ranking["produto"].map(normalize_label).str.contains(re.escape(token), regex=True)
            | product_ranking["codigo_produto"].map(normalize_label).str.contains(re.escape(token), regex=True)
        ]

    controls = st.columns([1.35, 1])
    with controls[0]:
        store_order = st.selectbox(
            "Ordenar produtos por",
            ["Maior valor no Estoque CD", "Maior valor em Ajustes Lojas", "Maior quantidade em estoque"],
            key=f"store_order_{nature}",
        )
    with controls[1]:
        show_without_adjustments = st.toggle(
            "Exibir produtos sem ajustes compatíveis", value=False, key=f"show_without_adjustments_{nature}"
        )
    if not show_without_adjustments:
        product_ranking = product_ranking[product_ranking["valor_identificado"].gt(0)]

    order_columns = {
        "Maior valor no Estoque CD": "valor_estoque",
        "Maior valor em Ajustes Lojas": "valor_identificado",
        "Maior quantidade em estoque": "quantidade_estoque",
    }
    product_ranking = product_ranking.sort_values(
        [order_columns[store_order], "valor_identificado"], ascending=False
    ).head(top_n).reset_index(drop=True)

    associated = adjustments[adjustments["codigo_produto"].isin(ranking["codigo_produto"])].copy()
    summary = st.columns(3)
    with summary[0]:
        metric_card(f"Total Ajustes Lojas ({counterpart})", brl(associated["valor_movimento"].sum()),
                    "Somente produtos deste grupo no CD")
    with summary[1]:
        metric_card("Produtos com ajustes", integer_br(associated["codigo_produto"].nunique()),
                    "Produtos com direção compatível")
    with summary[2]:
        metric_card("Lojas identificadas", integer_br(associated["subcategoria"].nunique()),
                    "Lojas com ajustes compatíveis")

    st.markdown("#### Produtos e ajustes por loja")
    st.caption(f"Detalhamento exclusivo de {counterpart} em Ajustes Lojas para {nature.lower()}s do CD.")
    if product_ranking.empty:
        st.info("Nenhum produto corresponde aos filtros atuais.")
        return

    for position, product in product_ranking.iterrows():
        label = (
            f'**{position + 1}. {product["produto"]}**  ·  Código {product["codigo_produto"]}'
            f'  ·  CD: **{brl(product["valor_estoque"])}**'
            f'  ·  Lojas: **{brl(product["valor_identificado"])}**'
        )
        with st.expander(label):
            product_stores = associated[associated["codigo_produto"].eq(product["codigo_produto"])]
            if product_stores.empty:
                st.info("Este produto não possui ajustes de sentido compatível.")
                continue
            store_table = (
                product_stores.groupby("subcategoria", as_index=False)
                .agg(
                    quantidade_ajustada=("quantidade_movimento", "sum"),
                    valor_ajustes=("valor_movimento", "sum"),
                    registros=("codigo_produto", "size"),
                )
                .sort_values("valor_ajustes", ascending=False).reset_index(drop=True)
            )
            store_table.insert(0, "ranking", range(1, len(store_table) + 1))
            total = store_table["valor_ajustes"].sum()
            store_table["participacao"] = store_table["valor_ajustes"] / total if total else 0
            store_table = store_table.rename(columns={
                "ranking": "#", "subcategoria": "Loja", "quantidade_ajustada": "Quantidade (absoluta)",
                "valor_ajustes": "Valor Ajustes Lojas", "participacao": "% dos ajustes", "registros": "Registros",
            })
            st.dataframe(
                store_table, hide_index=True, width="stretch",
                height=min(430, 39 + 35 * max(len(store_table), 1)),
                column_config={
                    "#": st.column_config.NumberColumn(width="small", format="%d"),
                    "Loja": st.column_config.TextColumn(width="large"),
                    "Quantidade (absoluta)": st.column_config.NumberColumn(format="%,.2f"),
                    "Valor Ajustes Lojas": st.column_config.NumberColumn(format="R$ %,.2f"),
                    "% dos ajustes": st.column_config.NumberColumn(format="percent"),
                    "Registros": st.column_config.NumberColumn(format="%d"),
                },
            )


def render_informative_movements(stock: pd.DataFrame, movements: pd.DataFrame,
                                 search: str, top_n: int) -> None:
    informative = movements[movements["categoria"].ne("Ajustes Lojas")].copy()
    st.markdown("#### Movimentações informativas")
    st.caption("Transferências, devoluções e inventário não participam da conciliação de quebras e sobras.")

    cards: list[tuple[str, float, str]] = []
    for category in ["Transferências", "Devoluções"]:
        subset = informative[informative["categoria"].eq(category)]
        cards.append((f"Total {category}", float(subset["valor_movimento_assinado"].sum()),
                      f'{subset["codigo_produto"].nunique()} produtos vinculados'))
    inventory = informative[informative["categoria"].eq("Inventário")]
    for cgo, subset in inventory.groupby("subcategoria", sort=True):
        cards.append((str(cgo).title(), float(subset["valor_movimento_assinado"].sum()),
                      f'{subset["codigo_produto"].nunique()} produtos vinculados'))
    if cards:
        for column, (label, value, detail) in zip(st.columns(len(cards)), cards):
            with column:
                metric_card(label, brl(value), detail)

    values_by_side = stock.pivot_table(
        index="codigo_produto", columns="tipo_estoque", values="valor_estoque",
        aggfunc="sum", fill_value=0,
    ).rename(columns={"Quebra": "Quebra CD", "Sobra": "Sobra CD"}).reset_index()
    for column in ["Quebra CD", "Sobra CD"]:
        if column not in values_by_side.columns:
            values_by_side[column] = 0.0
    products = stock.groupby("codigo_produto", as_index=False).agg(produto=("produto", first_valid))
    comparison = products.merge(values_by_side, on="codigo_produto", how="left")
    movement_columns: list[str] = []

    for category in ["Transferências", "Devoluções"]:
        total = (
            informative[informative["categoria"].eq(category)]
            .groupby("codigo_produto")["valor_movimento_assinado"].sum().rename(category)
        )
        comparison = comparison.merge(total, on="codigo_produto", how="left")
        movement_columns.append(category)
    if not inventory.empty:
        pivot = inventory.pivot_table(
            index="codigo_produto", columns="subcategoria", values="valor_movimento_assinado",
            aggfunc="sum", fill_value=0,
        )
        renamed = {}
        for column in pivot.columns:
            normalized = normalize_label(column)
            if "entrada" in normalized:
                renamed[column] = f"Inventário - Entrada ({column})"
            elif "saida" in normalized:
                renamed[column] = f"Inventário - Saída ({column})"
            else:
                renamed[column] = f"Inventário - {column}"
        pivot = pivot.rename(columns=renamed).reset_index()
        comparison = comparison.merge(pivot, on="codigo_produto", how="left")
        movement_columns.extend(renamed.values())

    comparison[movement_columns] = comparison[movement_columns].fillna(0)
    comparison["Total informativo"] = comparison[movement_columns].sum(axis=1)
    if search.strip():
        token = normalize_label(search)
        comparison = comparison[
            comparison["produto"].map(normalize_label).str.contains(re.escape(token), regex=True)
            | comparison["codigo_produto"].map(normalize_label).str.contains(re.escape(token), regex=True)
        ]
    comparison = comparison.assign(_ordem=comparison["Total informativo"].abs())
    comparison = comparison.sort_values("_ordem", ascending=False).drop(columns="_ordem").head(top_n)
    comparison.insert(0, "ranking", range(1, len(comparison) + 1))
    comparison = comparison.rename(columns={
        "ranking": "#", "codigo_produto": "Código", "produto": "Produto",
    })
    st.markdown("#### Movimentações por produto")
    column_config = {
        "#": st.column_config.NumberColumn(width="small", format="%d"),
        "Código": st.column_config.TextColumn(width="small"),
        "Produto": st.column_config.TextColumn(width="large"),
    }
    for column in ["Quebra CD", "Sobra CD", "Total informativo", *movement_columns]:
        column_config[column] = st.column_config.NumberColumn(format="R$ %,.2f")
    st.dataframe(comparison, hide_index=True, width="stretch",
                 height=min(720, 39 + 35 * max(len(comparison), 1)), column_config=column_config)


def render_product_detail(selected: pd.Series, product_movements: pd.DataFrame, nature: str) -> None:
    label = "Quebra" if nature == "Quebra" else "Sobra"
    counterpart = "Sobras" if nature == "Quebra" else "Quebras"
    st.markdown(
        f'<div class="detail-title"><h3 style="margin:0">{html.escape(str(selected["produto"]))}</h3>'
        f'<span class="detail-code">Código {html.escape(str(selected["codigo_produto"]))}</span></div>',
        unsafe_allow_html=True,
    )
    for column, (title, value, detail) in zip(st.columns(4), [
        ("Quantidade no CD", f'{selected["quantidade_estoque"]:,.2f}', "Magnitude da quantidade"),
        (f"Valor de {label.lower()} no CD", brl(selected["valor_estoque"]), "Valor absoluto, sem compensação"),
        (f"{counterpart} em Ajustes Lojas", brl(selected["valor_identificado"]), "Somente o sentido oposto"),
        ("Cobertura conciliada", percentage(selected["cobertura"]), selected["status"]),
    ]):
        with column:
            metric_card(title, value, detail)

    if product_movements.empty:
        st.warning("Nenhuma movimentação selecionada foi encontrada para este produto.")
        return
    summary = (
        product_movements.groupby("categoria", as_index=False)
        .agg(
            quantidade=("quantidade_movimento_assinada", lambda values: values.sum(min_count=1)),
            valor=("valor_movimento", "sum"),
            registros=("codigo_produto", "size"),
        )
    )
    summary["Quantidade assinada"] = summary["quantidade"].map(
        lambda value: "—" if pd.isna(value) else f"{value:,.2f}"
    )
    summary["Valor vinculado"] = summary["valor"].map(brl)
    summary["% do CD"] = summary["valor"].map(
        lambda value: percentage(value / selected["valor_estoque"]) if selected["valor_estoque"] else "—"
    )
    summary["Participa da conciliação"] = summary["categoria"].map(
        lambda category: "Sim" if category == "Ajustes Lojas" else "Não"
    )
    st.markdown("#### Movimentações do produto")
    st.caption(f"Apenas {counterpart.lower()} de Ajustes Lojas participam do confronto com {label.lower()}s do CD.")
    st.dataframe(
        summary.rename(columns={"categoria": "Categoria", "registros": "Registros"})[
            ["Categoria", "Quantidade assinada", "Valor vinculado", "% do CD", "Participa da conciliação", "Registros"]
        ], hide_index=True, width="stretch",
    )
    available_categories = [category for category in CATEGORIES if category in set(product_movements["categoria"])]
    for tab, category in zip(st.tabs(available_categories), available_categories):
        with tab:
            detail = detail_table(product_movements, category)
            if category == "Ajustes Lojas":
                st.caption(f"Lojas com {counterpart.lower()} (quantidade de sentido oposto ao CD).")
            else:
                st.caption("Movimentação informativa; não participa da conciliação.")
            st.dataframe(detail, hide_index=True, width="stretch", height=min(420, 72 + 36 * len(detail)))


def render_side_panel(stock: pd.DataFrame, movements: pd.DataFrame, nature: str,
                      search: str, top_n: int, order_label: str,
                      detail_categories: list[str]) -> pd.DataFrame:
    ranking = build_ranking(stock, movements, nature)
    counterpart = "Sobras" if nature == "Quebra" else "Quebras"
    plural = "Quebras" if nature == "Quebra" else "Sobras"
    direction = "negativa" if nature == "Quebra" else "positiva"
    st.markdown(f"### {plural} do Estoque CD × {counterpart} em Ajustes Lojas")
    st.caption(
        f"CD: quantidade {direction}; Ajustes Lojas: quantidade "
        f'{"positiva" if nature == "Quebra" else "negativa"}. '
        "Confronto por Código Produto (PLU), sem misturar sentidos nem compensar registros de sinais diferentes."
    )
    if ranking.empty:
        st.info(f"Nenhuma {nature.lower()} foi encontrada no Estoque CD.")
        return ranking

    total_stock = ranking["valor_estoque"].sum()
    total_adj = ranking["valor_identificado"].sum()
    total_matched = ranking["valor_conciliado"].sum()
    unmatched = ranking["saldo_justificar"].sum()
    coverage = total_matched / total_stock if total_stock else 0
    for column, (title, value, detail) in zip(st.columns(5), [
        (f"Total {plural} CD", brl(total_stock), f'{ranking["codigo_produto"].nunique()} produtos'),
        (f"{counterpart} em Ajustes Lojas", brl(total_adj), "Somente PLUs desta categoria"),
        ("Valor conciliado", brl(total_matched), "Limitado ao valor do CD por PLU"),
        ("Saldo a justificar", brl(unmatched), "Diferença não conciliada por PLU"),
        ("Cobertura", percentage(coverage), "Conciliado / Estoque CD"),
    ]):
        with column:
            metric_card(title, value, detail)

    comparison_tab, stores_tab = st.tabs(["Ranking e justificativas", "Ajustes por loja"])
    with comparison_tab:
        filtered = ranking.copy()
        if search.strip():
            token = normalize_label(search)
            filtered = filtered[
                filtered["produto"].map(normalize_label).str.contains(re.escape(token), regex=True)
                | filtered["codigo_produto"].map(normalize_label).str.contains(re.escape(token), regex=True)
            ]
        order_col = "quantidade_estoque" if order_label == "Maior quantidade" else "valor_estoque"
        filtered = filtered.sort_values([order_col, "valor_estoque"], ascending=False).head(top_n).reset_index(drop=True)
        filtered.insert(0, "ranking", range(1, len(filtered) + 1))
        display = filtered[[
            "ranking", "codigo_produto", "produto", "quantidade_estoque", "valor_estoque",
            "valor_identificado", "valor_conciliado", "saldo_justificar", "excedente_ajustes", "cobertura", "status",
        ]].rename(columns={
            "ranking": "#", "codigo_produto": "Código", "produto": "Produto",
            "quantidade_estoque": "Quantidade CD", "valor_estoque": f"Valor {nature} CD",
            "valor_identificado": f"Ajustes Lojas ({counterpart})", "valor_conciliado": "Valor conciliado",
            "saldo_justificar": "Saldo a justificar", "excedente_ajustes": "Excedente ajustes",
            "cobertura": "Cobertura", "status": "Status",
        })
        st.caption(
            "Selecione um produto para detalhar. A conciliação financeira é limitada ao menor valor do par CD × lojas "
            "para cada PLU; diferenças de custo aparecem como saldo ou excedente."
        )
        config = {
            "#": st.column_config.NumberColumn(width="small", format="%d"),
            "Código": st.column_config.TextColumn(width="small"),
            "Produto": st.column_config.TextColumn(width="large"),
            "Quantidade CD": st.column_config.NumberColumn(format="%,.2f"),
            "Cobertura": st.column_config.NumberColumn(format="percent"),
        }
        for name in [f"Valor {nature} CD", f"Ajustes Lojas ({counterpart})", "Valor conciliado",
                     "Saldo a justificar", "Excedente ajustes"]:
            config[name] = st.column_config.NumberColumn(format="R$ %,.2f")
        selection = st.dataframe(
            display, hide_index=True, width="stretch",
            height=min(720, 39 + 35 * max(len(display), 1)),
            on_select="rerun", selection_mode="single-row", key=f"ranking_selection_{nature}",
            column_config=config,
        )
        st.download_button(
            "Baixar comparativo filtrado (CSV)",
            data=display.to_csv(index=False, sep=";", decimal=",", encoding="utf-8-sig").encode("utf-8-sig"),
            file_name=f"comparativo_{'quebras' if nature == 'Quebra' else 'sobras'}_cd_ajustes_lojas.csv",
            mime="text/csv", width="stretch", key=f"download_{nature}",
        )
        selected_rows = selection.selection.rows if selection is not None else []
        if selected_rows and not filtered.empty:
            selected = filtered.iloc[selected_rows[0]]
        elif not filtered.empty:
            code = st.selectbox(
                "Ou escolha um produto para detalhar", options=filtered["codigo_produto"].tolist(),
                format_func=lambda code: f'{code} · {filtered.loc[filtered["codigo_produto"].eq(code), "produto"].iloc[0]}',
                key=f"select_product_{nature}",
            )
            selected = filtered[filtered["codigo_produto"].eq(code)].iloc[0]
        else:
            selected = None
        if selected is not None:
            selection_movements = movements[movements["categoria"].isin(detail_categories)].copy()
            selection_movements = selection_movements[
                selection_movements["categoria"].ne("Ajustes Lojas")
                | selection_movements["tipo_movimento"].eq("Sobra" if nature == "Quebra" else "Quebra")
            ]
            product_movements = selection_movements[
                selection_movements["codigo_produto"].eq(selected["codigo_produto"])
            ]
            render_product_detail(selected, product_movements, nature)

    with stores_tab:
        render_store_ranking(ranking, movements, nature, search, top_n)
    return ranking


def main() -> None:
    st.set_page_config(page_title="Análise de Quebras e Sobras CD", page_icon="📦", layout="wide")
    inject_css()
    with st.sidebar:
        st.markdown("### Bases de dados")
        st.caption("Envie uma nova planilha ou utilize a base incluída no projeto.")
        uploaded = st.file_uploader("Planilha de análise", type=["xlsx"],
                                    help="Planilha Excel com cinco abas: Estoque CD, Ajustes Lojas, Transferências, Devoluções e Inventário.")
        default_file = next((path for path in DEFAULT_FILES if path.exists()), None)
        if uploaded is not None:
            file_bytes = uploaded.getvalue()
            source_label = uploaded.name
        elif default_file is not None:
            file_bytes = default_file.read_bytes()
            source_label = default_file.name
        else:
            st.error("Inclua a planilha em data/Analise CD GERAL.xlsx ou faça o upload acima.")
            st.stop()
        st.success(f"Base ativa: {source_label}")
        st.divider()
        st.markdown("### Filtros")
        search = st.text_input("Produto ou código", placeholder="Ex.: 2203 ou Itaipava")
        detail_categories = st.multiselect(
            "Movimentações no detalhamento", CATEGORIES, default=CATEGORIES,
            help="A seleção afeta apenas o detalhamento. A conciliação utiliza sempre Ajustes Lojas do sentido oposto.",
        )
        order_label = st.radio("Ordenar ranking por", ["Maior quantidade", "Maior valor"])
        top_n = st.slider("Produtos exibidos", 10, 130, 30, step=10)

    try:
        with st.spinner("Lendo e conciliando os sentidos das cinco abas..."):
            stock, stock_by_cd, movements_all, quality = prepare_data(file_bytes)
    except Exception as error:
        st.error(f"Não foi possível processar a planilha: {error}")
        st.stop()

    st.markdown('<div class="eyebrow">Controle de estoque • Quebras e sobras</div>', unsafe_allow_html=True)
    st.title("Análise de Quebras e Sobras — Estoque CD")
    st.markdown(
        '<div class="hero-subtitle">Conciliação cruzada por PLU: quebra CD × sobra loja; sobra CD × quebra loja.</div>',
        unsafe_allow_html=True,
    )
    st.caption("O sentido é determinado pela QUANTIDADE, não pelo valor líquido. Os valores financeiros são exibidos em magnitude absoluta após a classificação.")
    cd_break_tab, cd_surplus_tab, break_tab, surplus_tab, quality_tab = st.tabs([
        "Ranking Quebra CD × Sobra Loja", "Ranking Sobra CD × Quebra Loja",
        "Quebras (CD negativo)", "Sobras (CD positivo)",
        "Qualidade dos dados",
    ])
    with cd_break_tab:
        render_cd_ranking(stock_by_cd, movements_all, "Quebra", search)
    with cd_surplus_tab:
        render_cd_ranking(stock_by_cd, movements_all, "Sobra", search)
    with break_tab:
        breaks = render_side_panel(stock, movements_all, "Quebra", search, top_n, order_label, detail_categories)
    with surplus_tab:
        surpluses = render_side_panel(stock, movements_all, "Sobra", search, top_n, order_label, detail_categories)
    with quality_tab:
        st.markdown("#### Qualidade e cobertura das abas")
        st.caption("Quantidade zero/ausente não é classificada como sobra ou quebra e permanece fora da conciliação.")
        st.dataframe(quality, hide_index=True, width="stretch")
        cd_conflicts = quality.loc[quality["Aba"].eq("Estoque CD"), "Sinais divergentes qtd/valor"].iloc[0]
        st.info(
            f"O CD contém {int(cd_conflicts)} linhas em que quantidade e valor financeiro têm sinais opostos. "
            "Essas linhas foram classificadas pela quantidade, com valor financeiro absoluto. "
            "Valide as divergências na origem quando necessário."
        )
        for nature, ranking in [("quebras", breaks), ("sobras", surpluses)]:
            st.markdown(f"#### Produtos com {nature} sem ajustes compatíveis")
            empty_match = ranking[ranking["valor_identificado"].eq(0)][
                ["codigo_produto", "produto", "valor_estoque", "quantidade_estoque"]
            ].sort_values("valor_estoque", ascending=False).head(200)
            if empty_match.empty:
                st.success(f"Todos os produtos com {nature} possuem ajustes de sinal oposto.")
            else:
                st.caption(f"{len(ranking[ranking['valor_identificado'].eq(0)])} produtos sem correspondência; exibindo até 200.")
                st.dataframe(empty_match.rename(columns={
                    "codigo_produto": "Código", "produto": "Produto", "valor_estoque": "Valor CD",
                    "quantidade_estoque": "Quantidade CD",
                }), hide_index=True, width="stretch")
        st.markdown(
            '<p class="method-note"><strong>Metodologia:</strong> apenas Ajustes Lojas participa da conciliação. '
            'O cruzamento ocorre pelo Código Produto, comparando quantidade negativa do CD com quantidade positiva das lojas, '
            'e vice-versa. As abas por CD exibem o total das lojas por PLU como referência, sem rateio entre CDs. '
            'Não há compensação entre sentidos; Transferências, Devoluções e Inventário são informativos. '
            'Valor conciliado = mínimo(CD, Ajustes Lojas) em cada PLU.</p>', unsafe_allow_html=True,
        )


if __name__ == "__main__":
    main()
