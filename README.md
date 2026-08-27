# Dashboard de Análise de Quebra — Estoque CD

Dashboard em Streamlit que usa a aba **Estoque CD** como universo de referência e cruza seus produtos com **Ajustes Lojas**, **Transferências**, **Devoluções** e **Inventário**.

## O que o sistema entrega

- KPIs de valor total da quebra, quebra identificada, saldo a justificar, cobertura e quantidade em estoque.
- Ranking padrão pela maior quantidade em estoque, com o valor da quebra ao lado.
- Seleção de uma linha do ranking para abrir as categorias que justificam o produto.
- Detalhamento de Ajustes Lojas por loja.
- Detalhamento de Inventário por CGO.
- Filtros por produto/código, categoria, status e quantidade de itens exibidos.
- Exclusão automática de produtos que aparecem nas outras abas, mas não existem no Estoque CD.
- Aba de qualidade dos dados e download do ranking filtrado em CSV.

## Regra da quebra identificada

A quebra identificada é calculada por produto como:

`soma absoluta dos valores vinculados em Ajustes Lojas, Transferências, Devoluções e Inventário`

A cobertura pode ultrapassar 100%. O único filtro de inclusão é o código do produto também existir na aba Estoque CD; produtos exclusivos das demais abas são ignorados.

## Como executar

1. Instale o Python 3.11 ou superior.
2. Abra o terminal na pasta do projeto.
3. Crie e ative um ambiente virtual (recomendado).
4. Instale as dependências:

   ```bash
   pip install -r requirements.txt
   ```

5. Inicie o dashboard:

   ```bash
   streamlit run app.py
   ```

O navegador abrirá normalmente em `http://localhost:8501`.

## Atualização da base

O projeto já inclui a planilha analisada na pasta `data`. Para analisar uma versão nova, use o campo de upload na barra lateral. A nova planilha deve manter as cinco abas e os cabeçalhos usados na base original.
