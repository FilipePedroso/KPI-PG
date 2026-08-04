# Dados

Coloque aqui os dois arquivos Excel:

- `Dados.xlsx` — abas `f_venda_total` e `f_ec_oniz`
- `Estrutura.xlsx` — abas `d_comercial`, `d_metas`, `d_clientes_braveo`

## Atualização automática

Basta **substituir o arquivo** e fazer o commit/push (ou publicar pelo Lovable).
O GitHub Actions (`.github/workflows/build-data.yml`) roda o
`scripts/build_data.py` sozinho e commita os arquivos gerados
(`docs/data.json`, `public/data.json`, `sem-estrutura.csv`).

Em ~1 min o GitHub Pages já mostra os dados novos. Nenhum comando manual.
