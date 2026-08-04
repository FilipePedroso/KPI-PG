# Dados

Coloque aqui os dois arquivos Excel:

- `Dados.xlsx` — abas `f_venda_total` e `f_ec_oniz`
- `Estrutura.xlsx` — abas `d_comercial`, `d_metas`, `d_clientes_braveo`

## Como atualizar

1. Substitua o `.xlsx` direto no GitHub (upload/commit ou push).
2. Pronto.

O workflow `.github/workflows/build-data.yml` dispara a cada push que altere
`data/`, roda `scripts/build_data.py` usando os arquivos do repositório e
commita `docs/data.json`, `public/data.json` e os `sem-estrutura.csv`.
Em ~1 min o GitHub Pages já mostra os dados novos — nenhum comando manual.

Obs.: no ambiente do Lovable o editor não aceita arquivos acima de 10 MB, por
isso os `.xlsx` grandes só entram pelo GitHub. O script mantém compatibilidade
com os `*.asset.json` (CDN) quando o `.xlsx` não está presente.
