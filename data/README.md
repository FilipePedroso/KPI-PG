# Dados

Os dois arquivos Excel ficam **fora do repositório** (limite de 10 MB por arquivo)
e são hospedados no CDN via `*.asset.json`:

- `Dados.xlsx` — abas `f_venda_total` e `f_ec_oniz`
- `Estrutura.xlsx` — abas `d_comercial`, `d_metas`, `d_clientes_braveo`

## Como atualizar (automático)

1. Solte o arquivo novo aqui em `data/` pelo Lovable.
2. Publique / faça deploy pelo Lovable.

O `prebuild` roda `scripts/build_data.py`, que:
- detecta que o arquivo local mudou, sobe o asset novo e atualiza o `*.asset.json`;
- regenera `docs/data.json`, `public/data.json` e os `sem-estrutura.csv`;
- esses arquivos gerados são commitados e sincronizados com o GitHub.

O GitHub Actions (`.github/workflows/build-data.yml`) também regenera o
`data.json` a cada push que altere `data/` ou o script, baixando os `.xlsx`
do CDN pelos pointers. Nenhum comando manual é necessário.
