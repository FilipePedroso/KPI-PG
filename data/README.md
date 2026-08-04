# Atualização de Dados

## Arquivos necessários

Coloque **dois arquivos** nesta pasta (`data/`) e faça o deploy:

| Arquivo | Abas usadas |
|---------|-------------|
| `Dados.xlsx` | `f_venda_total`, `f_ec_oniz` |
| `Estrutura.xlsx` | `d_comercial`, `d_metas`, `d_clientes_braveo` |

O build roda `scripts/build_data.py` automaticamente e gera `docs/data.json` e `public/data.json`.

> Os `.xlsx` são grandes demais para o repositório Git (limite de 10 MB por arquivo), por isso ficam ignorados no `.gitignore` e são guardados como assets externos (`*.asset.json`). O script detecta automaticamente quando o arquivo local mudou e faz o upload do novo asset; se o arquivo não estiver localmente, ele baixa o último asset do CDN.

## Fluxo de atualização

1. Substitua o arquivo em `data/` (ex: `Estrutura.xlsx`).
2. Rode o deploy no Lovable ou execute localmente:
   ```bash
   python3 scripts/build_data.py
   ```
   O script sobe automaticamente o novo asset e atualiza o `*.asset.json` se o arquivo mudou.
3. Se publicar pelo **GitHub Pages**, commit também o `data.json` e o `*.asset.json` gerados, pois o GitHub Pages não roda o build novamente.

## Colunas esperadas

**`f_venda_total`** (Dados.xlsx)
`ds_uf`, `cd_vendedor`, `nr_cnpj_cpf`, `vl_financeiro`, `vl_faturamento` (1 = faturado, 0 = não faturado), `Plataforma` (`Escolha Certa` / `Store Platform`), `Canal` (`Alimentar` / `Farma`)

**`d_comercial`** (Estrutura.xlsx)
`RV`, `ds_uf`, `CONCATENAÇÃO RV + NOME`, `CONCATENAÇÃO SV + NOME`, `CONCATENAÇÃO CV + NOME`

**`d_metas`** (Estrutura.xlsx)
`RV`, `Uf`, `Meta Financeira Total`, `Meta Financeira Escolha Certa`, `Meta Financeira Store Platform`, `Meta Financeira Alimentar`, `Meta Financeira Farma`, `Objetivo Positivação Total`, `Objetivo Positivação Alimentar`, `Objetivo Positivação Farma`

## Relacionamento

`d_comercial.RV + ds_uf` ↔ `f_venda_total.cd_vendedor + ds_uf` ↔ `d_metas.RV + Uf`

## Rodar manualmente (opcional)

```bash
python3 scripts/build_data.py data/Dados.xlsx data/Estrutura.xlsx
```

