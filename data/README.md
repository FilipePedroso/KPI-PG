# Atualização de Dados

## Arquivos necessários

Coloque **dois arquivos** nesta pasta (`data/`) e faça o deploy:

| Arquivo | Abas usadas |
|---------|-------------|
| `Dados.xlsx` | `f_venda_total` |
| `Estrutura.xlsx` | `d_comercial`, `d_metas` |

O build roda `scripts/build_data.py` automaticamente e gera `docs/data.json` e `public/data.json`.

> Os `.xlsx` são grandes demais para o repositório Git (limite de 10 MB por arquivo), por isso ficam ignorados no `.gitignore` e são guardados como assets externos (`*.asset.json`). O que vai para o site é o `data.json` gerado.

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
