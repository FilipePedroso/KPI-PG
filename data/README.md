# Atualização de Dados — Modo Automático

## Como funciona

Basta colocar seu arquivo `.xlsx` dentro desta pasta (`data/`) e fazer **deploy** (clicar em "Update" ou "Publish" no Lovable).

O build do projeto roda automaticamente o script `scripts/build_data.py`, que:
1. Lê o primeiro arquivo `.xlsx` encontrado nesta pasta
2. Extrai os dados das abas `d_comercial`, `f_vendas_total` e `d_metas_fin`
3. Gera os arquivos `docs/data.json` e `public/data.json`
4. Atualiza a data/hora no header do dashboard

**Não precisa rodar nenhum comando manual.**

## Estrutura esperada do Excel

As abas e colunas devem manter os mesmos nomes do modelo original:

| Aba | Colunas principais |
|-----|-------------------|
| `d_comercial` | `RV`, `CONCATENAÇÃO RV + NOME`, `CONCATENAÇÃO SV + NOME`, `CONCATENAÇÃO CV + NOME`, `ds_uf` |
| `f_vendas_total` | `cd_vendedor`, `ds_uf`, `vl_financeiro`, `Plataforma`, `Store Channel` |
| `d_metas_fin` | `vd`, `ds_uf`, `vl_objetivo`, `Plataforma`, `Store Channel` |

## Passo a passo

1. Suba ou substitua o arquivo `.xlsx` nesta pasta `data/`
2. Clique em **Publish** ou **Update** no Lovable
3. Pronto — os dados do dashboard serão atualizados automaticamente

> Dica: mantenha apenas um arquivo `.xlsx` nesta pasta para evitar confusão. O script sempre usa o último arquivo (ordem alfabética) se houver mais de um.
