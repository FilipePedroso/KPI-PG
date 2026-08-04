# Dados do KPI Ranking P&G

Esta pasta contém os arquivos Excel-fonte usados para gerar o dashboard.

## Arquivos esperados

| Arquivo | Abas utilizadas | Conteúdo |
|---|---|---|
| `Dados_f_venda_total.parquet` | — (tabela fato) | Vendas, positivação |
| `Dados_f_ec_oniz.parquet` | — (tabela fato) | Chaves e Platinum Points |
| `Dados_SC.xlsx` | `Pos Relação`, `Marcas Relação`, `Escolha Certa`, `Platinum Points` | Indicadores já tratados de SC (sobrescrevem o cálculo normal para a UF SC) |
| `Estrutura.xlsx` | `d_comercial`, `d_metas`, `d_clientes_braveo` | Hierarquia comercial, metas e potencial de clientes |

## Como atualizar os dados (GitHub)

O dashboard é estático e hospedado no GitHub Pages. A atualização dos dados é feita **diretamente no repositório do GitHub**, sem precisar rodar comandos manualmente.

1. Acesse o repositório `KPI Ranking P&G` no GitHub.
2. Navegue até a pasta `data/`.
3. Substitua o arquivo desejado (`Dados_f_venda_total.parquet`, `Dados_f_ec_oniz.parquet`, `Dados_SC.xlsx` ou `Estrutura.xlsx`) por upload na web ou via `git push`.
4. Commit com uma mensagem descritiva, por exemplo: `Atualização dados agosto/2026`.

Pronto. O resto é automático.

## O que acontece depois do commit

O workflow `.github/workflows/build-data.yml` dispara automaticamente sempre que algum arquivo dentro de `data/` for alterado. Ele:

1. Roda `scripts/build_data.py` usando os arquivos `.parquet` (tabelas fato) e `.xlsx` (estrutura) do repositório.
2. Gera/atualiza:
   - `docs/data.json`
   - `public/data.json`
   - `sem-estrutura.csv`
3. Commita esses arquivos de volta no repositório.

Em aproximadamente 1 minuto o GitHub Pages já reflete os novos dados no dashboard.

## Importante sobre o Lovable editor

O editor do Lovable não aceita uploads de arquivos acima de ~10 MB. Por isso, **arquivos Excel grandes devem ser atualizados diretamente no GitHub**, nunca arrastando para dentro do editor do Lovable.

O script `build_data.py` mantém compatibilidade com os arquivos de ponteiro `*.asset.json` (CDN) caso o `.xlsx` não esteja presente no repositório, mas o fluxo principal e recomendado é o de arquivos locais no GitHub.

## Estrutura esperada das abas

### Tabelas fato (parquet)

- **`Dados_f_venda_total.parquet`** — colunas esperadas: `cd_vendedor`, `ds_uf`, `vl_financeiro`, `vl_faturamento`, `Store Channel`, `Plataforma`, `Canal Ranking`, `nm_grupo`, `nm_produto`, `cd_gerente`, `cd_vendedor_superior`, `CNPJ`.
- **`Dados_f_ec_oniz.parquet`** — colunas esperadas: `nr_doc`, `nr_chave`, `ds_combo_sku_lista_ativacao`, `Platinum Point?`, `cd_vendedor`, `ds_sigla`, `Plataforma`.

### `Estrutura.xlsx`

- **`d_comercial`** — colunas esperadas: `RV`, `ds_uf`, `NOME`, `CONCATENAÇÃO RV + NOME`, `CONCATENAÇÃO SV + NOME`, `CONCATENAÇÃO CV + NOME`, `SV`, `GV`.
- **`d_metas`** — colunas esperadas: `RV`, `ds_uf`, `TT Positivação`, `OBJ PRODUTIVIDADE HFS`, `OBJ PRODUTIVIDADE FARMA`.
- **`d_clientes_braveo`** — colunas esperadas: `cd_vendedor`, `ds_uf`, `Potencial` (ou equivalente usado para potencial de positivação).

## Verificando se a atualização funcionou

- Acesse o GitHub Actions do repositório e confirme se a execução de `Build Dashboard Data` terminou com sucesso (✅ verde).
- No dashboard, o header mostra a data e hora da última geração dos dados.
- Caso alguma combinação de vendedor + UF não exista na estrutura, o arquivo `sem-estrutura.csv` será gerado com as combinações que faltavam e foram inseridas automaticamente como "CÓDIGO - -".

## `Dados_SC.xlsx` (fonte alternativa de SC)

Quando este arquivo está em `data/`, os indicadores abaixo passam a vir dele para
todas as combinações da UF `SC` (o cálculo pelas tabelas fato é desconsiderado):

| Card | Aba | Regra |
|---|---|---|
| Positivação HFS | `Pos Relação` | distintos de `CNPJ` onde `TIPO` = `HFS` |
| Positivação Farma | `Pos Relação` | distintos de `CNPJ` onde `TIPO` = `FARMA` |
| Positivação Always Noturno | `Marcas Relação` | distintos de `CNPJ` onde `TIPO` = `ALWAYS NOTURNO` |
| Positivação Pampers | `Marcas Relação` | distintos de `CNPJ` onde `TIPO` = `PAMPERS` |
| Chaves >= 2 | `Escolha Certa` | distintos de `CNPJ` |
| Platinum Points | `Platinum Points` | distintos de `CNPJ` + `GRUPO DE ATIVAÇÃO` |

Relacionamento com os filtros: `RV` (ou `cd_vendedor`/`Rv`) + `ds_uf`.
Basta substituir o arquivo em `data/` — o workflow roda automaticamente.
