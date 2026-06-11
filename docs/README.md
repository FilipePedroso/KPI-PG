# KPI Ranking P&G — GitHub Pages

Esta pasta (`/docs`) é servida pelo GitHub Pages como site estático.

## Como ativar o GitHub Pages

1. No GitHub, abra o repositório → **Settings** → **Pages**.
2. Em **Source**, selecione **Deploy from a branch**.
3. Branch: `main` · Folder: `/docs` · clique em **Save**.
4. Em ~1 min o site fica disponível em `https://<seu-usuario>.github.io/<repo>/`.

## Estrutura

- `index.html` — dashboard (cópia de `public/dashboard.html`).
- `dados.xlsx` — (futuro) planilha com os dados reais. Basta colocar aqui ao lado do `index.html`.

## Quando os dados chegarem

Quando você tiver a planilha pronta, me avise e eu adapto o `index.html` para
ler `dados.xlsx` via [SheetJS](https://sheetjs.com/) (carregado por CDN, 100% no navegador — sem servidor).

## Renomear o repositório

Eu não consigo renomear repositórios do GitHub por você. Faça assim:
GitHub → repositório → **Settings** → campo **Repository name** → mude para
`KPI-Ranking-P-G` (o GitHub não aceita `&` em nomes de repo, então recomendo
`KPI-Ranking-PG` ou `kpi-ranking-pg`) → **Rename**.
