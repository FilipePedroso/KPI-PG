# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repository is

A fully static KPI dashboard for a P&G commercial team ("KPI Ranking P&G"), hosted on GitHub Pages from `/docs` on `main`. There is no backend: the browser fetches `docs/data.json` directly and does all filtering/ranking client-side. A Python script regenerates that JSON from Excel/Parquet source files whenever they change, via a GitHub Action that commits the result back to the repo.

**The real, deployed dashboard is `docs/index.html`.** It is self-contained — all CSS and JS are inline inside it (three `<script>` blocks + one `<style>` block). `docs/script.js` and `docs/style.css` exist but are **not referenced by `docs/index.html` and have no effect on the live site** — they are stale leftovers from an earlier refactor that was abandoned. Always edit `docs/index.html` directly for UI/behavior changes.

`public/dashboard.html` is a separate, **manually-diverged copy** of `docs/index.html`, kept only so the Lovable.dev Vite preview can serve something from `/public`. It is not auto-synced and is frequently behind — don't treat it as a source of truth, and if it needs to match `docs/index.html`, port changes over by hand.

`src/` is an unused TanStack Start / Lovable.dev scaffold. Its only real route (`src/routes/index.tsx`) just renders `<iframe src="/dashboard.html">` so Lovable's editor preview can display the static dashboard; `src/components/ui/*` (shadcn/ui) is boilerplate never imported anywhere. Don't build dashboard features in `src/` — it doesn't reach the deployed site.

## Commands

Package manager is **bun** (`bun.lock`, `bunfig.toml`).

- `bun run dev` — start the Vite dev server for the `src/` TanStack Start scaffold (only useful for the Lovable.dev preview iframe wrapper, not for iterating on the actual dashboard).
- `bun run build` — Vite production build (runs `prebuild` → `build:data` first).
- `bun run build:data` — regenerates `docs/data.json`, `public/data.json`, `docs/clientes.json`, `public/clientes.json`, `sem-estrutura.csv` by running `scripts/build_data.py` (tries `python3`, then `python`, then skips with a warning if neither has `openpyxl`/`pyarrow` available).
- `bun run lint` — ESLint over `**/*.ts,tsx` (does not cover `docs/index.html`, since its JS is inline in an HTML file).
- `bun run format` — Prettier write, repo-wide (`.prettierrc`: 100-char print width, double quotes, trailing commas everywhere).
- No test suite is configured in this repository.

To preview the actual dashboard locally, just open `docs/index.html` in a browser (or serve `docs/` with any static file server) — no build step is required for dashboard changes, only for regenerating `data.json`.

## Data pipeline architecture

Source files live in `data/` and are the real inputs an analyst updates (see `data/README.md` for the full column-level contract):

- `Dados_f_venda_total.parquet` / `Dados_f_ec_oniz.parquet` — fact tables (sales/positivação, and chaves/Platinum Points respectively).
- `Estrutura.xlsx` — sheets `d_comercial` (RV/SV/GV + UF hierarchy and display names), `d_metas` (targets per RV+UF), `d_clientes_braveo` (client potential), and a `data` sheet holding the single reference month used for daily-average math.
- `Dados_SC.xlsx` (optional) — for any UF present in this file, its sheets (`Pos Relação`, `Marcas Relação`, `Escolha Certa`, `Platinum Points`) **replace** the values `build_data.py` would otherwise compute from the fact tables for that UF, and can even inject `rv|uf` combinations absent from the fact tables entirely.

`scripts/build_data.py` aggregates everything by composite key `rv|uf`: faturamento by segment (Total/Escolha Certa/Store Platform/Alimentar/Farma), positivação (distinct CNPJs with value > 0), and ranking indicators (HFS, Farma, Always Noturno via a hardcoded EAN whitelist, Pampers with a minimum-volume noise floor) gated on `plat_ok` (Escolha Certa or Store Platform). Any `rv|uf` combo found in sales/metas but missing from `d_comercial` is auto-appended with placeholder names and reported in `sem-estrutura.csv`, so filters never silently drop data.

Editing `scripts/build_data.py` changes what KPIs are computed. Editing `docs/index.html` changes how they're displayed/filtered. The GitHub Action `.github/workflows/build-data.yml` re-runs the script and commits the regenerated JSON/CSV whenever `data/**` or the script itself changes on `main` (or via manual `workflow_dispatch`) — so `docs/data.json` is a generated artifact tracked in git, not something to hand-edit.

## Frontend architecture (`docs/index.html`)

- Cards: Faturamento (Total/EC/SP/Alimentar/Farma), Positivação (Total/Alimentar/Farma), Ranking (HFS/Farma/Always Noturno/Pampers), Escolha Certa (Positivação/Platinum Points), plus two "Resumo" roll-up cards (Ranking, Faturamento).
- Filtering is a custom multi-select chip/popover component over the RV/SV/GV/UF hierarchy (click = single-select, Ctrl/Cmd-click = multi-select), driving `window.dashboardAPI`.
- Print/export uses `html2canvas` (screenshot-to-clipboard/PNG) and `jspdf` + `jspdf-autotable` (PDF export), both loaded from a CDN. Resumo cards force a `snapshotting-grid` CSS layout (e.g. 2x2 instead of the default 3x3) when filtered down to a specific Gerente/Supervisor/Vendedor.
- Card expand/collapse uses FLIP animation; the Faturamento-Total card supports drilling into a per-Gerente/Supervisor/Vendedor team view.
