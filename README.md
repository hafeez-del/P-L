# VapeTime P&L Leadership Dashboard

A single-page, animated, dual-theme dashboard built on top of the merged `Combined_PnL_by_Store.xlsx`.

Credentials: **admin / admin123** (client-side only; keep the repo private).

## What's in the box
- `index.html` — the whole dashboard (opens in any browser)
- `assets/` — logo + login backgrounds (also inlined into the HTML)
- `data/pnl.json` — the raw monthly data for every store
- `scripts/merge_pnl.py` — Excel merger (takes the 5 single-store P&L files and builds the combined workbook with Parker-style formatting)
- `scripts/build_data.py` — rebuilds `data/pnl.json` from a new combined workbook and injects it into `index.html`
- `vercel.json` — one-line Vercel config

## Updating with next month's data
```
python scripts/merge_pnl.py path\to\PnL_folder
python scripts/build_data.py Combined_PnL_by_Store.xlsx
# commit and push
```

## Pages
- **Overview** — 6 KPIs, monthly bar+line, profit bridge (waterfall), operating expense donut, net income by store (click a bar to drill).
- **Metrics** — pick any metric (sales, cost, gross profit, inventory, each expense …), see trend, by-store comparison, heatmap and the full sub-category table.
- **Stores** — one card per store, side-by-side bar chart, monthly net income lines.
- **Rankings** — top performers / needs attention for every headline metric.
- **Statement** — the full P&L with frozen headers.

Built with Chart.js (CDN) and vanilla JS. No build step.
