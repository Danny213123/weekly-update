# Dashboard Builder (Fullstack)

This app now runs as a fullstack service with an Express API, SQLite persistence, and the existing dashboard UIs served from the backend. Node.js 18+ is recommended (Node 20 LTS tends to work best with `better-sqlite3`).

## Quick Start

1. Install dependencies

```bash
npm install
```

2. Install Playwright browser (required for export)

```bash
npx playwright install chromium
```

3. Run the server

```bash
npm run dev
```

Open:
- Dashboard Studio V2 (primary): `http://localhost:3000/` (auto-routes to `/v2`)
- Direct V2 link: `http://localhost:3000/v2`

## JSON Import/Export

Use the “JSON Import/Export” panel in the sidebar to paste JSON and apply it to the current mode. You can also copy the current state as JSON or use the built-in template button.

Template files:
- `public/json-template.monthly.json`
- `public/json-template.yearly.json`

## Theme Studio

Use the “Theme Studio” panel to customize colors, opacity, typography, and panel sizing. Per‑KPI gradient and text colors are editable under each KPI card.

## CLI

The CLI mirrors the git-ui styling and lets you build exports, list saves, and import/export JSON directly from the local SQLite database.

Run interactive mode:

```bash
npm run cli
```

Examples:

```bash
dashboard-studio list
dashboard-studio history --type monthly
dashboard-studio export-json --type monthly --out exports/monthly.json
dashboard-studio import-json --type yearly --file public/json-template.yearly.json --save-version
dashboard-studio export-png --type monthly --out exports/monthly.png
dashboard-studio build
dashboard-studio launch --open
dashboard-studio quick-launch --open
```

## Export Output

Exports are generated server-side using Playwright + Chromium for pixel-perfect layout fidelity.

## Data Storage

SQLite data lives in `data/app.db` by default. You can override with `DB_PATH` in `.env`.
