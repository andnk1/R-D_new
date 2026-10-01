# R&D Tax Credit – Feasibility Study (web app)

Internal CFO Associates tool. Six screens collect the facts, the existing calculation engine
works out the credit, and the existing feasibility generator produces the client-facing study
(HTML + PDF) and the calculation workbook.

## Run it on your computer

1. Double-click **run_local.bat** (first run installs the packages – takes a minute), or in a terminal:
   ```
   cd webapp
   pip install -r requirements.txt
   python -m playwright install chromium
   python app.py
   ```
2. Open **http://localhost:5000**. Stop it with Ctrl+C in the window.

## What's in the folder

| File / folder | What it does |
|---|---|
| `app.py` | The web app: one route per screen, saving answers, downloads |
| `study.py` | Turns the screen answers into the engine's inputs, runs the engine and the study generator |
| `workbook.py` | The downloadable expense workbook and reading it back when uploaded; reads .docx/.txt descriptions |
| `store.py` | Where answers are kept while you work (temporary folder, cleared after 24 h). Swap for a database later |
| `engine/` | **Unchanged copies** of `rd_credit_calculator_v2.py` and `rd_feasibility_generator.py`. From now on edit the engine here |
| `templates/` | The HTML screens (built from the approved mockups) |
| `static/style.css` | TaxPro UI kit – unchanged copy from 1099DA |
| `static/app.css` | Styles for this app only |
| `static/app.js` | On-screen helpers: live totals, tabs, filing-status line, blank/zero tags |
| `static/logo.png` | Logo used on the study cover |
| `Dockerfile` | Recipe Railway uses to build the app (includes Chromium for the PDF) |
| `requirements.txt` | Python packages the app needs |

## Deploying to Railway (your steps)

- Push the `webapp` folder as its own repo (it contains no client data).
- Railway finds the `Dockerfile` automatically. Set one variable: `SECRET_KEY` = any long random text.
- Optional: `STUDY_DIR` to choose where temporary study files go (default: the system temp folder).
- Put Cloudflare Access in front if you want a login.

## Adding a database later

Only `store.py` changes: write a class with the same four methods (`get`, `save`, `folder`, `delete`)
backed by Railway Postgres, and use it in `app.py` instead of `FileStore`.

## Known engine items (not changed yet – to fix together)

- The study PDF labels the method "Regular Credit Method" whenever a Regular figure exists, even when the
  higher ASC credit is the one shown. The Review screen shows the correct label.
- Regular credit, QRE year 11+: engine uses years 5–9; the rule allows any 5 years from 5–10.
- Items collected but not yet in the engine (noted on screen): first year with gross receipts, controlled
  groups, prior payroll elections, filing-status / 280C logic, special contract-research percentages,
  foreign R&E amortization, Section G.
