# Invoice Processor — Web App

The browser version of the Coffee Lab Invoice Processor. Everything runs
client-side: PDFs are parsed, tagged, and merged **in your browser** — no
file ever leaves your machine, there is no server, no login, and no LLM.

[invoice_processor.py](../invoice_processor.py) in the repo root is the
reference implementation; the pipeline in `src/lib/` is a function-for-function
TypeScript port of it, verified by golden tests against real invoices.

## What it does

1. Drop one or more invoice PDFs onto the page
2. Vendor is detected (Sysco, Italco, InnerMountain, … or generic)
3. Line items are categorised by keyword rules; Sysco order forms use their
   pre-printed category codes; multi-page invoices are grouped
4. Colored category tags and dollar-amount highlights are drawn onto the PDF
5. Download each tagged PDF, plus an updated master PDF — optionally pick
   your existing `master_invoices.pdf` first and the new pages are appended

Scanned/photographed invoices (no selectable text) are not supported yet;
the app shows a clear error for those.

## Categories & profiles

Click **Edit Categories** to manage categories, keywords, and colors. The
header has a **Profile** dropdown for keeping several independent category
setups side by side (e.g. different shops using the app for different
purposes) — **+ New** creates a profile (copying the current categories or
starting from the defaults), and profiles can be renamed and deleted.

Everything is stored in your browser (localStorage) and persists between
visits, so each machine keeps its own profiles — there are no accounts.
A config saved by the pre-profiles version of the app is migrated into a
"Default" profile automatically on first load. Use **Export JSON** /
**Import JSON** to back up a profile's categories or copy them to another
machine. The built-in defaults come from
[src/lib/defaultConfig.json](src/lib/defaultConfig.json) (a snapshot of the
desktop app's `config/categories.json`).

## Development

```bash
npm install
npm run dev      # local dev server
npm test         # golden tests against the Python reference output
npm run build    # type-check + production build into dist/
```

### Golden tests

`tests/golden.json` holds the Python pipeline's output (extracted text,
per-page category totals, invoice groups) for the sample invoices. If you
change pipeline logic, change it in invoice_processor.py too and regenerate:

```bash
# from the repo root
venv/bin/python3 tools/generate_golden.py processed/Invoice.pdf processed/Invoice2.pdf \
  processed/Invoice3.pdf "processed/Invoice3 copy.pdf" > web/tests/golden.json
```

## Deploying to Vercel

1. Push this repo to GitHub
2. In Vercel: **Add New Project** → import the repo
3. Set **Root Directory** to `web` — Vercel auto-detects Vite
4. Deploy. The free (Hobby) tier is sufficient indefinitely: this is a fully
   static site with no serverless functions and no storage.

The deployed URL is open to anyone who has it; options for restricting access
(not yet implemented) are written up in
[docs/access-control-options.md](docs/access-control-options.md).
