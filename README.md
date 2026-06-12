# Coffee Lab Invoice Processor

A zero-install web app for processing PDF invoices — drop invoices onto the
page and it detects the vendor, assigns cost category tags based on keyword
rules, overlays colored labels directly onto the PDF pages, and appends the
tagged pages to a running master PDF.

Everything runs **in your browser**: no uploads, no server, no login, no LLM.
The app lives in [`web/`](web/) and deploys to Vercel as a static site — see
[web/README.md](web/README.md) for usage, development, and deploy steps.

> The original Python desktop app (`invoice_processor.py`) is no longer the
> product — it remains in the repo as the **reference implementation** that
> the web pipeline is golden-tested against. The old tkinter UI, EXE build
> scripts, and `inbox/`/`processed/` folder workflow are kept only for that
> purpose.

---

## What it does

1. **Drop invoice PDFs** onto the page (or click to choose files)
2. **Detects the vendor** (Sysco, InnerMountain, Italco, Crested Bucha, Sisu,
   Vermont Sticky, or generic)
3. **Categorises line items** by matching descriptions against your keyword
   rules; Sysco order forms use their pre-printed category codes
4. **Groups multi-page invoices** — consecutive pages belonging to the same
   invoice get a single combined total tag on the first page
5. **Overlays colored tags** showing category code, name, and total amount,
   and **highlights individual dollar amounts** in the matching category color
6. **Downloads tagged PDFs** with a timestamp in the filename, plus an updated
   `master_invoices.pdf` — optionally pick your existing master first and the
   new pages are appended

Scanned/photographed invoices (no selectable text) are not supported yet; the
app shows a clear error for those.

## Categories & profiles

Categories are managed in the app via **Edit Categories** — no JSON editing
required. Each category has a code (e.g. `52000`), a name, a color, and a
keyword list matched case-insensitively against line item descriptions.

The default categories were built for a coffee shop:

| Code  | Name    | Color  |
|-------|---------|--------|
| 52000 | Merch   | Blue   |
| 50900 | F&B     | Yellow |
| 53100 | Kitchen | Purple |
| 61600 | Cafe    | Green  |

If you run a different kind of shop, edit the categories to match your own
inventory — and use **profiles** to keep separate category setups side by
side (e.g. one per shop), switchable from the dropdown in the header.
Everything is saved in your browser's localStorage and persists between
visits; **Export JSON** / **Import JSON** lets you back a setup up or move it
to another machine.

## Repo layout

```
InvoiceProcessor/
├── web/                    # The product: client-side web app (Vite + React)
│   ├── src/lib/            # TypeScript port of the processing pipeline
│   ├── tests/              # Golden tests against the Python reference
│   └── docs/               # Decision docs (e.g. access control options)
├── invoice_processor.py    # Python reference implementation (golden-test source)
├── tools/                  # generate_golden.py — regenerates golden test data
├── webredesign/            # Static design mock used for the UI redesign
├── config/, inbox/, processed/, master/   # Desktop-era folders, kept for the reference app
└── Start Invoice Processor.bat / .sh, build_exe.bat, build_app.sh   # Legacy launchers
```

## Supported vendors

| Vendor           | Detection keyword   | Notes                                          |
|------------------|---------------------|------------------------------------------------|
| Sysco            | `SYSCO`             | Uses pre-printed category codes; multi-page invoices automatically grouped (delivery copy pages merged with their summary page) |
| InnnerMountain   | `INNERMOUNTAIN`     | Keyword line-item matching                     |
| Italco           | `ITALCO`            | Keyword line-item matching                     |
| Crested Bucha    | `CRESTED BUCHA`     | Keyword line-item matching                     |
| Sisu Studios     | `SISU STUDIOS`      | Keyword line-item matching                     |
| Vermont Sticky   | `VERMONT STICKY`    | Keyword line-item matching                     |
| Gunnison County  | `GUNNISON COUNTY`   | Skipped — license/permit invoices, no tagging  |
| Generic          | *(fallback)*        | Keyword line-item matching                     |

## Running the Python reference (development only)

Only needed when changing pipeline logic — the web pipeline must stay in sync
with it, verified by the golden tests (see [web/README.md](web/README.md)):

```
pip install pypdf pdfplumber reportlab
python invoice_processor.py
```
