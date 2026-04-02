# Coffee Lab Invoice Processor

A desktop tool for processing PDF invoices — it scans each invoice, assigns cost category tags based on keyword rules, overlays colored labels directly onto the PDF pages, and appends the tagged pages to a running master PDF.

IMPORTANT NOTE: These categories have been made specifically for a coffee shop- if you are using this for a different resteraunt, follow the template but you will need to essentially input your own inventory, or have an LLM do it for you if you are comfortable with that.

Requires Python to work so pls install on your device if you want to use this. Still working out smooth installation so requires running python -m pip install pypdf pdfplumber reportlab pyinstaller

This is a known bug in the installer- in theory pyinstaller should be able to handle this?

---

## What it does

1. **Reads PDFs** from the `inbox/` folder (or via file picker)
2. **Detects the vendor** (Sysco, InnnerMountain, Italco, Crested Bucha, Sisu, Vermont Sticky, or generic)
3. **Categorises line items** by matching descriptions against keyword lists in `config/categories.json`
4. **Overlays colored tags** onto each page showing the category code, name, and total amount
5. **Highlights individual dollar amounts** on the page in the matching category color
6. **Saves tagged PDFs** to `processed/` with a timestamp in the filename
7. **Appends each tagged invoice** to `master/master_invoices.pdf` for a consolidated record

---

## Folder structure

```
CoffeeLabInvoiceProcessor/
├── inbox/                  # Drop invoice PDFs here before processing
├── processed/              # Tagged output PDFs land here
├── master/
│   └── master_invoices.pdf # All tagged invoices appended in order
├── config/
│   └── categories.json     # Keyword rules and category colors (auto-created on first run)
├── invoice_processor.py    # Main application
├── build_exe.bat           # Build script to produce a standalone Windows EXE
└── Start Invoice Processor.bat
```

---

## Categories

Defined in `config/categories.json`. Each category has:
- A **code** (e.g. `52000`)
- A **name** (e.g. `Merch`)
- A **color** (RGB, used for tags and highlights)
- A list of **keywords** matched case-insensitively against line item descriptions

Default categories:

| Code  | Name    | Color       |
|-------|---------|-------------|
| 52000 | Merch   | Blue        |
| 50900 | F&B     | Yellow      |
| 53100 | Kitchen | Purple      |
| 61600 | Cafe    | Green       |

Edit `config/categories.json` directly (the **Edit Keywords** button opens it in Notepad) and re-run to apply changes. The file is created with defaults on first launch if it doesn't exist.

---

## Running the app

**From source** (requires Python + dependencies):
```
pip install pypdf pdfplumber reportlab
python invoice_processor.py
```

**As a standalone EXE** (no Python needed on the target machine):
```
build_exe.bat        # run once to build dist\CoffeeLabInvoiceProcessor.exe
```
Then distribute or run `dist\CoffeeLabInvoiceProcessor.exe` directly. All folders (`inbox/`, `processed/`, `master/`, `config/`) are created next to the exe on first launch.

---

## Supported vendors

| Vendor           | Detection keyword   | Notes                                          |
|------------------|---------------------|------------------------------------------------|
| Sysco            | `SYSCO`             | Uses pre-printed category codes where present  |
| InnnerMountain   | `INNERMOUNTAIN`     | Keyword line-item matching                     |
| Italco           | `ITALCO`            | Keyword line-item matching                     |
| Crested Bucha    | `CRESTED BUCHA`     | Keyword line-item matching                     |
| Sisu Studios     | `SISU STUDIOS`      | Keyword line-item matching                     |
| Vermont Sticky   | `VERMONT STICKY`    | Keyword line-item matching                     |
| Gunnison County  | `GUNNISON COUNTY`   | Skipped — license/permit invoices, no tagging  |
| Generic          | *(fallback)*        | Keyword line-item matching                     |
