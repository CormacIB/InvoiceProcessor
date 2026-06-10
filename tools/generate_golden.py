#!/usr/bin/env python3
"""
Dump golden reference data from the Python pipeline for the TypeScript port.
Runs extraction + categorisation on sample PDFs and writes JSON the web
app's tests compare against. Run from the repo root:

    venv/bin/python3 tools/generate_golden.py processed/Invoice.pdf ... > web/tests/golden.json
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))
import pdfplumber

from invoice_processor import (
    detect_vendor, get_page_categories, find_invoice_groups,
    find_amount_positions, load_config,
)


def golden_for(path: Path, config: dict) -> dict:
    pages_text = []
    pages = []
    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            pages_text.append(text)
            pages.append({"plumber_text": text})

    full_text = "\n".join(pages_text)
    vendor = detect_vendor(full_text)

    cats_per_page = []
    with pdfplumber.open(str(path)) as pdf:
        for i, page in enumerate(pdf.pages):
            cats, matched = get_page_categories(pages_text[i], config, vendor)
            highlights = find_amount_positions(page, matched, float(page.height))
            cats_per_page.append(cats)
            pages[i].update({
                "categories": cats,
                "matched_items": [[d, a, l] for d, a, l in matched],
                "highlight_count": len(highlights),
                "highlight_labels": [h[4] for h in highlights],
            })

    groups = find_invoice_groups(pages_text, vendor)
    group_cats = []
    for start, end in groups:
        agg: dict = {}
        for cats in cats_per_page[start:end]:
            for label, amount in cats.items():
                agg[label] = round(agg.get(label, 0.0) + amount, 2)
        group_cats.append({"start": start, "end": end, "categories": agg})

    return {
        "file": path.name,
        "vendor": vendor,
        "pages": pages,
        "groups": group_cats,
    }


def main():
    config = load_config()
    files = [Path(p) for p in sys.argv[1:]]
    out = {"config_source": "config/categories.json", "invoices": [golden_for(f, config) for f in files]}
    json.dump(out, sys.stdout, indent=2)


if __name__ == "__main__":
    main()
