#!/usr/bin/env python3
"""
Coffee Lab Invoice Processor
Scans PDF invoices, assigns cost categories via keyword rules,
overlays colored tags, and appends pages to a master PDF.
"""

import io
import json
import os
import re
import shutil
import subprocess
import sys
import traceback
from datetime import datetime
from pathlib import Path

import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.lib.colors import Color
from reportlab.pdfgen import canvas as rl_canvas
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
CONFIG_FILE = BASE_DIR / "config" / "categories.json"
INBOX_DIR   = BASE_DIR / "inbox"
PROCESSED_DIR = BASE_DIR / "processed"
MASTER_DIR  = BASE_DIR / "master"
MASTER_PDF  = MASTER_DIR / "master_invoices.pdf"

# ── Default config (written if categories.json missing) ───────────────────────
DEFAULT_CONFIG = {
    "categories": [
        {
            "code": "52000", "name": "Merch",
            "color": [0, 191, 255],
            "keywords": [
                "yerba mate", "red bull", "open water", "recess",
                "amy & brian", "coconut juice", "sticker", "card",
                "painting", "hot cocoa", "maple syrup", "bulk gallon",
                "candy", "colored bites", "cookie", "cookies",
                "peanut butter", "triple chocolate chunk", "donut cake"
            ]
        },
        {
            "code": "50900", "name": "F&B",
            "color": [255, 200, 0],
            "keywords": [
                "kombucha", "burrito", "brkfst", "muffin", "croissant",
                "dough croissant", "gelato", "banana walnut", "loaf",
                "milk", "creamer", "half & half", "yogurt", "tamale",
                "pretzel", "churro", "donut churro", "mushroom", "chai",
                "sprite", "coke", "mineral water", "topochi", "salsa",
                "cholula", "noodle cup", "sandwich", "apple cid", "big b",
                "spiced apple", "ginger grapefruit", "pineapple peach",
                "vanilla bean gelato", "energy organic", "drink energy",
                "cream soda", "sparkling water", "galette"
            ]
        },
        {
            "code": "53100", "name": "Kitchen",
            "color": [153, 51, 255],
            "keywords": [
                "cup/lid", "cup lid", "lid plas", "compostable",
                "ecoprod", "fiber", "sip lid", "sleeve", "straw",
                "napkin", "glove", "lid hot cup"
            ]
        },
        {
            "code": "61600", "name": "cafe",
            "color": [100, 200, 100],
            "keywords": [
                "rinse aid", "keyston", "chemical", "janitorial",
                "detergent", "sanitizer", "cleaning", "low temp liq disp"
            ]
        }
    ]
}

# ── Setup ─────────────────────────────────────────────────────────────────────
def ensure_dirs():
    for d in [INBOX_DIR, PROCESSED_DIR, MASTER_DIR, CONFIG_FILE.parent]:
        d.mkdir(parents=True, exist_ok=True)


def load_config() -> dict:
    if not CONFIG_FILE.exists():
        CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
        with open(CONFIG_FILE, "w") as f:
            json.dump(DEFAULT_CONFIG, f, indent=2)
        return DEFAULT_CONFIG
    with open(CONFIG_FILE) as f:
        return json.load(f)


# ── Vendor detection ──────────────────────────────────────────────────────────
def detect_vendor(text: str) -> str:
    t = text.upper()
    if "SYSCO" in t:
        return "sysco"
    if "INNERMOUNTAIN" in t:
        return "innermountain"
    if "ITALCO" in t:
        return "italco"
    if "CRESTED BUCHA" in t:
        return "crested_bucha"
    if "SISU STUDIOS" in t:
        return "sisu"
    if "VERMONT STICKY" in t:
        return "vermont_sticky"
    if "GUNNISON COUNTY" in t:
        return "skip"   # license/permit invoices — no category tagging
    return "generic"


# ── Sysco: extract pre-printed category codes + amounts ───────────────────────
def extract_sysco_categories(text: str) -> dict:
    """
    For Sysco SINGLE-ITEM order forms (not delivery copy invoices),
    the category code is pre-printed in the format:
        50900 F&B
        12/10/25          <- date line (optional)
        $85.00
    Extract all such pairs and return {label: amount}.
    """
    cats = {}
    # Allow an optional intervening line (e.g. a date) between code and amount
    pat = re.compile(
        r'(\d{5})\s+([A-Za-z&/ ]+?)\s*\n[^\n]*\n\s*\$([\d,]+\.\d{2})',
        re.MULTILINE
    )
    for m in pat.finditer(text):
        code   = m.group(1).strip()
        name   = m.group(2).strip()
        amount = float(m.group(3).replace(",", ""))
        label  = f"{code} {name}"
        cats[label] = cats.get(label, 0.0) + amount

    # Fallback: code+name then $ directly on next line (no date)
    if not cats:
        pat2 = re.compile(
            r'(\d{5})\s+([A-Za-z&/ ]+?)\s*\n\s*\$([\d,]+\.\d{2})',
            re.MULTILINE
        )
        for m in pat2.finditer(text):
            code   = m.group(1).strip()
            name   = m.group(2).strip()
            amount = float(m.group(3).replace(",", ""))
            label  = f"{code} {name}"
            cats[label] = cats.get(label, 0.0) + amount
    return cats


# ── Generic: extract (description, amount) line items ────────────────────────
_SKIP_WORDS = {
    "total", "subtotal", "tax", "balance", "payment", "due",
    "amount", "price", "extended", "invoice", "surcharge",
    "misc", "page", "terms", "group total", "order summary",
    "remit", "cases", "split", "cube", "gross",
    "sysco", "confidential", "paca", "driver", "sign",
    "important", "authorized", "retains", "receivables", "proceeds",
    "dispute", "representative", "capacity", "claimants",
    "open:", "close:", "5:00 am", "9:00 pm",   # Sysco footer time strings
    "fuel surcharge", "misc charges", "misc tax",
}

def extract_line_items(text: str) -> list:
    """
    Returns [(description, amount), ...] from arbitrary invoice text.
    Grabs lines that end with a dollar amount and aren't header/footer lines.
    """
    items = []
    # Match: any text ... $XX.XX  OR  text ... XX.XX  at end of line
    line_re = re.compile(
        r'^(.+?)\s+\$?([\d,]{1,7}\.\d{2})\s*$',
        re.MULTILINE
    )
    for m in line_re.finditer(text):
        desc   = m.group(1).strip()
        amount = float(m.group(2).replace(",", ""))

        # Skip zero-dollar lines and unreasonably large amounts
        if amount <= 0 or amount > 49999:
            continue
        # Skip lines that are clearly totals/headers/footers
        desc_low = desc.lower()
        if any(w in desc_low for w in _SKIP_WORDS):
            continue
        # Also skip Sysco footer lines like "OPEN: 5:00 AM  CLOSE: 9:00 PM"
        if re.search(r'\d+:\d{2}\s*(am|pm)', desc_low):
            continue
        # Skip very short descriptions (likely column headers)
        if len(desc) < 4:
            continue

        items.append((desc, amount))
    return items


# ── Keyword categorisation ────────────────────────────────────────────────────
def categorize_items(items: list, config: dict) -> tuple:
    """
    Match each (description, amount) against keyword lists.
    Returns ({label: total_amount}, [(desc, amount, label), ...]).
    """
    cats_cfg = config["categories"]
    totals: dict = {}
    matched_items = []
    uncategorized = 0.0

    for desc, amount in items:
        desc_low = desc.lower()
        matched = False
        for cat in cats_cfg:
            if any(kw.lower() in desc_low for kw in cat["keywords"]):
                label = f"{cat['code']} {cat['name']}"
                totals[label] = totals.get(label, 0.0) + amount
                matched_items.append((desc, amount, label))
                matched = True
                break
        if not matched:
            uncategorized += amount

    if not totals and uncategorized > 0:
        totals["REVIEW"] = round(uncategorized, 2)

    return {k: round(v, 2) for k, v in totals.items()}, matched_items


def extract_invoice_total(text: str) -> float:
    """Pull the headline total from invoice text as a fallback."""
    patterns = [
        r'invoice\s+total\s*\$?\s*([\d,]+\.\d{2})',
        r'total\s+due\s*\$?\s*([\d,]+\.\d{2})',
        r'total\s+sales\s*\$?\s*([\d,]+\.\d{2})',
        r'net\s+amount\s+due\s*\$?\s*([\d,]+\.\d{2})',
        r'balance\s*\$?\s*([\d,]+\.\d{2})',
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return float(m.group(1).replace(",", ""))
    return 0.0


# ── Per-page category dispatch ────────────────────────────────────────────────
def adjust_for_surcharges(cats: dict, text: str) -> dict:
    """
    If the invoice total > sum of categorized items (e.g. fuel surcharge,
    small taxes), add the difference to the largest category.
    Only adjusts upward; ignores discrepancies over $30 to avoid errors.
    """
    invoice_total = extract_invoice_total(text)
    if invoice_total <= 0:
        return cats
    item_sum = sum(cats.values())
    diff = round(invoice_total - item_sum, 2)
    if 0 < diff <= 30:
        largest = max(cats, key=cats.get)
        cats[largest] = round(cats[largest] + diff, 2)
    return cats


def get_page_categories(text: str, config: dict, vendor: str) -> tuple:
    """
    Returns ({label: amount}, [(desc, amount, label), ...]) for a single page.
    matched_items is empty for Sysco regex-extracted pages (no line positions).
    """
    if vendor == "skip":
        return {}, []

    if vendor == "sysco":
        is_delivery = "DELIVERY COPY" in text.upper()
        if not is_delivery:
            cats = extract_sysco_categories(text)
            if cats:
                return cats, []   # regex path — no per-line positions
        t_up = text.upper()
        if "ORDER SUMMARY" in t_up and not re.search(r'\b(DAIRY|FROZEN|CANNED|PAPER|CHEMICAL)\b', t_up):
            return {}, []

    # All other vendors: keyword match on line items
    items = extract_line_items(text)
    if items:
        cats, matched = categorize_items(items, config)
        if cats:
            cats = adjust_for_surcharges(cats, text)
            return cats, matched

    # Last resort: whole-invoice keyword scan using total
    total = extract_invoice_total(text)
    if total > 0:
        cats, matched = categorize_items([(text[:800].lower(), total)], config)
        if cats:
            return cats, matched
        return {"REVIEW": round(total, 2)}, []

    return {}, []


# ── Amount position lookup ─────────────────────────────────────────────────────
def find_amount_positions(plumber_page, matched_items: list, page_h: float) -> list:
    """
    For each matched line item, find the bounding box of its dollar amount
    on the page using pdfplumber word positions.

    When multiple items share the same amount, positions are assigned
    top-to-bottom to match the order items appear on the page.

    Returns [(x0, y_bottom, x1, y_top, label), ...] in reportlab coordinates
    (origin bottom-left).
    """
    if not matched_items:
        return []

    words = plumber_page.extract_words()

    # Build a map: amount_value -> list of word bounding boxes, sorted top-to-bottom
    # Keep only the rightmost column for each amount (extended/total column)
    from collections import defaultdict
    word_map: dict = defaultdict(list)
    for w in words:
        cleaned = w["text"].lstrip("$").replace(",", "")
        try:
            val = round(float(cleaned), 2)
            word_map[val].append(w)
        except ValueError:
            pass

    # For each amount value, keep only the rightmost column, sorted top-to-bottom
    rightmost_map: dict = {}
    for val, wlist in word_map.items():
        max_x0 = max(w["x0"] for w in wlist)
        col = [w for w in wlist if abs(w["x0"] - max_x0) < 5]
        col.sort(key=lambda w: w["top"])
        rightmost_map[val] = col

    # Assign positions to matched items in the order they appear
    used: dict = defaultdict(int)
    highlights = []

    for desc, amount, label in matched_items:
        key = round(amount, 2)
        col = rightmost_map.get(key, [])
        idx = used[key]
        if idx >= len(col):
            continue
        best = col[idx]
        used[key] += 1

        pad = 2
        x0    = best["x0"] - pad
        x1    = best["x1"] + pad
        y_bot = page_h - best["bottom"] - pad
        y_top = page_h - best["top"]    + pad
        highlights.append((x0, y_bot, x1, y_top, label))

    return highlights


# ── PDF tag overlay ───────────────────────────────────────────────────────────
def _build_color_map(config: dict) -> dict:
    cmap = {}
    for cat in config["categories"]:
        label = f"{cat['code']} {cat['name']}"
        cmap[label] = cat["color"]
    return cmap


def create_tag_overlay(page_w: float, page_h: float,
                        categories: dict, config: dict,
                        highlights: list = None) -> bytes:
    """
    Draws colored tag boxes near the top-left of the page, and optionally
    highlights each matched dollar amount with its category color.
    Returns bytes of a single-page PDF overlay.
    """
    color_map = _build_color_map(config)
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=(page_w, page_h))

    # ── Amount highlights ──────────────────────────────────────────────────
    if highlights:
        for x0, y_bot, x1, y_top, label in highlights:
            rgb = color_map.get(label, [180, 180, 180])
            r, g, b = [v / 255 for v in rgb]
            c.setFillColor(Color(r, g, b, alpha=0.55))
            c.setStrokeColor(Color(0, 0, 0, alpha=0))
            c.rect(x0, y_bot, x1 - x0, y_top - y_bot, fill=1, stroke=0)

    # ── Tag boxes ─────────────────────────────────────────────────────────
    TAG_W, TAG_H = 135, 40
    TAG_X_START  = 22
    TAG_Y_TOP    = page_h - 44
    GAP          = 6

    x = TAG_X_START
    for label, amount in categories.items():
        rgb = color_map.get(label, [180, 180, 180])
        r, g, b = [v / 255 for v in rgb]
        dr, dg, db = r * 0.65, g * 0.65, b * 0.65  # darker border

        # Background rectangle
        c.setFillColor(Color(r, g, b, alpha=0.72))
        c.setStrokeColor(Color(dr, dg, db, alpha=0.85))
        c.setLineWidth(1.5)
        c.rect(x, TAG_Y_TOP - TAG_H, TAG_W, TAG_H, fill=1, stroke=1)

        # Label text  (e.g. "52000 Merch")
        c.setFillColor(Color(0, 0, 0))
        c.setFont("Helvetica-Bold", 9)
        c.drawString(x + 6, TAG_Y_TOP - 15, label)

        # Amount text  (e.g. "$446.90")
        c.setFont("Helvetica-Bold", 11)
        c.drawString(x + 6, TAG_Y_TOP - 30, f"${amount:,.2f}")

        x += TAG_W + GAP

    c.save()
    buf.seek(0)
    return buf.read()


def overlay_tags_on_pdf(input_path: Path,
                         cats_per_page: list,
                         highlights_per_page: list,
                         config: dict) -> bytes:
    """Merge tag overlays (and amount highlights) onto each page; return full PDF bytes."""
    reader = PdfReader(str(input_path))
    writer = PdfWriter()

    for i, page in enumerate(reader.pages):
        cats       = cats_per_page[i]       if i < len(cats_per_page)       else {}
        highlights = highlights_per_page[i] if i < len(highlights_per_page) else []
        if cats or highlights:
            pw = float(page.mediabox.width)
            ph = float(page.mediabox.height)
            overlay_bytes = create_tag_overlay(pw, ph, cats, config, highlights)
            overlay_page  = PdfReader(io.BytesIO(overlay_bytes)).pages[0]
            page.merge_page(overlay_page)
        writer.add_page(page)

    out = io.BytesIO()
    writer.write(out)
    out.seek(0)
    return out.read()


# ── Master PDF ────────────────────────────────────────────────────────────────
def append_to_master(processed_bytes: bytes, master_path: Path):
    writer = PdfWriter()
    if master_path.exists():
        for page in PdfReader(str(master_path)).pages:
            writer.add_page(page)
    for page in PdfReader(io.BytesIO(processed_bytes)).pages:
        writer.add_page(page)
    with open(master_path, "wb") as f:
        writer.write(f)


# ── Main pipeline ─────────────────────────────────────────────────────────────
def process_invoice(input_path: Path, config: dict, log=print) -> bool:
    log(f"  File : {input_path.name}")
    try:
        pages_text    = []
        plumber_pages = []
        with pdfplumber.open(str(input_path)) as pdf:
            for page in pdf.pages:
                pages_text.append(page.extract_text() or "")
                plumber_pages.append(page)

        if not pages_text or not any(pages_text):
            log("  ⚠  No text could be extracted — is this a scanned image PDF?")
            return False

        full_text = "\n".join(pages_text)
        vendor    = detect_vendor(full_text)
        log(f"  Vendor: {vendor}")

        cats_per_page       = []
        highlights_per_page = []

        for i, (text, plumber_page) in enumerate(zip(pages_text, plumber_pages)):
            cats, matched = get_page_categories(text, config, vendor)
            cats_per_page.append(cats)

            ph = float(plumber_page.height)
            highlights = find_amount_positions(plumber_page, matched, ph)
            highlights_per_page.append(highlights)

            if cats:
                summary = "  |  ".join(f"{k}  ${v:,.2f}" for k, v in cats.items())
                log(f"  Page {i+1}: {summary}")
            else:
                log(f"  Page {i+1}: (no tag)")

        if not any(cats_per_page):
            log("  ⚠  No categories found — check keywords in config/categories.json")

        processed_bytes = overlay_tags_on_pdf(
            input_path, cats_per_page, highlights_per_page, config
        )

        timestamp   = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_name    = f"{input_path.stem}_tagged_{timestamp}.pdf"
        out_path    = PROCESSED_DIR / out_name
        out_path.write_bytes(processed_bytes)
        log(f"  Saved : {out_path.name}")

        append_to_master(processed_bytes, MASTER_PDF)
        log(f"  Master: appended -> {MASTER_PDF.name}")
        return True

    except Exception as e:
        log(f"  ERROR: {e}")
        log(traceback.format_exc())
        return False


# ── GUI ───────────────────────────────────────────────────────────────────────
class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Coffee Lab Invoice Processor")
        self.root.geometry("740x520")
        self.root.resizable(True, True)
        self._build()

    # ── UI construction ───────────────────────────────────────────────────────
    def _build(self):
        # Button row
        btn_row = tk.Frame(self.root, pady=8, padx=10)
        btn_row.pack(fill=tk.X)

        def btn(parent, text, cmd, bg):
            return tk.Button(parent, text=text, command=cmd,
                             bg=bg, fg="white", font=("Arial", 10, "bold"),
                             width=18, height=2, relief=tk.FLAT, cursor="hand2")

        btn(btn_row, "Select File(s)...",    self.select_files,    "#8C4CAF").pack(side=tk.LEFT, padx=4)
        btn(btn_row, "Process Inbox",        self.process_inbox,   "#8C4CAF").pack(side=tk.LEFT, padx=4)
        btn(btn_row, "Edit Keywords",        self.open_config,     "#8C4CAF").pack(side=tk.LEFT, padx=4)
        btn(btn_row, "Clear Master PDF",     self.clear_master,    "#8C4CAF").pack(side=tk.LEFT, padx=4)

        # Status bar
        self.status_var = tk.StringVar(value="Ready.")
        tk.Label(self.root, textvariable=self.status_var,
                 anchor="w", font=("Arial", 9), fg="#444",
                 padx=12).pack(fill=tk.X)

        # Log panel
        frame = tk.LabelFrame(self.root, text=" Log ", padx=4, pady=4, font=("Arial", 9))
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 4))
        self.log_box = scrolledtext.ScrolledText(frame, font=("Courier New", 9),
                                                  bg="#1e1e1e", fg="#d4d4d4",
                                                  insertbackground="white")
        self.log_box.pack(fill=tk.BOTH, expand=True)

        # Footer paths
        footer = (f"Inbox: {INBOX_DIR}    |    "
                  f"Processed: {PROCESSED_DIR}    |    "
                  f"Master: {MASTER_PDF}")
        tk.Label(self.root, text=footer, anchor="w",
                 font=("Courier New", 7), fg="#999",
                 padx=10).pack(fill=tk.X, pady=(0, 4))

    # ── Logging helpers ───────────────────────────────────────────────────────
    def log(self, msg: str):
        self.log_box.insert(tk.END, msg + "\n")
        self.log_box.see(tk.END)
        self.root.update_idletasks()

    def status(self, msg: str):
        self.status_var.set(msg)
        self.root.update_idletasks()

    # ── Actions ───────────────────────────────────────────────────────────────
    def select_files(self):
        paths = filedialog.askopenfilenames(
            title="Select Invoice PDF(s)",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        if paths:
            self._run_batch([Path(p) for p in paths])

    def process_inbox(self):
        pdfs = sorted(INBOX_DIR.glob("*.pdf"))
        if not pdfs:
            messagebox.showinfo("Inbox Empty",
                                f"No PDF files found in:\n{INBOX_DIR}\n\n"
                                "Drop invoice PDFs into the inbox folder and try again.")
            return
        self._run_batch(pdfs, move_after=True)

    def _run_batch(self, paths: list, move_after: bool = False):
        config  = load_config()
        total   = len(paths)
        success = 0

        self.log(f"\n{'─'*60}")
        self.log(f"Batch: {total} file(s)  —  {datetime.now():%Y-%m-%d %H:%M}")
        self.log(f"{'─'*60}")

        for path in paths:
            self.log(f"\n> {path.name}")
            ok = process_invoice(path, config, self.log)
            if ok:
                success += 1
                if move_after:
                    dest = PROCESSED_DIR / path.name
                    # Avoid name collision
                    if dest.exists():
                        ts  = datetime.now().strftime("%Y%m%d_%H%M%S")
                        dest = PROCESSED_DIR / f"{path.stem}_orig_{ts}.pdf"
                    shutil.move(str(path), str(dest))

        result = f"Done — {success}/{total} succeeded."
        self.log(f"\n{'─'*60}")
        self.log(result)
        self.status(result)

    def open_config(self):
        subprocess.Popen(f'notepad "{CONFIG_FILE}"')
        self.log("Opened categories.json in Notepad — save and re-run to apply changes.")

    def clear_master(self):
        if not MASTER_PDF.exists():
            messagebox.showinfo("Nothing to clear", "Master PDF does not exist yet.")
            return
        if messagebox.askyesno("Confirm",
                               f"Delete master PDF?\n{MASTER_PDF}\n\nThis cannot be undone."):
            MASTER_PDF.unlink()
            self.log("Master PDF deleted.")
            self.status("Master PDF cleared.")


# ── Entry point ───────────────────────────────────────────────────────────────
def main():
    ensure_dirs()
    root = tk.Tk()
    try:
        # Set a window icon if one exists next to the script
        icon = BASE_DIR / "icon.ico"
        if icon.exists():
            root.iconbitmap(str(icon))
    except Exception:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
