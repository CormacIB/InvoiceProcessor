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
import sys
import traceback
from datetime import datetime
from pathlib import Path

import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.lib.colors import Color
from reportlab.pdfgen import canvas as rl_canvas
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QDialog,
    QVBoxLayout, QHBoxLayout, QGridLayout,
    QPushButton, QLabel, QListWidget, QTextEdit,
    QLineEdit, QGroupBox, QFileDialog, QMessageBox,
    QAbstractItemView,
)
from PySide6.QtGui import QFont, QTextCursor, QIcon, QDesktopServices
from PySide6.QtCore import Qt, QUrl

# ── Paths ─────────────────────────────────────────────────────────────────────
# When frozen by PyInstaller, use the exe's location so that
# inbox/processed/master folders are created next to the exe.
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
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
                "peanut butter", "triple chocolate chunk", "donut cake",
                "12 oz", "12oz"
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
                "cream soda", "sparkling water", "galette",
                "5 lb", "5lb"
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
    "amount", "price", "extended", "invoice",
    "misc", "page", "terms", "group total", "order summary",
    "remit", "cases", "split", "cube", "gross",
    "sysco", "confidential", "paca", "driver", "sign",
    "important", "authorized", "retains", "receivables", "proceeds",
    "dispute", "representative", "capacity", "claimants",
    "open:", "close:", "5:00 am", "9:00 pm",   # Sysco footer time strings
    "misc charges", "misc tax",
}

def extract_line_items(text: str) -> list:
    """
    Returns [(description, amount), ...] from arbitrary invoice text.
    Grabs lines that end with a dollar amount and aren't header/footer lines.
    """
    items = []
    # Match: any text ... $XX.XX  OR  text ... XX.XX  at end of line
    line_re = re.compile(
        r'^(.+?)\s+\$?([\d,]{0,7}\.\d{2})\s*[A-Za-z*]?\s*$',
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
            if "surcharge" not in desc_low and "retail delivery fee" not in desc_low:
                continue
        # Skip Sysco lines marked OUT (not delivered, no extended price charged)
        if re.match(r'^(?:[a-z]\s+)?out\s', desc_low):
            continue
        # Also skip Sysco footer lines like "OPEN: 5:00 AM  CLOSE: 9:00 PM"
        if re.search(r'\d+:\d{2}\s*(am|pm)', desc_low):
            continue
        # Skip very short descriptions (likely column headers)
        if len(desc) < 4:
            continue
        # Skip descriptions with no letters — catches bare item codes (e.g. "7267150")
        # and stray symbol lines (e.g. "****") that aren't real line items
        if not re.search(r'[a-zA-Z]', desc):
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
            if "FUEL SURCHARGE" not in t_up:
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
        cleaned = w["text"].lstrip("$").replace(",", "").rstrip("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz*")
        try:
            val = round(float(cleaned), 2)
            word_map[val].append(w)
        except ValueError:
            pass

    # For each amount value, keep only the rightmost column, sorted top-to-bottom
    rightmost_map: dict = {}
    for val, wlist in word_map.items():
        max_x1 = max(w["x1"] for w in wlist)
        col = [w for w in wlist if abs(w["x1"] - max_x1) < 10]
        col.sort(key=lambda w: w["top"])
        rightmost_map[val] = col

    # Assign positions to matched items in the order they appear
    used: dict = defaultdict(int)
    highlights = []

    for _, amount, label in matched_items:
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


# ── GUI helpers ───────────────────────────────────────────────────────────────
def _darken(hex_color: str, factor: float = 0.75) -> str:
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return f"#{int(r*factor):02x}{int(g*factor):02x}{int(b*factor):02x}"


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


# ── Invoice boundary detection ────────────────────────────────────────────────
def find_invoice_groups(pages_text: list, vendor: str) -> list:
    """
    Groups consecutive pages that belong to the same invoice.
    Returns [(start_idx, end_idx), ...] where end_idx is exclusive.

    Sysco: a new invoice begins on any non-delivery-copy page that contains
    pre-printed category codes.  Delivery-copy pages trail the preceding
    summary page and belong to the same invoice group.

    Other vendors: "Page 1 of N" signals a new invoice; otherwise each page
    is its own group (preserving the original per-page behaviour).
    """
    if len(pages_text) <= 1:
        return [(0, len(pages_text))]

    groups = []
    group_start = 0

    for i, text in enumerate(pages_text):
        if i == 0:
            continue

        if vendor == "sysco":
            is_delivery = "DELIVERY COPY" in text.upper()
            if not is_delivery and extract_sysco_categories(text):
                groups.append((group_start, i))
                group_start = i
        else:
            if re.search(r'page\s+1\s+of\s+\d+', text, re.IGNORECASE):
                groups.append((group_start, i))
                group_start = i

    groups.append((group_start, len(pages_text)))
    return groups


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

        # Detect invoice boundaries within the PDF and aggregate each group's
        # categories onto the first page of that group only.
        invoice_groups = find_invoice_groups(pages_text, vendor)
        cats_for_overlay = [{} for _ in cats_per_page]
        for start, end in invoice_groups:
            group_cats: dict = {}
            for cats in cats_per_page[start:end]:
                for label, amount in cats.items():
                    group_cats[label] = round(group_cats.get(label, 0.0) + amount, 2)
            cats_for_overlay[start] = group_cats
            if group_cats and end - start > 1:
                page_range = f"{start+1}–{end}"
                summary = "  |  ".join(f"{k}  ${v:,.2f}" for k, v in group_cats.items())
                log(f"  Invoice (p{page_range}): {summary}")

        processed_bytes = overlay_tags_on_pdf(
            input_path, cats_for_overlay, highlights_per_page, config
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


# ── Category Editor ───────────────────────────────────────────────────────────
class CategoryEditor(QDialog):
    """Dialog for managing categories and their keywords."""

    PURPLE = "#8C4CAF"
    RED    = "#AA3333"
    GREEN  = "#2E8B57"
    GREY   = "#666666"

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Categories")
        self.resize(820, 540)
        self.setModal(True)

        self.config_data = load_config()
        self.categories  = self.config_data["categories"]
        self._selected   = None

        self._build()
        self._refresh_cat_list()
        if self.categories:
            self.cat_list.setCurrentRow(0)

    # ── Layout ────────────────────────────────────────────────────────────────
    def _build(self):
        def sbtn(text, color, width=None):
            b = QPushButton(text)
            b.setStyleSheet(
                f"QPushButton{{background:{color};color:white;font-weight:bold;"
                f"border:none;padding:5px 10px;border-radius:3px;}}"
                f"QPushButton:hover{{background:{_darken(color)};}}"
            )
            if width:
                b.setFixedWidth(width)
            return b

        root = QHBoxLayout(self)
        root.setContentsMargins(10, 10, 10, 10)

        # ── Left panel: category list ─────────────────────────────────────
        left = QWidget()
        left.setFixedWidth(200)
        ll = QVBoxLayout(left)
        ll.setContentsMargins(0, 0, 4, 0)

        hdr = QLabel("Categories")
        hdr.setStyleSheet("font-weight:bold; font-size:10pt;")
        ll.addWidget(hdr)

        self.cat_list = QListWidget()
        self.cat_list.setFont(QFont("Arial", 10))
        self.cat_list.currentRowChanged.connect(self._on_cat_select)
        ll.addWidget(self.cat_list)

        cat_btns = QHBoxLayout()
        add_btn = sbtn("+ Add",    self.PURPLE)
        del_btn = sbtn("- Delete", self.RED)
        add_btn.clicked.connect(self._add_category)
        del_btn.clicked.connect(self._delete_category)
        cat_btns.addWidget(add_btn)
        cat_btns.addWidget(del_btn)
        ll.addLayout(cat_btns)

        root.addWidget(left)

        # ── Right panel: details + keywords ──────────────────────────────
        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(4, 0, 0, 0)

        # Name / Code / Color row
        meta = QWidget()
        ml = QGridLayout(meta)
        ml.setContentsMargins(0, 0, 0, 6)

        for col, (label, attr, width) in enumerate([
            ("Name",    "_entry_name", 140),
            ("Code",    "_entry_code",  70),
            ("Color R", "_entry_r",     45),
            ("Color G", "_entry_g",     45),
            ("Color B", "_entry_b",     45),
        ]):
            lbl = QLabel(label)
            lbl.setStyleSheet("font-size:9pt;")
            ml.addWidget(lbl, 0, col)
            e = QLineEdit()
            e.setFixedWidth(width)
            e.setFont(QFont("Arial", 10))
            ml.addWidget(e, 1, col)
            setattr(self, attr, e)

        apply_btn = sbtn("Apply", self.PURPLE)
        apply_btn.clicked.connect(self._apply_meta)
        ml.addWidget(apply_btn, 1, 5)

        self._swatch = QLabel()
        self._swatch.setFixedSize(28, 24)
        self._swatch.setStyleSheet("border:1px solid #888; background:#b4b4b4;")
        ml.addWidget(self._swatch, 1, 6)

        for entry in (self._entry_r, self._entry_g, self._entry_b):
            entry.textChanged.connect(self._update_swatch)

        rl.addWidget(meta)

        kw_lbl = QLabel("Keywords  (one per line, case-insensitive)")
        kw_lbl.setStyleSheet("font-size:9pt;")
        rl.addWidget(kw_lbl)

        self.kw_list = QListWidget()
        self.kw_list.setFont(QFont("Courier New", 10))
        self.kw_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        rl.addWidget(self.kw_list)

        add_kw_row = QHBoxLayout()
        self._kw_entry = QLineEdit()
        self._kw_entry.setFont(QFont("Arial", 10))
        self._kw_entry.setPlaceholderText("new keyword…")
        self._kw_entry.returnPressed.connect(self._add_keyword)
        add_kw_row.addWidget(self._kw_entry)
        add_kw_btn = sbtn("+ Add Keyword",     self.PURPLE)
        rem_kw_btn = sbtn("- Remove Selected", self.RED)
        add_kw_btn.clicked.connect(self._add_keyword)
        rem_kw_btn.clicked.connect(self._remove_keywords)
        add_kw_row.addWidget(add_kw_btn)
        add_kw_row.addWidget(rem_kw_btn)
        rl.addLayout(add_kw_row)

        footer = QHBoxLayout()
        footer.addStretch()
        cancel_btn = sbtn("Cancel",       self.GREY,  width=90)
        save_btn   = sbtn("Save Changes", self.GREEN, width=120)
        cancel_btn.clicked.connect(self.reject)
        save_btn.clicked.connect(self._save)
        footer.addWidget(cancel_btn)
        footer.addWidget(save_btn)
        rl.addLayout(footer)

        root.addWidget(right)

    # ── Helpers ───────────────────────────────────────────────────────────────
    def _keywords(self):
        return [self.kw_list.item(i).text() for i in range(self.kw_list.count())]

    def _refresh_cat_list(self):
        self.cat_list.clear()
        for cat in self.categories:
            self.cat_list.addItem(f"{cat['code']}  {cat['name']}")

    def _on_cat_select(self, row=None):
        if row is None:
            row = self.cat_list.currentRow()
        if row < 0:
            return
        self._selected = row
        cat = self.categories[row]
        for entry, val in [
            (self._entry_name, cat["name"]),
            (self._entry_code, cat["code"]),
            (self._entry_r,    cat["color"][0]),
            (self._entry_g,    cat["color"][1]),
            (self._entry_b,    cat["color"][2]),
        ]:
            entry.setText(str(val))
        self._update_swatch()
        self.kw_list.clear()
        for kw in cat["keywords"]:
            self.kw_list.addItem(kw)

    def _update_swatch(self):
        try:
            r, g, b = int(self._entry_r.text()), int(self._entry_g.text()), int(self._entry_b.text())
            self._swatch.setStyleSheet(f"background-color:rgb({r},{g},{b}); border:1px solid #888;")
        except ValueError:
            pass

    def _apply_meta(self):
        if self._selected is None:
            return
        try:
            r, g, b = int(self._entry_r.text()), int(self._entry_g.text()), int(self._entry_b.text())
            if not all(0 <= v <= 255 for v in (r, g, b)):
                raise ValueError
        except ValueError:
            QMessageBox.critical(self, "Invalid Color", "R, G, B must be integers 0–255.")
            return
        cat = self.categories[self._selected]
        cat["name"]  = self._entry_name.text().strip()
        cat["code"]  = self._entry_code.text().strip()
        cat["color"] = [r, g, b]
        self._refresh_cat_list()
        self.cat_list.setCurrentRow(self._selected)

    def _add_keyword(self):
        if self._selected is None:
            return
        kw = self._kw_entry.text().strip().lower()
        if not kw:
            return
        if kw in self._keywords():
            QMessageBox.information(self, "Duplicate", f'"{kw}" is already in this category.')
            return
        self.kw_list.addItem(kw)
        self.categories[self._selected]["keywords"] = self._keywords()
        self._kw_entry.clear()

    def _remove_keywords(self):
        if self._selected is None:
            return
        for item in self.kw_list.selectedItems():
            self.kw_list.takeItem(self.kw_list.row(item))
        self.categories[self._selected]["keywords"] = self._keywords()

    def _add_category(self):
        new_cat = {"code": "00000", "name": "New Category",
                   "color": [180, 180, 180], "keywords": []}
        self.categories.append(new_cat)
        self._refresh_cat_list()
        idx = len(self.categories) - 1
        self.cat_list.setCurrentRow(idx)

    def _delete_category(self):
        if self._selected is None:
            return
        cat = self.categories[self._selected]
        reply = QMessageBox.question(
            self, "Delete Category",
            f'Delete category "{cat["name"]}" ({cat["code"]})?\nThis cannot be undone.',
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return
        self.categories.pop(self._selected)
        self._selected = None
        self._refresh_cat_list()
        self.kw_list.clear()
        for e in (self._entry_name, self._entry_code, self._entry_r, self._entry_g, self._entry_b):
            e.clear()
        if self.categories:
            self.cat_list.setCurrentRow(0)

    def _save(self):
        if self._selected is not None:
            self.categories[self._selected]["keywords"] = self._keywords()
        self.config_data["categories"] = self.categories
        try:
            with open(CONFIG_FILE, "w") as f:
                json.dump(self.config_data, f, indent=2)
            QMessageBox.information(self, "Saved", "categories.json updated successfully.")
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Save Error", str(e))


# ── Main Window ───────────────────────────────────────────────────────────────
class App(QMainWindow):

    PURPLE = "#8C4CAF"
    RED    = "#AA3333"

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Coffee Lab Invoice Processor")
        self.resize(760, 540)
        self._build()

    # ── UI construction ───────────────────────────────────────────────────────
    def _build(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(10, 8, 10, 4)

        def sbtn(text, color):
            b = QPushButton(text)
            b.setStyleSheet(
                f"QPushButton{{background:{color};color:white;font-weight:bold;"
                f"border:none;padding:6px 14px;border-radius:3px;font-size:10pt;}}"
                f"QPushButton:hover{{background:{_darken(color)};}}"
            )
            b.setMinimumHeight(36)
            return b

        # Button row
        btn_row = QHBoxLayout()
        btn_row.setSpacing(6)
        b_select = sbtn("Select File(s)...", self.PURPLE)
        b_inbox  = sbtn("Process Inbox",     self.PURPLE)
        b_config = sbtn("Edit Categories",   self.PURPLE)
        b_clear  = sbtn("Clear Master PDF",  self.PURPLE)
        b_debug  = sbtn("DEBUG: Dump Text",  self.RED)
        b_select.clicked.connect(self.select_files)
        b_inbox.clicked.connect(self.process_inbox)
        b_config.clicked.connect(self.open_config)
        b_clear.clicked.connect(self.clear_master)
        b_debug.clicked.connect(self.debug_dump_text)
        for b in (b_select, b_inbox, b_config, b_clear, b_debug):
            btn_row.addWidget(b)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        # Status bar
        self.status_label = QLabel("Ready.")
        self.status_label.setStyleSheet("color:#444; font-size:9pt; padding:2px;")
        layout.addWidget(self.status_label)

        # Log panel
        log_group = QGroupBox(" Log ")
        log_group.setStyleSheet("QGroupBox{font-size:9pt;}")
        log_inner = QVBoxLayout(log_group)
        log_inner.setContentsMargins(4, 4, 4, 4)
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setFont(QFont("Courier New", 9))
        self.log_box.setStyleSheet("background:#1e1e1e; color:#d4d4d4;")
        log_inner.addWidget(self.log_box)
        layout.addWidget(log_group)

        # Footer paths
        footer = QLabel(
            f"Inbox: {INBOX_DIR}    |    Processed: {PROCESSED_DIR}    |    Master: {MASTER_PDF}"
        )
        footer.setStyleSheet("color:#999; font-size:7pt; font-family:'Courier New';")
        layout.addWidget(footer)

    # ── Logging helpers ───────────────────────────────────────────────────────
    def log(self, msg: str):
        self.log_box.append(msg)
        self.log_box.moveCursor(QTextCursor.End)
        QApplication.processEvents()

    def status(self, msg: str):
        self.status_label.setText(msg)
        QApplication.processEvents()

    # ── Actions ───────────────────────────────────────────────────────────────
    def select_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select Invoice PDF(s)", "",
            "PDF Files (*.pdf);;All Files (*.*)"
        )
        if paths:
            self._run_batch([Path(p) for p in paths])

    def process_inbox(self):
        pdfs = sorted(INBOX_DIR.glob("*.pdf"))
        if not pdfs:
            QMessageBox.information(
                self, "Inbox Empty",
                f"No PDF files found in:\n{INBOX_DIR}\n\n"
                "Drop invoice PDFs into the inbox folder and try again."
            )
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
                    if dest.exists():
                        ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
                        dest = PROCESSED_DIR / f"{path.stem}_orig_{ts}.pdf"
                    shutil.move(str(path), str(dest))

        result = f"Done — {success}/{total} succeeded."
        self.log(f"\n{'─'*60}")
        self.log(result)
        self.status(result)

    def open_config(self):
        editor = CategoryEditor(self)
        editor.exec()

    def clear_master(self):
        if not MASTER_PDF.exists():
            QMessageBox.information(self, "Nothing to clear", "Master PDF does not exist yet.")
            return
        reply = QMessageBox.question(
            self, "Confirm",
            f"Delete master PDF?\n{MASTER_PDF}\n\nThis cannot be undone.",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            MASTER_PDF.unlink()
            self.log("Master PDF deleted.")
            self.status("Master PDF cleared.")

    def debug_dump_text(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select PDF to dump text", "",
            "PDF Files (*.pdf);;All Files (*.*)"
        )
        if not paths:
            return
        for path in paths:
            self.log(f"\n{'═'*60}")
            self.log(f"DEBUG TEXT DUMP: {Path(path).name}")
            self.log(f"{'═'*60}")
            try:
                with pdfplumber.open(path) as pdf:
                    for i, page in enumerate(pdf.pages):
                        text = page.extract_text() or ""
                        self.log(f"\n── Page {i+1} ──────────────────────────")
                        for lineno, line in enumerate(text.splitlines(), 1):
                            self.log(f"{lineno:3}: {line}")
            except Exception as e:
                self.log(f"ERROR: {e}")


# ── Entry point ───────────────────────────────────────────────────────────────
def main():
    ensure_dirs()
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = App()
    try:
        icon_path = BASE_DIR / "icon.ico"
        if icon_path.exists():
            window.setWindowIcon(QIcon(str(icon_path)))
    except Exception:
        pass
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
