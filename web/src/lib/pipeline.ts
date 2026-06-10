/**
 * Pure pipeline logic ported from invoice_processor.py (the reference
 * implementation). Function-for-function, regex-for-regex — if you change
 * behaviour here, change it there too or the golden tests will drift.
 */
import type {
  CategoryTotals,
  Config,
  Highlight,
  MatchedItem,
  Vendor,
  Word,
} from "./types";

export function round2(v: number): number {
  return Math.round((v + Number.EPSILON) * 100) / 100;
}

// ── Vendor detection ──────────────────────────────────────────────────────────
export function detectVendor(text: string): Vendor {
  const t = text.toUpperCase();
  if (t.includes("SYSCO")) return "sysco";
  if (t.includes("INNERMOUNTAIN")) return "innermountain";
  if (t.includes("ITALCO")) return "italco";
  if (t.includes("CRESTED BUCHA")) return "crested_bucha";
  if (t.includes("SISU STUDIOS")) return "sisu";
  if (t.includes("VERMONT STICKY")) return "vermont_sticky";
  if (t.includes("GUNNISON COUNTY")) return "skip"; // license/permit invoices — no tagging
  return "generic";
}

// ── Sysco: extract pre-printed category codes + amounts ───────────────────────
export function extractSyscoCategories(text: string): CategoryTotals {
  const cats: CategoryTotals = {};
  // Allow an optional intervening line (e.g. a date) between code and amount
  const pat = /(\d{5})\s+([A-Za-z&/ ]+?)\s*\n[^\n]*\n\s*\$([\d,]+\.\d{2})/g;
  for (const m of text.matchAll(pat)) {
    const label = `${m[1].trim()} ${m[2].trim()}`;
    cats[label] = (cats[label] ?? 0) + parseFloat(m[3].replace(/,/g, ""));
  }
  if (Object.keys(cats).length === 0) {
    // Fallback: code+name then $ directly on next line (no date)
    const pat2 = /(\d{5})\s+([A-Za-z&/ ]+?)\s*\n\s*\$([\d,]+\.\d{2})/g;
    for (const m of text.matchAll(pat2)) {
      const label = `${m[1].trim()} ${m[2].trim()}`;
      cats[label] = (cats[label] ?? 0) + parseFloat(m[3].replace(/,/g, ""));
    }
  }
  return cats;
}

// ── Generic: extract (description, amount) line items ────────────────────────
const SKIP_WORDS = [
  "total", "subtotal", "tax", "balance", "payment", "due",
  "amount", "price", "extended", "invoice",
  "misc", "page", "terms", "group total", "order summary",
  "remit", "cases", "split", "cube", "gross",
  "sysco", "confidential", "paca", "driver", "sign",
  "important", "authorized", "retains", "receivables", "proceeds",
  "dispute", "representative", "capacity", "claimants",
  "open:", "close:", "5:00 am", "9:00 pm", // Sysco footer time strings
  "misc charges", "misc tax",
];

export function extractLineItems(text: string): [string, number][] {
  const items: [string, number][] = [];
  const lineRe = /^(.+?)\s+\$?([\d,]{0,7}\.\d{2})\s*[A-Za-z*]?\s*$/gm;
  for (const m of text.matchAll(lineRe)) {
    const desc = m[1].trim();
    const amount = parseFloat(m[2].replace(/,/g, ""));

    // Skip zero-dollar lines and unreasonably large amounts
    if (amount <= 0 || amount > 49999) continue;
    // Skip lines that are clearly totals/headers/footers
    const descLow = desc.toLowerCase();
    if (SKIP_WORDS.some((w) => descLow.includes(w))) {
      if (!descLow.includes("surcharge") && !descLow.includes("retail delivery fee")) {
        continue;
      }
    }
    // Skip Sysco lines marked OUT (not delivered, no extended price charged)
    if (/^(?:[a-z]\s+)?out\s/.test(descLow)) continue;
    // Also skip Sysco footer lines like "OPEN: 5:00 AM  CLOSE: 9:00 PM"
    if (/\d+:\d{2}\s*(am|pm)/.test(descLow)) continue;
    // Skip very short descriptions (likely column headers)
    if (desc.length < 4) continue;
    // Skip descriptions with no letters — bare item codes / symbol lines
    if (!/[a-zA-Z]/.test(desc)) continue;

    items.push([desc, amount]);
  }
  return items;
}

// ── Keyword categorisation ────────────────────────────────────────────────────
export function categorizeItems(
  items: [string, number][],
  config: Config,
): { totals: CategoryTotals; matched: MatchedItem[] } {
  const totals: CategoryTotals = {};
  const matched: MatchedItem[] = [];
  let uncategorized = 0;

  for (const [desc, amount] of items) {
    const descLow = desc.toLowerCase();
    let hit = false;
    for (const cat of config.categories) {
      if (cat.keywords.some((kw) => descLow.includes(kw.toLowerCase()))) {
        const label = `${cat.code} ${cat.name}`;
        totals[label] = (totals[label] ?? 0) + amount;
        matched.push({ desc, amount, label });
        hit = true;
        break;
      }
    }
    if (!hit) uncategorized += amount;
  }

  if (Object.keys(totals).length === 0 && uncategorized > 0) {
    totals["REVIEW"] = round2(uncategorized);
  }
  for (const k of Object.keys(totals)) totals[k] = round2(totals[k]);
  return { totals, matched };
}

export function extractInvoiceTotal(text: string): number {
  const patterns = [
    /invoice\s+total\s*\$?\s*([\d,]+\.\d{2})/i,
    /total\s+due\s*\$?\s*([\d,]+\.\d{2})/i,
    /total\s+sales\s*\$?\s*([\d,]+\.\d{2})/i,
    /net\s+amount\s+due\s*\$?\s*([\d,]+\.\d{2})/i,
    /balance\s*\$?\s*([\d,]+\.\d{2})/i,
  ];
  for (const pat of patterns) {
    const m = text.match(pat);
    if (m) return parseFloat(m[1].replace(/,/g, ""));
  }
  return 0;
}

// ── Per-page category dispatch ────────────────────────────────────────────────
export function adjustForSurcharges(cats: CategoryTotals, text: string): CategoryTotals {
  // If the invoice total > sum of categorized items (e.g. fuel surcharge,
  // small taxes), add the difference to the largest category.
  const invoiceTotal = extractInvoiceTotal(text);
  if (invoiceTotal <= 0) return cats;
  const itemSum = Object.values(cats).reduce((a, b) => a + b, 0);
  const diff = round2(invoiceTotal - itemSum);
  if (diff > 0 && diff <= 30) {
    const largest = Object.keys(cats).reduce((a, b) => (cats[a] >= cats[b] ? a : b));
    cats[largest] = round2(cats[largest] + diff);
  }
  return cats;
}

export function getPageCategories(
  text: string,
  config: Config,
  vendor: Vendor,
): { totals: CategoryTotals; matched: MatchedItem[] } {
  if (vendor === "skip") return { totals: {}, matched: [] };

  if (vendor === "sysco") {
    const isDelivery = text.toUpperCase().includes("DELIVERY COPY");
    if (!isDelivery) {
      const cats = extractSyscoCategories(text);
      if (Object.keys(cats).length > 0) {
        return { totals: cats, matched: [] }; // regex path — no per-line positions
      }
    }
    const tUp = text.toUpperCase();
    if (
      tUp.includes("ORDER SUMMARY") &&
      !/\b(DAIRY|FROZEN|CANNED|PAPER|CHEMICAL)\b/.test(tUp)
    ) {
      if (!tUp.includes("FUEL SURCHARGE")) return { totals: {}, matched: [] };
    }
  }

  // All other vendors: keyword match on line items
  const items = extractLineItems(text);
  if (items.length > 0) {
    const { totals, matched } = categorizeItems(items, config);
    if (Object.keys(totals).length > 0) {
      return { totals: adjustForSurcharges(totals, text), matched };
    }
  }

  // Last resort: whole-invoice keyword scan using total
  const total = extractInvoiceTotal(text);
  if (total > 0) {
    const { totals, matched } = categorizeItems(
      [[text.slice(0, 800).toLowerCase(), total]],
      config,
    );
    if (Object.keys(totals).length > 0) return { totals, matched };
    return { totals: { REVIEW: round2(total) }, matched: [] };
  }

  return { totals: {}, matched: [] };
}

// ── Amount position lookup ─────────────────────────────────────────────────────
const NUMERIC_RE = /^[+-]?(\d+\.?\d*|\.\d+)$/;

export function findAmountPositions(words: Word[], matchedItems: MatchedItem[]): Highlight[] {
  if (matchedItems.length === 0) return [];

  // amount value -> word boxes, rightmost column only, sorted top-to-bottom
  const wordMap = new Map<number, Word[]>();
  for (const w of words) {
    const cleaned = w.text
      .replace(/^\$+/, "")
      .replace(/,/g, "")
      .replace(/[A-Za-z*]+$/, "");
    if (!NUMERIC_RE.test(cleaned)) continue;
    const val = round2(parseFloat(cleaned));
    if (!wordMap.has(val)) wordMap.set(val, []);
    wordMap.get(val)!.push(w);
  }

  const rightmostMap = new Map<number, Word[]>();
  for (const [val, wlist] of wordMap) {
    const maxX1 = Math.max(...wlist.map((w) => w.x1));
    const col = wlist.filter((w) => Math.abs(w.x1 - maxX1) < 10);
    col.sort((a, b) => b.yTop - a.yTop); // top of page first
    rightmostMap.set(val, col);
  }

  // Items sharing an amount get positions assigned top-to-bottom in item order
  const used = new Map<number, number>();
  const highlights: Highlight[] = [];
  for (const { amount, label } of matchedItems) {
    const key = round2(amount);
    const col = rightmostMap.get(key) ?? [];
    const idx = used.get(key) ?? 0;
    if (idx >= col.length) continue;
    const best = col[idx];
    used.set(key, idx + 1);

    const pad = 2;
    highlights.push({
      x0: best.x0 - pad,
      yBottom: best.yBottom - pad,
      x1: best.x1 + pad,
      yTop: best.yTop + pad,
      label,
    });
  }
  return highlights;
}

// ── Invoice boundary detection ────────────────────────────────────────────────
export function findInvoiceGroups(pagesText: string[], vendor: Vendor): [number, number][] {
  // Groups consecutive pages belonging to the same invoice; end is exclusive.
  if (pagesText.length <= 1) return [[0, pagesText.length]];

  const groups: [number, number][] = [];
  let groupStart = 0;

  for (let i = 1; i < pagesText.length; i++) {
    const text = pagesText[i];
    if (vendor === "sysco") {
      // A new invoice begins on any non-delivery-copy page with pre-printed
      // category codes; delivery-copy pages trail their summary page.
      const isDelivery = text.toUpperCase().includes("DELIVERY COPY");
      if (!isDelivery && Object.keys(extractSyscoCategories(text)).length > 0) {
        groups.push([groupStart, i]);
        groupStart = i;
      }
    } else {
      if (/page\s+1\s+of\s+\d+/i.test(text)) {
        groups.push([groupStart, i]);
        groupStart = i;
      }
    }
  }

  groups.push([groupStart, pagesText.length]);
  return groups;
}
