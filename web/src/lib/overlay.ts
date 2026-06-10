/**
 * PDF tagging and master-append via pdf-lib. Geometry and styling are a
 * direct port of create_tag_overlay() in invoice_processor.py.
 */
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import type { CategoryTotals, Config, Highlight } from "./types";

function buildColorMap(config: Config): Map<string, [number, number, number]> {
  const m = new Map<string, [number, number, number]>();
  for (const cat of config.categories) {
    m.set(`${cat.code} ${cat.name}`, cat.color);
  }
  return m;
}

const FALLBACK_COLOR: [number, number, number] = [180, 180, 180];

function fmtMoney(amount: number): string {
  return amount.toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

/**
 * Draw tag boxes (and amount highlights) onto each page of the invoice.
 * Returns the tagged PDF bytes.
 */
export async function overlayTagsOnPdf(
  inputBytes: Uint8Array,
  catsPerPage: CategoryTotals[],
  highlightsPerPage: Highlight[][],
  config: Config,
): Promise<Uint8Array> {
  const colorMap = buildColorMap(config);
  const doc = await PDFDocument.load(inputBytes);
  const font = await doc.embedFont(StandardFonts.HelveticaBold);

  doc.getPages().forEach((page, i) => {
    const cats = catsPerPage[i] ?? {};
    const highlights = highlightsPerPage[i] ?? [];
    const { height: pageH } = page.getSize();

    // ── Amount highlights ──────────────────────────────────────────────
    for (const hl of highlights) {
      const [r, g, b] = colorMap.get(hl.label) ?? FALLBACK_COLOR;
      page.drawRectangle({
        x: hl.x0,
        y: hl.yBottom,
        width: hl.x1 - hl.x0,
        height: hl.yTop - hl.yBottom,
        color: rgb(r / 255, g / 255, b / 255),
        opacity: 0.55,
      });
    }

    // ── Tag boxes ─────────────────────────────────────────────────────
    const TAG_W = 135;
    const TAG_H = 40;
    const TAG_X_START = 22;
    const TAG_Y_TOP = pageH - 44;
    const GAP = 6;

    let x = TAG_X_START;
    for (const [label, amount] of Object.entries(cats)) {
      const [cr, cg, cb] = colorMap.get(label) ?? FALLBACK_COLOR;
      const r = cr / 255, g = cg / 255, b = cb / 255;

      page.drawRectangle({
        x,
        y: TAG_Y_TOP - TAG_H,
        width: TAG_W,
        height: TAG_H,
        color: rgb(r, g, b),
        opacity: 0.72,
        borderColor: rgb(r * 0.65, g * 0.65, b * 0.65),
        borderOpacity: 0.85,
        borderWidth: 1.5,
      });
      page.drawText(label, {
        x: x + 6,
        y: TAG_Y_TOP - 15,
        size: 9,
        font,
        color: rgb(0, 0, 0),
      });
      page.drawText(`$${fmtMoney(amount)}`, {
        x: x + 6,
        y: TAG_Y_TOP - 30,
        size: 11,
        font,
        color: rgb(0, 0, 0),
      });

      x += TAG_W + GAP;
    }
  });

  return doc.save();
}

/**
 * Append tagged invoice pages to an existing master PDF (or start a new
 * master if none provided). Returns the updated master bytes.
 */
export async function appendToMaster(
  masterBytes: Uint8Array | null,
  taggedBytes: Uint8Array[],
): Promise<Uint8Array> {
  const out = await PDFDocument.create();
  if (masterBytes) {
    const master = await PDFDocument.load(masterBytes);
    const pages = await out.copyPages(master, master.getPageIndices());
    pages.forEach((p) => out.addPage(p));
  }
  for (const bytes of taggedBytes) {
    const doc = await PDFDocument.load(bytes);
    const pages = await out.copyPages(doc, doc.getPageIndices());
    pages.forEach((p) => out.addPage(p));
  }
  return out.save();
}
