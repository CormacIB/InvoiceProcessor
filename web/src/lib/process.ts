/**
 * Orchestrator — port of process_invoice() in invoice_processor.py, minus
 * the filesystem: takes invoice bytes, returns tagged bytes + a log.
 */
import type { CategoryTotals, Config, Highlight, Vendor } from "./types";
import { extractPages } from "./textExtract";
import {
  detectVendor,
  findAmountPositions,
  findInvoiceGroups,
  getPageCategories,
  round2,
} from "./pipeline";
import { overlayTagsOnPdf } from "./overlay";

export interface ProcessResult {
  name: string;
  ok: boolean;
  vendor?: Vendor;
  taggedBytes?: Uint8Array;
  taggedName?: string;
  /** Per-page category totals, for display. */
  pages?: CategoryTotals[];
  /** Whole-file category totals (all invoice groups aggregated). */
  totals?: CategoryTotals;
  log: string[];
}

function fmtCats(cats: CategoryTotals): string {
  return Object.entries(cats)
    .map(([k, v]) => `${k}  $${v.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`)
    .join("  |  ");
}

function timestamp(): string {
  const d = new Date();
  const p = (n: number, w = 2) => String(n).padStart(w, "0");
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
}

export async function processInvoice(
  name: string,
  bytes: Uint8Array,
  config: Config,
): Promise<ProcessResult> {
  const log: string[] = [`> ${name}`];
  try {
    const pages = await extractPages(bytes);

    if (pages.length === 0 || pages.every((p) => !p.text)) {
      log.push("  ⚠  No text could be extracted — is this a scanned image PDF?");
      log.push("     (Scanned/photographed invoices aren't supported yet — the PDF needs selectable text.)");
      return { name, ok: false, log };
    }

    const fullText = pages.map((p) => p.text).join("\n");
    const vendor = detectVendor(fullText);
    log.push(`  Vendor: ${vendor}`);

    const catsPerPage: CategoryTotals[] = [];
    const highlightsPerPage: Highlight[][] = [];

    pages.forEach((page, i) => {
      const { totals, matched } = getPageCategories(page.text, config, vendor);
      catsPerPage.push(totals);
      highlightsPerPage.push(findAmountPositions(page.words, matched));
      log.push(
        Object.keys(totals).length > 0
          ? `  Page ${i + 1}: ${fmtCats(totals)}`
          : `  Page ${i + 1}: (no tag)`,
      );
    });

    if (catsPerPage.every((c) => Object.keys(c).length === 0)) {
      log.push("  ⚠  No categories found — check your category keywords.");
    }

    // Aggregate each invoice group's categories onto its first page only
    const groups = findInvoiceGroups(pages.map((p) => p.text), vendor);
    const catsForOverlay: CategoryTotals[] = catsPerPage.map(() => ({}));
    for (const [start, end] of groups) {
      const groupCats: CategoryTotals = {};
      for (const cats of catsPerPage.slice(start, end)) {
        for (const [label, amount] of Object.entries(cats)) {
          groupCats[label] = round2((groupCats[label] ?? 0) + amount);
        }
      }
      catsForOverlay[start] = groupCats;
      if (Object.keys(groupCats).length > 0 && end - start > 1) {
        log.push(`  Invoice (p${start + 1}–${end}): ${fmtCats(groupCats)}`);
      }
    }

    const taggedBytes = await overlayTagsOnPdf(bytes, catsForOverlay, highlightsPerPage, config);
    const stem = name.replace(/\.pdf$/i, "");
    const taggedName = `${stem}_tagged_${timestamp()}.pdf`;
    log.push(`  Tagged: ${taggedName}`);

    const totals: CategoryTotals = {};
    for (const cats of catsPerPage) {
      for (const [label, amount] of Object.entries(cats)) {
        totals[label] = round2((totals[label] ?? 0) + amount);
      }
    }

    return {
      name,
      ok: true,
      vendor,
      taggedBytes,
      taggedName,
      pages: catsPerPage,
      totals,
      log,
    };
  } catch (e) {
    log.push(`  ERROR: ${e instanceof Error ? e.message : String(e)}`);
    return { name, ok: false, log };
  }
}
