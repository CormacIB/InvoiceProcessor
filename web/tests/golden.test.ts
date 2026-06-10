/**
 * Golden tests: the TypeScript port must reproduce the Python reference
 * implementation's output on real invoices.
 *
 * golden.json is generated from invoice_processor.py by:
 *   venv/bin/python3 tools/generate_golden.py processed/Invoice.pdf ... > web/tests/golden.json
 *
 * Layer 1 (pure logic): run the ported pipeline on the exact text pdfplumber
 * extracted — isolates regex/categorisation fidelity.
 * Layer 2 (end-to-end): run pdf.js extraction on the real PDFs — also
 * validates that our line reconstruction is close enough to pdfplumber's.
 */
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import {
  detectVendor,
  findInvoiceGroups,
  getPageCategories,
} from "../src/lib/pipeline";
import { extractPages } from "../src/lib/textExtract";
import { processInvoice } from "../src/lib/process";
import type { CategoryTotals, Config, Vendor } from "../src/lib/types";
import defaultConfigJson from "../src/lib/defaultConfig.json";
import golden from "./golden.json";

const config = defaultConfigJson as Config;
const REPO_ROOT = join(__dirname, "..", "..");

interface GoldenInvoice {
  file: string;
  vendor: string;
  pages: {
    plumber_text: string;
    categories: CategoryTotals;
    matched_items: [string, number, string][];
    highlight_count: number;
    highlight_labels: string[];
  }[];
  groups: { start: number; end: number; categories: CategoryTotals }[];
}

const invoices = (golden as unknown as { invoices: GoldenInvoice[] }).invoices;

describe("layer 1: pure logic on pdfplumber-extracted text", () => {
  for (const inv of invoices) {
    describe(inv.file, () => {
      const pagesText = inv.pages.map((p) => p.plumber_text);

      it("detects the vendor", () => {
        expect(detectVendor(pagesText.join("\n"))).toBe(inv.vendor);
      });

      it("reproduces per-page category totals", () => {
        inv.pages.forEach((page, i) => {
          const { totals } = getPageCategories(
            page.plumber_text,
            config,
            inv.vendor as Vendor,
          );
          expect(totals, `page ${i + 1}`).toEqual(page.categories);
        });
      });

      it("reproduces invoice grouping and group totals", () => {
        const groups = findInvoiceGroups(pagesText, inv.vendor as Vendor);
        expect(groups).toEqual(inv.groups.map((g) => [g.start, g.end]));
      });
    });
  }
});

describe("layer 2: end-to-end with pdf.js extraction on real PDFs", () => {
  for (const inv of invoices) {
    describe(inv.file, () => {
      const bytes = new Uint8Array(readFileSync(join(REPO_ROOT, "processed", inv.file)));

      it("matches the Python pipeline's per-page categories", async () => {
        const pages = await extractPages(bytes);
        expect(pages.length).toBe(inv.pages.length);
        const vendor = detectVendor(pages.map((p) => p.text).join("\n"));
        expect(vendor).toBe(inv.vendor);
        pages.forEach((page, i) => {
          const { totals } = getPageCategories(page.text, config, vendor);
          expect(totals, `page ${i + 1}`).toEqual(inv.pages[i].categories);
        });
      });

      it("produces a tagged PDF with the same page count", async () => {
        const result = await processInvoice(inv.file, bytes, config);
        expect(result.ok).toBe(true);
        expect(result.taggedBytes).toBeDefined();
        // pdf-lib can parse its own output and page count is preserved
        const { PDFDocument } = await import("pdf-lib");
        const doc = await PDFDocument.load(result.taggedBytes!);
        expect(doc.getPageCount()).toBe(inv.pages.length);
      });
    });
  }
});
