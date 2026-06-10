/**
 * PDF text extraction via pdf.js, shaped to mimic pdfplumber's
 * extract_text()/extract_words() output that the pipeline regexes were
 * written against: lines sorted top-to-bottom, words left-to-right,
 * single spaces between words.
 */
import * as pdfjsLib from "pdfjs-dist";
import type { ExtractedPage, Word } from "./types";

// pdfplumber's default y_tolerance: words whose baselines are within this
// many points are considered the same line.
const Y_TOLERANCE = 3;

interface RawTextItem {
  str: string;
  transform: number[];
  width: number;
  height: number;
}

function itemToWords(item: RawTextItem): Word[] {
  const str = item.str;
  if (!str.trim()) return [];
  const x = item.transform[4];
  const y = item.transform[5]; // baseline, origin bottom-left
  const h = item.height || Math.hypot(item.transform[1], item.transform[3]);
  const charW = str.length > 0 ? item.width / str.length : 0;

  const words: Word[] = [];
  const re = /\S+/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(str)) !== null) {
    words.push({
      text: m[0],
      x0: x + m.index * charW,
      x1: x + (m.index + m[0].length) * charW,
      // approximate descent/ascent around the baseline
      yBottom: y - 0.2 * h,
      yTop: y + 0.8 * h,
    });
  }
  return words;
}

function wordsToText(words: Word[]): string {
  if (words.length === 0) return "";
  // Cluster words into lines by baseline proximity, top of page first
  const sorted = [...words].sort((a, b) => b.yBottom - a.yBottom);
  const lines: Word[][] = [];
  let current: Word[] = [sorted[0]];
  let lineY = sorted[0].yBottom;
  for (let i = 1; i < sorted.length; i++) {
    const w = sorted[i];
    if (Math.abs(w.yBottom - lineY) <= Y_TOLERANCE) {
      current.push(w);
    } else {
      lines.push(current);
      current = [w];
      lineY = w.yBottom;
    }
  }
  lines.push(current);

  return lines
    .map((line) =>
      line
        .sort((a, b) => a.x0 - b.x0)
        .map((w) => w.text)
        .join(" "),
    )
    .join("\n");
}

/**
 * True when the text runs horizontally left-to-right. Rotated text (e.g.
 * Sysco's sideways "EQUAL OPPORTUNITY..." margin boilerplate) shares
 * baselines with table rows and would corrupt line reconstruction, so it
 * goes on lines of its own — pdfplumber does the same for non-upright text.
 */
function isUpright(item: RawTextItem): boolean {
  const [, b, c] = item.transform;
  return Math.abs(b) < 0.001 && Math.abs(c) < 0.001;
}

export async function extractPages(data: Uint8Array): Promise<ExtractedPage[]> {
  // pdf.js transfers (detaches) the buffer it's given — pass a copy so the
  // caller's bytes stay usable for the pdf-lib overlay step afterwards.
  const doc = await pdfjsLib.getDocument({ data: data.slice() }).promise;
  const pages: ExtractedPage[] = [];
  try {
    for (let p = 1; p <= doc.numPages; p++) {
      const page = await doc.getPage(p);
      const viewport = page.getViewport({ scale: 1 });
      const content = await page.getTextContent();
      const words: Word[] = [];
      const rotatedLines: string[] = [];
      for (const item of content.items as unknown as RawTextItem[]) {
        if (!("str" in item)) continue;
        if (isUpright(item)) {
          words.push(...itemToWords(item));
        } else if (item.str.trim()) {
          rotatedLines.push(item.str.trim());
        }
      }
      const text = [wordsToText(words), ...rotatedLines].filter(Boolean).join("\n");
      pages.push({
        text,
        words,
        width: viewport.width,
        height: viewport.height,
      });
    }
  } finally {
    await doc.destroy();
  }
  return pages;
}
