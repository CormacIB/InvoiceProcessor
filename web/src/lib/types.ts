export interface Category {
  code: string;
  name: string;
  color: [number, number, number];
  keywords: string[];
}

export interface Config {
  categories: Category[];
}

/** {"50900 F&B": 169.43, ...} */
export type CategoryTotals = Record<string, number>;

export interface MatchedItem {
  desc: string;
  amount: number;
  label: string;
}

/** A word on the page, in PDF coordinates (origin bottom-left). */
export interface Word {
  text: string;
  x0: number;
  x1: number;
  yBottom: number;
  yTop: number;
}

export interface ExtractedPage {
  text: string;
  words: Word[];
  width: number;
  height: number;
}

/** Highlight box in PDF coordinates (origin bottom-left). */
export interface Highlight {
  x0: number;
  yBottom: number;
  x1: number;
  yTop: number;
  label: string;
}

export type Vendor =
  | "sysco"
  | "innermountain"
  | "italco"
  | "crested_bucha"
  | "sisu"
  | "vermont_sticky"
  | "skip"
  | "generic";
