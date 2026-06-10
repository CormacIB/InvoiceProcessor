import type { Config } from "./types";
import defaultConfigJson from "./defaultConfig.json";

const STORAGE_KEY = "invoice-processor-config-v1";

export function defaultConfig(): Config {
  return structuredClone(defaultConfigJson as Config);
}

export function loadConfig(): Config {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (validateConfig(parsed)) return parsed;
    }
  } catch {
    // fall through to defaults
  }
  return defaultConfig();
}

export function saveConfig(config: Config): void {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(config));
}

export function resetConfig(): Config {
  localStorage.removeItem(STORAGE_KEY);
  return defaultConfig();
}

export function validateConfig(data: unknown): data is Config {
  if (typeof data !== "object" || data === null) return false;
  const cats = (data as Config).categories;
  if (!Array.isArray(cats) || cats.length === 0) return false;
  return cats.every(
    (c) =>
      typeof c.code === "string" &&
      typeof c.name === "string" &&
      Array.isArray(c.color) &&
      c.color.length === 3 &&
      c.color.every((v) => typeof v === "number" && v >= 0 && v <= 255) &&
      Array.isArray(c.keywords) &&
      c.keywords.every((k) => typeof k === "string"),
  );
}

export function exportConfig(config: Config): string {
  return JSON.stringify(config, null, 2);
}

export function importConfig(json: string): Config {
  const parsed = JSON.parse(json);
  if (!validateConfig(parsed)) {
    throw new Error(
      "Invalid config: expected { categories: [{ code, name, color: [r,g,b], keywords: [...] }] }",
    );
  }
  return parsed;
}
