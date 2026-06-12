import type { Config } from "./types";
import defaultConfigJson from "./defaultConfig.json";

const PROFILES_KEY = "invoice-processor-profiles-v1";
const LEGACY_CONFIG_KEY = "invoice-processor-config-v1";

export interface ProfileStore {
  active: string;
  profiles: Record<string, Config>;
}

export function defaultConfig(): Config {
  return structuredClone(defaultConfigJson as Config);
}

export function defaultStore(): ProfileStore {
  return { active: "Default", profiles: { Default: defaultConfig() } };
}

export function loadStore(): ProfileStore {
  try {
    const raw = localStorage.getItem(PROFILES_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (validateStore(parsed)) return parsed;
    }
    // One-time migration from the pre-profiles single-config key
    const legacy = localStorage.getItem(LEGACY_CONFIG_KEY);
    if (legacy) {
      const parsed = JSON.parse(legacy);
      if (validateConfig(parsed)) {
        const store: ProfileStore = {
          active: "Default",
          profiles: { Default: parsed },
        };
        saveStore(store);
        localStorage.removeItem(LEGACY_CONFIG_KEY);
        return store;
      }
    }
  } catch {
    // fall through to defaults
  }
  return defaultStore();
}

export function saveStore(store: ProfileStore): void {
  localStorage.setItem(PROFILES_KEY, JSON.stringify(store));
}

export function validateStore(data: unknown): data is ProfileStore {
  if (typeof data !== "object" || data === null) return false;
  const { active, profiles } = data as ProfileStore;
  if (typeof active !== "string") return false;
  if (typeof profiles !== "object" || profiles === null) return false;
  const entries = Object.values(profiles);
  if (entries.length === 0 || !(active in profiles)) return false;
  return entries.every(validateConfig);
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
