/**
 * Tercihler: per-document preferred spellings, and export/import of the
 * document's Vefeyât preferences (VBA 0.3.41 `VfpPrefer`, `VfpExport`).
 */
import { type Person, isValidAlias, nameQuery, personKey } from "../core/index.js";
import { type Settings, SETTING_CHOICES, type SettingName, parseConfirmations, parseExclusions } from "./logic.js";
import { type LocalPerson, parseLocalPersons, upsertLocal } from "./persons.js";

/** person key -> preferred spelling (without a Turkish ending). */
export type Spellings = Readonly<Record<string, string>>;

export function parseSpellings(stored: unknown): Record<string, string> {
  const map = parseConfirmations(stored); // same JSON object shape
  return Object.fromEntries(Object.entries(map).filter(([, v]) => isValidAlias(v)));
}

export function preferredSpelling(spellings: Spellings, person: Person): string | null {
  return spellings[personKey(person)] ?? null;
}

/** The selected literal as a preference: the name without its ending. */
export function preferenceOf(literal: string): string {
  return nameQuery(literal);
}

/** Spelling options with the document preference first. */
export function withPreferred<T extends { value: string; label: string }>(options: readonly T[], preferred: string | null): Array<T | { value: string; label: string }> {
  if (!preferred) return [...options];
  return [{ value: preferred, label: "belge tercihi" }, ...options.filter((o) => o.value !== preferred)];
}

export interface PrefsFile {
  readonly format: "vefeyat-tercihler";
  readonly version: 1;
  readonly settings: Partial<Settings>;
  readonly spellings: Record<string, string>;
  readonly exclusions: string[];
  readonly confirmations: Record<string, string>;
  /** Kendi şahıslarım (absent in files from older versions). */
  readonly persons: LocalPerson[];
}

export function exportPrefs(
  settings: Settings,
  spellings: Spellings,
  exclusions: readonly string[],
  confirmations: Readonly<Record<string, string>>,
  persons: readonly LocalPerson[] = [],
): string {
  const file: PrefsFile = {
    format: "vefeyat-tercihler",
    version: 1,
    settings: { ...settings },
    spellings: { ...spellings },
    exclusions: [...exclusions],
    confirmations: { ...confirmations },
    persons: [...persons],
  };
  return JSON.stringify(file, null, 2);
}

export type ImportResult = { ok: true; file: PrefsFile } | { ok: false; reason: string };

export function parsePrefs(text: string): ImportResult {
  let value: unknown;
  try {
    value = JSON.parse(text);
  } catch {
    return { ok: false, reason: "Dosya okunamadı: geçerli bir JSON değil." };
  }
  const raw = value as Partial<PrefsFile> | null;
  if (!raw || raw.format !== "vefeyat-tercihler") return { ok: false, reason: "Bu bir Vefeyât tercih dosyası değil." };
  if (raw.version !== 1) return { ok: false, reason: "Tercih dosyasının sürümü desteklenmiyor." };
  const settings: Partial<Record<SettingName, string>> = {};
  for (const name of Object.keys(SETTING_CHOICES) as SettingName[]) {
    const v = (raw.settings as Record<string, unknown> | undefined)?.[name];
    if (typeof v === "string" && v in SETTING_CHOICES[name]) settings[name] = v;
  }
  return {
    ok: true,
    file: {
      format: "vefeyat-tercihler",
      version: 1,
      settings: settings as Partial<Settings>,
      spellings: parseSpellings(JSON.stringify(raw.spellings ?? {})),
      exclusions: parseExclusions(JSON.stringify(raw.exclusions ?? [])),
      confirmations: parseConfirmations(JSON.stringify(raw.confirmations ?? {})),
      persons: parseLocalPersons(JSON.stringify(raw.persons ?? [])),
    },
  };
}

/** Import merges lists into the document's own; imported values win. */
export function mergePrefs(
  current: { spellings: Spellings; exclusions: readonly string[]; confirmations: Readonly<Record<string, string>>; persons?: readonly LocalPerson[] },
  file: PrefsFile,
): { spellings: Record<string, string>; exclusions: string[]; confirmations: Record<string, string>; persons: LocalPerson[] } {
  return {
    spellings: { ...current.spellings, ...file.spellings },
    exclusions: [...new Set([...current.exclusions, ...file.exclusions])],
    confirmations: { ...current.confirmations, ...file.confirmations },
    persons: file.persons.reduce((list, person) => upsertLocal(list, person), [...(current.persons ?? [])]),
  };
}
