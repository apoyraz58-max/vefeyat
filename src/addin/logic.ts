/**
 * Task pane decisions that do not need Word. Word calls live in word.ts so
 * that everything here is unit tested.
 */
import {
  type Catalog,
  type Person,
  type PrefixKey,
  alreadyHasDate,
  confirmationCandidates,
  formatDate,
  isValidAlias,
  isValidDate,
  nameCandidates,
  nameQuery,
  normalize,
  personKey,
  singleSafe,
  startsWithDate,
  strength,
  textName,
} from "../core/index.js";

export interface Lookup {
  /** Selected name without a Turkish ending. */
  readonly query: string;
  /** Strong candidates (same name), then partial ones. */
  readonly candidates: readonly Person[];
  readonly strong: ReadonlySet<string>;
  /** Best name score per person key: 4/3 same name, 2 part of the name, 1 similar spelling. */
  readonly scores: ReadonlyMap<string, number>;
  /** Set when exactly one strong candidate exists: no choice needed. */
  readonly automatic: Person | null;
}

/**
 * Candidates for a selected name, as the VBA "Ölüm tarihi ekle" command
 * gathers them: name matches only (no description matches), one entry per
 * person, strong matches first.
 */
export function lookup(catalog: Catalog, alias: string): Lookup {
  const found = confirmationCandidates(catalog.find(alias, false), alias);
  const { exact, partial } = nameCandidates(found, alias);
  const strong = new Set(exact.map(personKey));
  return {
    query: nameQuery(alias),
    candidates: [...exact, ...partial],
    strong,
    scores: new Map(found.map((person) => [personKey(person), strength(person, alias)])),
    automatic: singleSafe(found, alias) ? (found[0] as Person) : null,
  };
}

export type DatePlan =
  | { readonly ok: true; readonly text: string }
  | { readonly ok: false; readonly reason: string };

/**
 * Whether a date may be written after the name, and its text. `following` is
 * the rest of the paragraph after the selected name.
 */
export function planDate(person: Person, prefix: PrefixKey, following: string): DatePlan {
  if (!isValidDate(person.deathDate)) {
    return { ok: false, reason: "Katalogda bu şahıs için kullanılabilir bir ölüm tarihi yok; tarih eklenmedi." };
  }
  if (startsWithDate(following) || alreadyHasDate(following, person.deathDate)) {
    return { ok: false, reason: "Adın hemen arkasında zaten bir tarih var; mevcut tarihe dokunulmadı." };
  }
  return { ok: true, text: formatDate(person.deathDate, prefix) };
}

/**
 * Literal spellings searched before the selection to warn about an earlier
 * mention: the selected spelling, the three catalog names, their reading
 * order and the family part of the short name ("Râzî" of "Fahreddin er-Râzî").
 * Very short or invalid values are skipped; at most 12 spellings.
 */
export function earlierSpellings(person: Person, alias: string): string[] {
  const result: string[] = [];
  const seen = new Set<string>();
  const add = (value: string): void => {
    value = value.trim();
    const key = normalize(value);
    if (value.length < 3 || value.length > 200 || key.length < 3 || !isValidAlias(value) || seen.has(value)) return;
    seen.add(value);
    result.push(value);
  };
  add(nameQuery(alias));
  for (const name of [person.shortName, person.fullName, person.originalName]) {
    add(name.includes(",") ? textName(name) : name);
  }
  const short = person.shortName.includes(",")
    ? person.shortName.slice(0, person.shortName.indexOf(","))
    : (person.shortName.trim().split(/[\s\-–]+/).pop() ?? "");
  add(short);
  return result.slice(0, 12);
}

/** Content control tag of a Vefeyât date: short, stable per person. */
export function dateTag(person: Person): string {
  return `vefeyat:date:${hash(personKey(person))}`;
}

function hash(value: string): string {
  // FNV-1a 32-bit; only used to group dates of one person, not for security.
  let h = 0x811c9dc5;
  for (let i = 0; i < value.length; i++) {
    h ^= value.charCodeAt(i);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return h.toString(16).padStart(8, "0");
}

/** Short note explaining a non-exact candidate, or "" for the same name. */
export function matchHint(score: number): string {
  if (score >= 3) return "";
  if (score === 2) return "adın bir parçası";
  return "benzer yazım, doğrulayın";
}

/** One-line summary of a candidate for the list. */
export function candidateLabel(person: Person): string {
  const name = textName(person.shortName);
  return person.deathDate ? `${name} ${person.deathDate}` : name;
}

export type View = "secili" | "dokum" | "teslim" | "ayarlar";
export type Tool = "tutarlilik" | "rapor" | "dizin" | "temizlik";

export interface Launch {
  readonly view: View;
  readonly action: "insert" | "spelling" | null;
  readonly tool: Tool | null;
}

/**
 * Which view and action a ribbon or context-menu button asked for, from the
 * task pane URL query (?action=insert, ?view=dokum, ?tool=dizin …).
 * Unknown values fall back to the selected-name view.
 */
export function parseLaunch(search: string): Launch {
  const query = new URLSearchParams(search);
  const action = query.get("action");
  const tool = query.get("tool");
  const view = query.get("view");
  if (action === "insert" || action === "spelling") return { view: "secili", action, tool: null };
  if (tool === "tutarlilik" || tool === "rapor" || tool === "dizin" || tool === "temizlik") {
    return { view: "teslim", action: null, tool };
  }
  if (view === "dokum" || view === "denetim") return { view: "dokum", action: null, tool: null };
  if (view === "teslim" || view === "ayarlar") return { view, action: null, tool: null };
  return { view: "secili", action: null, tool: null };
}

/**
 * Why the ribbon's "Ölüm tarihi ekle" cannot write at once, or null when the
 * selection is unambiguous and the normal checks may run.
 */
export function quickInsertBlocker(hasCatalog: boolean, alias: string | null, found: Lookup | null): string | null {
  if (!hasCatalog) return "Önce Ayarlar bölümünden DİA kataloğunuzu (CSV) seçin.";
  if (!alias) return "Metinde bir şahıs adı seçip yeniden deneyin.";
  if (!found || found.candidates.length === 0) return "Katalogda bu adla eşleşen şahıs yok.";
  if (!found.automatic) return "Birden fazla aday veya benzer yazım var. Listeden şahsı seçip Ölüm tarihi ekle'ye basın.";
  return null;
}
