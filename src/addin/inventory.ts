/**
 * Şahıs dökümü: every person mention in the main text and footnotes, with its
 * catalog match, the date written after it and a status. Word-independent;
 * word.ts supplies the paragraph texts. Follows VBA 0.3.41 `VfjPrepare` and
 * `VfjRefreshItem`.
 */
import {
  type Catalog,
  type Match,
  type NameHit,
  NameIndex,
  type Person,
  alreadyHasDate,
  bibliographyFlags,
  fieldsOf,
  sentenceAround,
  sharedFields,
  calendarsAgree,
  checkWrittenDate,
  type DateCheck,
  isAdjectiveUse,
  isTitleLike,
  words,
  lowerCaseWords,
  citationSpans,
  containsSpan,
  findMatches,
  findUnknownPeople,
  isValidDate,
  nameQuery,
  naturalOrder,
  normalize,
  personKey,
  safeCandidate,
  sourceName,
  startsWithDate,
  textName,
} from "../core/index.js";
import { type Confirmations, confirmedPerson, exclusionKey, mentionKey } from "./logic.js";

/** Segment.order of an endnote: ENDNOTE_BASE + its number in the paragraph (footnotes are 1, 2 …). */
export const ENDNOTE_BASE = 1000;

/** "Metin", "Dipnot" or "Sonnot" for a segment. */
export function placeOf(segment: Pick<Segment, "part" | "order"> | undefined): "Metin" | "Dipnot" | "Sonnot" {
  if (!segment || segment.part === "main") return "Metin";
  return segment.order > ENDNOTE_BASE ? "Sonnot" : "Dipnot";
}

/** A paragraph of the main text or of a footnote, in reading order. */
export interface Segment {
  readonly part: "main" | "note";
  /** Main-text paragraph index (for a footnote: the paragraph holding its reference). */
  readonly anchor: number;
  /** Position inside the anchor: 0 for the main paragraph, 1.. for its footnotes. */
  readonly order: number;
  readonly text: string;
}

export type Status = "match" | "spelling" | "near" | "choose" | "unknown";

export interface InventoryRow {
  readonly id: number;
  readonly segment: number;
  readonly part: "main" | "note";
  /** 0-based offset in the segment text. */
  readonly start: number;
  readonly length: number;
  /** Literal text, including a Turkish ending. */
  readonly text: string;
  readonly candidates: readonly Person[];
  /** Set when the identity is certain (single strong match). */
  readonly person: Person | null;
  /** Date already written right after the name, or "". */
  readonly date: string;
  readonly status: Status;
  /** First certain mention of this person in its part (main text, or footnotes and endnotes), in reading order. */
  readonly first: boolean;
  /** Identity taken from a confirmation saved in the document. */
  readonly confirmed: boolean;
  /**
   * How the identity was settled: the catalog alone, the user's confirmation,
   * the death date written after the name, or another certain mention of one
   * of the candidates in the document.
   */
  readonly resolved: "catalog" | "confirmed" | "date" | "context" | "topic" | null;
  /** For "topic": the shared field(s), e.g. "tefsir". */
  readonly topic?: string;
  /** The date written after the name against the catalog: "different", "inconsistent" (hijri vs Gregorian) or "ok". */
  readonly dateCheck: DateCheck;
}

const LINK_WORDS = new Set(["b", "bin", "bint", "ibn", "ibni", "el", "er", "es", "et", "ez", "en", "ed", "es", "ebu", "ebul", "ebi"]);

/** Every name word written ("Ebû Hâmid el-Gazzâlî") occurs in one of the person's names. */
const ownWords = new WeakMap<Person, Set<string>>();

function wordsContained(person: Person, literal: string): boolean {
  if (!/\s/.test(literal.trim())) return false;
  let own = ownWords.get(person);
  if (!own) {
    own = new Set([person.shortName, person.fullName, person.originalName].flatMap((name) => normalize(name).split(" ")).filter(Boolean));
    ownWords.set(person, own);
  }
  const written = normalize(nameQuery(literal)).split(" ").filter((w) => w && !LINK_WORDS.has(w));
  return written.length >= 2 && written.every((w) => own.has(w));
}

/** Date text right after a name: "(ö. 505/1111)" etc., or "". */
export function followingDate(after: string): string {
  const trimmed = after.replace(/^[  ]+/, "");
  if (!startsWithDate(trimmed)) return "";
  const closing = trimmed.indexOf(")");
  return closing > 0 && closing < 80 ? trimmed.slice(0, closing + 1) : "";
}

/** Whether the selected spelling already is the catalog spelling (`VfjSpellingMatches`). */
export function spellingMatches(person: Person, literal: string): boolean {
  const bare = nameQuery(literal);
  if ([person.shortName, person.fullName, person.originalName].some((name) => name === bare || textName(name) === bare)) return true;
  if (sourceName(person, bare) === bare) return true;
  // A usual short form built from the person's own name words: "Ebû Hâmid el-Gazzâlî".
  return wordsContained(person, literal);
}

/**
 * Person mentions of one paragraph. Main text: strict matches with context
 * rules. Footnotes (VBA test harness `FootnotePeople`): citation islands are
 * computed from all catalog spellings and strict matches inside them dropped.
 */
export function notePeopleOrMatches(segment: Pick<Segment, "part" | "text">, index: NameIndex): Match[] {
  // Citation islands ("Gazzâlî, el-Mustasfâ, I, 45", "(Gazzâlî, 1993: 45)")
  // are left out in both parts; the rest of the paragraph or footnote is read.
  const names = findMatches(segment.text, index, false).map((h): NameHit => [h.start, h.length, h.text]);
  let spans = citationSpans(segment.text, names);
  if (segment.part === "main") {
    // Main text narrates more than it cites: only islands with a locator
    // (page, volume, year) count, and "Gazzâlî (1993: 45) vurgular" keeps
    // its subject.
    spans = spans.filter(([a, b]) => {
      const island = segment.text.slice(a, b);
      return /\d/.test(island) && !/^[^,(]+\(\s*\d{4}/.test(island);
    });
  }
  return findMatches(segment.text, index, true, segment.part !== "note").filter((hit) => !containsSpan(spans, hit.start, hit.length));
}

/** Fields of the mention's sentence; when it names none, of its paragraph. */
function mentionFields(text: string, start: number, length: number): Set<string> {
  const sentence = fieldsOf(sentenceAround(text, start, length));
  return sentence.size > 0 ? sentence : fieldsOf(text);
}

interface Draft {
  segment: number;
  part: "main" | "note";
  start: number;
  length: number;
  text: string;
  candidates: readonly Person[];
  person: Person | null;
  resolved: "catalog" | "confirmed" | "date" | "context" | "topic" | null;
  after: string;
  /** Fields of the sentence around the mention (tefsir, tıp …). */
  fields: Set<string>;
  topic?: string;
}

export function buildInventory(
  segments: readonly Segment[],
  catalog: Catalog,
  index = new NameIndex(catalog.people),
  confirmations: Confirmations = {},
  exclusions: readonly string[] = [],
  spellings: Readonly<Record<string, string>> = {},
): InventoryRow[] {
  const excluded = new Set(exclusions);
  const ordered = segments
    .map((segment, i) => ({ segment, i }))
    .sort((a, b) => a.segment.anchor - b.segment.anchor || a.segment.order - b.segment.order);
  const rows: InventoryRow[] = [];
  const drafts: Draft[] = [];
  const seen = new Set<string>();
  const mains = ordered.filter((o) => o.segment.part === "main");
  const flags = bibliographyFlags(mains.map((o) => o.segment.text));
  const bibliography = new Set(mains.filter((_, k) => flags[k]).map((o) => o.segment.anchor));
  const lowerWords = lowerCaseWords(segments.map((s) => s.text));
  for (const { segment, i } of ordered) {
    if (bibliography.has(segment.anchor)) continue;
    const known = notePeopleOrMatches(segment, index);
    const unknown = findUnknownPeople(segment.text, index).filter(
      (u) => !known.some((k) => u.start < k.start + k.length && k.start < u.start + u.length),
    );
    const hits = [...known, ...unknown].sort((a, b) => a.start - b.start);
    const title = segment.part === "main" && isTitleLike(segment.text);
    for (const hit of hits) {
      if (excluded.size > 0 && excluded.has(exclusionKey(hit.text))) continue;
      if (isAdjectiveUse(segment.text, hit.start, hit.length, hit.text, lowerWords)) continue;
      // In a heading every word is capitalised: a one-word match that the
      // document also writes in lower case is that word ("Hayatı ve Eserleri").
      if (title && !/\s/.test(nameQuery(hit.text).trim()) && lowerWords.has(words(nameQuery(hit.text)).trim())) continue;
      let candidates = hit.key ? index.rank(hit.key, nameQuery(hit.text)) : [];
      const after = segment.text.slice(hit.start + hit.length);
      let person: Person | null = null;
      let resolved: Draft["resolved"] = null;
      let start = hit.start;
      let literal = hit.text;
      // "F. er-Râzî": an initial before the name picks the bearer whose other name starts with it.
      const initial = /(^|[\s(])(\p{Lu})\.\s?$/u.exec(segment.text.slice(Math.max(0, hit.start - 4), hit.start));
      if (initial && candidates.length > 1) {
        const letter = normalize(initial[2] as string);
        const head = normalize(nameQuery(hit.text)).split(" ").filter((w) => !LINK_WORDS.has(w))[0] ?? "";
        // The word the initial stands for comes right before the written name: "Fahreddin er-Râzî".
        const fits = candidates.filter((c) =>
          [c.shortName, naturalOrder(c.shortName), c.fullName].some((name) => {
            // Words of a künye ("Ebü’l-Fazl") are not abbreviated to an initial.
            const raw = normalize(name).split(" ").filter(Boolean);
            const list = raw.filter((w, k) => !LINK_WORDS.has(w) && w !== "l" && !["ebu", "ebi", "ebul", "ummu", "l"].includes(raw[k - 1] ?? ""));
            const at = list.indexOf(head);
            return at > 0 && (list[at - 1] as string).startsWith(letter);
          }),
        );
        if (fits.length === 1) {
          const at = segment.text.lastIndexOf(initial[2] as string, hit.start);
          start = at;
          literal = segment.text.slice(at, hit.start + hit.length);
          candidates = fits;
          person = fits[0] as Person;
          resolved = "catalog";
        }
      }
      const safe = safeCandidate(candidates, hit.text);
      if (!person && (safe || (candidates.length === 1 && wordsContained(candidates[0] as Person, hit.text)))) {
        person = safe ?? (candidates[0] as Person);
        resolved = "catalog";
      }
      if (!person) {
        const key = mentionKey(hit.text, segment.text.slice(0, hit.start), after);
        person = confirmedPerson(confirmations, key, candidates);
        if (person) resolved = "confirmed";
      }
      if (!person && candidates.length > 1) {
        // "Râzî (ö. 313/925)": the date written after the name picks the bearer.
        const dated = candidates.filter((c) => isValidDate(c.deathDate) && alreadyHasDate(after, c.deathDate));
        if (dated.length === 1) {
          person = dated[0] as Person;
          resolved = "date";
        }
      }
      drafts.push({
        segment: i, part: segment.part, start, length: hit.start + hit.length - start, text: literal, candidates, person, resolved, after,
        fields: person || candidates.length < 2 ? new Set() : mentionFields(segment.text, hit.start, hit.length),
      });
    }
  }
  // Document context: among a mention's candidates, the only one the document
  // names for certain elsewhere (full name, date or confirmation) is meant.
  const known = new Set(drafts.filter((d) => d.person).map((d) => personKey(d.person as Person)));
  const lastAt = new Map<string, number>(); // person key -> last draft index with that identity from the text itself
  drafts.forEach((draft, at) => {
    if (draft.person) {
      lastAt.set(personKey(draft.person), at);
      return;
    }
    if (draft.candidates.length === 0) return;
    const inDocument = draft.candidates.filter((c) => known.has(personKey(c)));
    if (inDocument.length === 1) {
      draft.person = inDocument[0] as Person;
      draft.resolved = "context";
    } else if (inDocument.length > 1) {
      // Several bearers in the document: the one named last before this
      // mention first, then the others; the choice stays with the user.
      const lastSeen = (c: Person) => lastAt.get(personKey(c)) ?? -1;
      inDocument.sort((a, b) => lastSeen(b) - lastSeen(a));
      // The sentence's subject decides between bearers the document names: "Râzî tefsirinde…".
      const scored = inDocument.map((c) => ({ c, shared: sharedFields(draft.fields, c.description) }));
      const best = Math.max(0, ...scored.map((s) => s.shared.length));
      const top = scored.filter((s) => s.shared.length === best);
      if (best > 0 && top.length === 1) {
        draft.person = (top[0] as { c: Person }).c;
        draft.resolved = "topic";
        draft.topic = (top[0] as { shared: string[] }).shared.join(", ");
        inDocument.sort((a, b) => (a === draft.person ? -1 : b === draft.person ? 1 : 0));
      }
      draft.candidates = [...inDocument, ...draft.candidates.filter((c) => !known.has(personKey(c)))];
    } else if (draft.fields.size > 0) {
      // No bearer named elsewhere: the subject only orders the candidates.
      const score = (c: Person) => sharedFields(draft.fields, c.description).length;
      draft.candidates = [...draft.candidates].sort((a, b) => score(b) - score(a));
    }
  });
  for (const draft of drafts) {
    const person = draft.person;
    let date = followingDate(draft.after);
    if (!date && person && alreadyHasDate(draft.after, person.deathDate)) date = person.deathDate;
    let status: Status;
    if (draft.candidates.length === 0) status = "unknown";
    else if (person) {
      const preferred = spellings[personKey(person)];
      const ok = preferred ? nameQuery(draft.text) === preferred : spellingMatches(person, draft.text);
      status = ok ? "match" : "spelling";
    } else if (draft.candidates.length === 1) status = "near";
    else status = "choose";
    let first = false;
    if (person) {
      // The main text and the notes each have their own first mention.
      const key = `${draft.part}|${personKey(person)}`;
      first = !seen.has(key);
      seen.add(key);
    }
    rows.push({
      id: rows.length,
      segment: draft.segment,
      part: draft.part,
      start: draft.start,
      length: draft.length,
      text: draft.text,
      candidates: draft.candidates,
      person,
      date,
      status,
      first,
      confirmed: draft.resolved === "confirmed",
      resolved: draft.resolved,
      ...(draft.topic ? { topic: draft.topic } : {}),
      dateCheck: person && date && isValidDate(person.deathDate) ? checkWrittenDate(date, person.deathDate) : date ? (calendarsAgree(date) ? "ok" : "inconsistent") : null,
    });
  }
  return rows;
}

export function hasDateIssue(row: Pick<InventoryRow, "dateCheck">): boolean {
  return row.dateCheck === "different" || row.dateCheck === "inconsistent";
}

/** Catalog column text for a row. */
export function catalogLabel(row: InventoryRow): string {
  if (row.person) return textName(row.person.shortName);
  if (row.candidates.length === 1) return `${textName((row.candidates[0] as Person).shortName)} (?)`;
  if (row.candidates.length > 1) return `${row.candidates.length} aday; seçim gerekli`;
  return "Katalogda yok";
}

/** Status text for a row, with date hints as in the VBA inventory. */
export function statusLabel(row: InventoryRow): string {
  const base: Record<Status, string> = {
    match: "Kaynak yazımıyla uyumlu",
    spelling: "Uyarı: kaynak yazımı farklı",
    near: "Yakın eşleşme; doğrulayın",
    choose: "Birden fazla şahıs eşleşiyor",
    unknown: "Katalogda yok",
  };
  let text = base[row.status];
  if (row.confirmed) text += " · onaylı kimlik";
  if (row.dateCheck === "different" && row.person) text += ` · Uyarı: metindeki tarih katalogdan farklı (katalog: ${row.person.deathDate})`;
  if (row.dateCheck === "inconsistent") text += ` · Uyarı: yazılan hicrî ve milâdî yıl birbirini tutmuyor${row.person ? ` (katalog: ${row.person.deathDate})` : ""}`;
  if (row.resolved === "date") text += " · kimlik yazılı tarihten";
  if (row.resolved === "context") text += " · kimlik belgedeki diğer geçişlerden";
  if (row.resolved === "topic") text += ` · kimlik konudan (${row.topic}); doğrulayın`;
  if (row.person && !isValidDate(row.person.deathDate)) text += " · Kaynakta tarih yok";
  else if (row.person && row.first && !row.date) text += " · İlk geçişte tarih yok";
  return text;
}

export interface InventorySummary {
  readonly mentions: number;
  readonly people: number;
  readonly firstWithoutDate: number;
  readonly spelling: number;
  readonly uncertain: number;
  /** Written dates that differ from the catalog or whose two calendars disagree. */
  readonly dateIssues: number;
}

export function summarize(rows: readonly InventoryRow[]): InventorySummary {
  const people = new Set(rows.filter((r) => r.person).map((r) => personKey(r.person as Person)));
  return {
    mentions: rows.length,
    people: people.size,
    firstWithoutDate: rows.filter((r) => r.first && !r.date && r.person && isValidDate(r.person.deathDate)).length,
    spelling: rows.filter((r) => r.status === "spelling").length,
    uncertain: rows.filter((r) => r.status === "near" || r.status === "choose").length,
    dateIssues: rows.filter((r) => hasDateIssue(r)).length,
  };
}
