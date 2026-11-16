/**
 * Thin Word layer (Office.js, WordApi 1.3). Every decision is made in
 * logic.ts; these functions only read or write the document.
 */
import type { NameIndex, Person } from "../core/index.js";
import { ENDNOTE_BASE, type InventoryRow, type Segment } from "./inventory.js";
import { parseSpellings } from "./prefs.js";
import { type LocalPerson, parseLocalPersons } from "./persons.js";
import { DATE_TAG_PREFIX, type DateControl, type LegacyResolution, isLegacyDateField, legacyPerson } from "./health.js";
import { type IndexLine, staticIndex, INDEX_CODE, isVefeyatIndexField, xeCode } from "./submit.js";
import { type SettingName, type Settings, parseConfirmations, parseExclusions, settingValue } from "./logic.js";

export interface SelectionText {
  /** Selected text as Word returns it. */
  readonly text: string;
  /** Rest of the paragraph after the selection. */
  readonly following: string;
  /** Paragraph text before the selection. */
  readonly preceding: string;
}

export async function readSelection(): Promise<SelectionText> {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    const paragraphEnd = selection.paragraphs.getLast().getRange("End");
    const following = selection.getRange("End").expandTo(paragraphEnd);
    following.load("text");
    const preceding = selection.paragraphs.getFirst().getRange("Start").expandTo(selection.getRange("Start"));
    preceding.load("text");
    await context.sync();
    return { text: selection.text, following: following.text, preceding: preceding.text };
  });
}

/** Writes " <date>" right after the selection inside a hidden, tagged content control. */
export async function insertDate(dateText: string, tag: string, title: string): Promise<number> {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    // The name stops being underlined as a misspelling: its right-click menu then carries Vefeyât.
    if (Office.context.requirements.isSetSupported("WordApiDesktop", "1.3")) selection.hasNoProofing = true;
    const inserted = selection.getRange("End").insertText(` ${dateText}`, "After");
    const control = inserted.insertContentControl();
    control.tag = tag;
    control.title = title;
    control.appearance = "Hidden";
    // Continue typing after the date, outside the control.
    control.getRange("After").select("Start");
    control.load("id");
    await context.sync();
    return control.id;
  });
}

/** Replaces the selected text and keeps the new text selected. */
export async function replaceSelection(text: string): Promise<void> {
  await Word.run(async (context) => {
    const replaced = context.document.getSelection().insertText(text, "Replace");
    replaced.select("Select");
    await context.sync();
  });
}

function searchText(value: string): string {
  // "^" starts a Word special character in search strings.
  return value.replace(/\^/g, "^^");
}

const BEFORE: ReadonlySet<string> = new Set(["Before", "AdjacentBefore"]);

/**
 * Searches the main text for each spelling before the selection.
 * With `select`, the first earlier mention found is selected.
 * Footnotes are not searched in this version.
 */
export async function findEarlier(spellings: readonly string[], select = false): Promise<boolean> {
  if (spellings.length === 0) return false;
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    const searches = spellings
      .filter((value) => value.length <= 250)
      .map((value) => {
        const results = context.document.body.search(searchText(value), { matchCase: false, matchWholeWord: false });
        results.load("items");
        return results;
      });
    await context.sync();
    const checks: Array<{ range: Word.Range; relation: OfficeExtension.ClientResult<Word.LocationRelation> }> = [];
    for (const results of searches) {
      for (const range of results.items.slice(0, 50)) {
        checks.push({ range, relation: range.compareLocationWith(selection) });
      }
    }
    await context.sync();
    const earlier = checks.find((check) => BEFORE.has(check.relation.value));
    if (earlier) {
      if (select) {
        earlier.range.select("Select");
        await context.sync();
      }
      return true;
    }
    // Footnotes (WordApi 1.5): a note whose reference precedes the selection.
    if (!Office.context.requirements.isSetSupported("WordApi", "1.5")) return false;
    const foot = context.document.body.footnotes;
    const end = context.document.body.endnotes;
    foot.load("items/body/text");
    end.load("items/body/text");
    await context.sync();
    const lowered = spellings.map((value) => value.toLocaleLowerCase("tr"));
    const containing = [...foot.items, ...end.items].filter((note) => {
      const text = note.body.text.toLocaleLowerCase("tr");
      return lowered.some((value) => text.includes(value));
    });
    const relations = containing.map((note) => ({ note, relation: note.reference.compareLocationWith(selection) }));
    await context.sync();
    const before = relations.find((item) => BEFORE.has(item.relation.value));
    if (!before) return false;
    if (select) {
      for (const value of spellings) {
        const found = before.note.body.search(searchText(value), { matchCase: false, matchWholeWord: false });
        found.load("items");
        await context.sync();
        if (found.items[0]) {
          found.items[0].select("Select");
          await context.sync();
          break;
        }
      }
    }
    return true;
  });
}

const CONFIRMATIONS_KEY = "vefeyat.confirmations";
const AUDIT_KEY = "vefeyat.audit";
const EXCLUSIONS_KEY = "vefeyat.exclusions";
const SPELLINGS_KEY = "vefeyat.spellings";
const PERSONS_KEY = "vefeyat.persons";

/** Kendi şahıslarım of this document. */
export function readLocalPersons(): LocalPerson[] {
  return parseLocalPersons(Office.context.document.settings.get(PERSONS_KEY));
}

export function writeLocalPersons(list: readonly LocalPerson[]): Promise<void> {
  if (list.length === 0) Office.context.document.settings.remove(PERSONS_KEY);
  else Office.context.document.settings.set(PERSONS_KEY, JSON.stringify(list));
  return saveSettings();
}

/** Preferred spellings of this document (person key -> spelling). */
export function readSpellings(): Record<string, string> {
  return parseSpellings(Office.context.document.settings.get(SPELLINGS_KEY));
}

export function writeSpellings(map: Readonly<Record<string, string>>): Promise<void> {
  if (Object.keys(map).length === 0) Office.context.document.settings.remove(SPELLINGS_KEY);
  else Office.context.document.settings.set(SPELLINGS_KEY, JSON.stringify(map));
  return saveSettings();
}

function saveSettings(): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.document.settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error(result.error.message));
    });
  });
}

/** "Şahıs değil" phrases of this document (exclusion keys). */
export function readExclusions(): string[] {
  return parseExclusions(Office.context.document.settings.get(EXCLUSIONS_KEY));
}

export function writeExclusions(list: readonly string[]): Promise<void> {
  if (list.length === 0) Office.context.document.settings.remove(EXCLUSIONS_KEY);
  else Office.context.document.settings.set(EXCLUSIONS_KEY, JSON.stringify(list));
  return saveSettings();
}

/** Paused-audit bookmark saved in the document (raw JSON or null). */
export function readAuditBookmark(): unknown {
  return Office.context.document.settings.get(AUDIT_KEY);
}

export function writeAuditBookmark(value: string | null): Promise<void> {
  if (value === null) Office.context.document.settings.remove(AUDIT_KEY);
  else Office.context.document.settings.set(AUDIT_KEY, value);
  return new Promise((resolve) => Office.context.document.settings.saveAsync(() => resolve()));
}

/** Identity confirmations saved in the document. */
export function readConfirmations(): Record<string, string> {
  return parseConfirmations(Office.context.document.settings.get(CONFIRMATIONS_KEY));
}

export function writeConfirmations(map: Record<string, string>): Promise<void> {
  Office.context.document.settings.set(CONFIRMATIONS_KEY, JSON.stringify(map));
  return new Promise((resolve, reject) => {
    Office.context.document.settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error(result.error.message));
    });
  });
}

const SETTING_KEYS: Record<SettingName, string> = {
  prefix: "vefeyat.prefix",
  scope: "vefeyat.scope",
  nameForm: "vefeyat.nameForm",
  auditMode: "vefeyat.auditMode",
};

/** Re-reads settings saved by the ribbon commands (another runtime). */
export function refreshSettings(): Promise<void> {
  return new Promise((resolve) => {
    Office.context.document.settings.refreshAsync(() => resolve());
  });
}

export function readSettings(): Settings {
  const get = (name: SettingName): unknown => Office.context.document.settings.get(SETTING_KEYS[name]);
  return {
    prefix: settingValue("prefix", get("prefix")),
    scope: settingValue("scope", get("scope")),
    nameForm: settingValue("nameForm", get("nameForm")),
    auditMode: settingValue("auditMode", get("auditMode")),
  };
}

/** Stores one setting in the document (saved with the document). */
export function writeSetting(name: SettingName, value: string): Promise<void> {
  Office.context.document.settings.set(SETTING_KEYS[name], value);
  return new Promise((resolve, reject) => {
    Office.context.document.settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error(result.error.message));
    });
  });
}

export function onSelectionChanged(handler: () => void): void {
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handler);
}

// ---- Whole-document reading for Şahıs dökümü and Belgeyi denetle ----


/** The footnote or endnote a note segment comes from, in its anchor paragraph. */
function noteOf(lists: { foot: Word.NoteItemCollection; end?: Word.NoteItemCollection | undefined }, order: number): Word.NoteItem | undefined {
  return order > ENDNOTE_BASE ? lists.end?.items[order - ENDNOTE_BASE - 1] : lists.foot.items[order - 1];
}

/** Footnote and endnote bodies of the whole document (WordApi 1.5). */
async function noteBodies(context: Word.RequestContext): Promise<Word.Body[]> {
  if (!Office.context.requirements.isSetSupported("WordApi", "1.5")) return [];
  const foot = context.document.body.footnotes;
  const end = context.document.body.endnotes;
  foot.load("items");
  end.load("items");
  await context.sync();
  return [...foot.items, ...end.items].map((n) => n.body);
}

/** Word requests queued per sync in long operations. */
const CHUNK = 200;
/** Marks per sync: each one makes Word re-lay out the page, so keep batches small. */
const MARK_CHUNK = 40;

/** Lets Word repaint and take keystrokes between batches. */
const breathe = () => new Promise<void>((resolve) => setTimeout(resolve, 60));

export type Progress = (done: number, total: number) => void;

/** Main-text paragraphs and, where WordApi 1.5 is available, footnotes and endnotes. */
export async function readSegments(progress: Progress = () => {}): Promise<{ segments: Segment[]; footnotes: boolean }> {
  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/text");
    await context.sync();
    const segments: Segment[] = paragraphs.items.map((p, i) => ({ part: "main", anchor: i, order: 0, text: p.text }));
    const footnotes = Office.context.requirements.isSetSupported("WordApi", "1.5");
    if (footnotes) {
      // In chunks: one request per paragraph of a 500-page thesis is too much for one sync.
      for (let from = 0; from < paragraphs.items.length; from += CHUNK) {
        const notes = paragraphs.items.slice(from, from + CHUNK).map((p) => {
          const foot = p.getRange("Whole").footnotes;
          const end = p.getRange("Whole").endnotes;
          foot.load("items/body/text");
          end.load("items/body/text");
          return { foot, end };
        });
        await context.sync();
        notes.forEach(({ foot, end }, k) => {
          foot.items.forEach((note, i) => segments.push({ part: "note", anchor: from + k, order: i + 1, text: note.body.text }));
          end.items.forEach((note, i) => segments.push({ part: "note", anchor: from + k, order: ENDNOTE_BASE + i + 1, text: note.body.text }));
        });
        progress(Math.min(from + CHUNK, paragraphs.items.length), paragraphs.items.length);
      }
    }
    return { segments, footnotes };
  });
}

function occurrenceBefore(text: string, literal: string, start: number): number {
  let count = 0;
  for (let at = text.indexOf(literal); at >= 0 && at < start; at = text.indexOf(literal, at + 1)) count++;
  return count;
}

/** Selects one mention found by the inventory. Returns false when the text changed. */
export async function selectMention(segment: Segment, start: number, literal: string): Promise<boolean> {
  if (literal.length === 0 || literal.length > 250) return false;
  return Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs;
    paragraph.load("items");
    await context.sync();
    const host = paragraph.items[segment.anchor];
    if (!host) return false;
    let scope: Word.Body | Word.Range = host.getRange("Whole");
    if (segment.part === "note") {
      const foot = host.getRange("Whole").footnotes;
      const end = host.getRange("Whole").endnotes;
      foot.load("items");
      end.load("items");
      await context.sync();
      const note = noteOf({ foot, end }, segment.order);
      if (!note) return false;
      scope = note.body;
    }
    const results = scope.search(searchText(literal), { matchCase: true, matchWholeWord: false });
    results.load("items");
    await context.sync();
    const target = results.items[occurrenceBefore(segment.text, literal, start)];
    if (!target) return false;
    target.select("Select");
    await context.sync();
    return true;
  });
}

// ---- Teslim ve dizin ----

/** Opens a generated .docx as a new Word document (WordApi 1.3). */
export async function openNewDocument(base64: string): Promise<void> {
  await Word.run(async (context) => {
    context.application.createDocument(base64).open();
    await context.sync();
  });
}

/** Name of the current document, for report titles. */
export function documentName(): string {
  const url = Office.context.document.url ?? "";
  const name = decodeURIComponent(url.split(/[\\/]/).pop() ?? "");
  return name || "Adsız belge";
}

/**
 * Teslim öncesi temizle: removes Vefeyât's date content controls (keeping
 * the date text) from the body and footnotes, and the vefeyat.* settings.
 */
export async function finalizeDocument(): Promise<{ controls: number; settings: number; marks: number }> {
  let marks = 0;
  if (canIndex()) marks = (await removeIndexMarks()).marks;
  const controls = await Word.run(async (context) => {
    const lists = [context.document.body.contentControls, ...(await noteBodies(context)).map((b) => b.contentControls)];
    let count = 0;
    for (let from = 0; from < lists.length; from += CHUNK) {
      const chunk = lists.slice(from, from + CHUNK);
      for (const list of chunk) list.load("items/tag");
      await context.sync();
      for (const list of chunk) {
        for (const control of list.items) {
          if (!control.tag?.startsWith("vefeyat:")) continue;
          control.cannotDelete = false;
          control.delete(true);
          count++;
        }
      }
      await context.sync();
    }
    return count;
  });
  const settings = Office.context.document.settings;
  let removed = 0;
  for (const key of [CONFIRMATIONS_KEY, AUDIT_KEY, EXCLUSIONS_KEY, SPELLINGS_KEY, PERSONS_KEY, ...Object.values(SETTING_KEYS)]) {
    if (settings.get(key) !== null && settings.get(key) !== undefined) removed++;
    settings.remove(key);
  }
  await new Promise<void>((resolve, reject) => {
    settings.saveAsync((result) => (result.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error(result.error.message))));
  });
  return { controls, settings: removed, marks };
}

// ---- Şahıslar dizini (WordApi 1.5: insertField) ----

export function canIndex(): boolean {
  return Office.context.requirements.isSetSupported("WordApi", "1.5");
}

const INDEX_TAG = "vefeyat:index";
const INDEX_HEADING = "Şahıslar Dizini";

/** Deletes Vefeyât's XE and INDEX fields from body and footnotes. */
async function deleteIndexFields(context: Word.RequestContext): Promise<{ marks: number; indexes: number }> {
  const lists = [context.document.body.fields, ...(await noteBodies(context)).map((b) => b.fields)];
  let marks = 0;
  let indexes = 0;
  for (let from = 0; from < lists.length; from += CHUNK) {
    const chunk = lists.slice(from, from + CHUNK);
    for (const list of chunk) list.load("items/code");
    await context.sync();
    for (const list of chunk) {
      for (const field of list.items) {
        if (!isVefeyatIndexField(field.code)) continue;
        if (/^\s*INDEX\b/i.test(field.code)) indexes++;
        else marks++;
        field.delete();
      }
    }
    await context.sync();
  }
  return { marks, indexes };
}

/**
 * The index block: heading, page break and INDEX field inside one content
 * control, so removing it leaves nothing behind. Indexes made by the first
 * version (no control) are recognised by their INDEX field and heading.
 */
async function deleteIndexBlock(context: Word.RequestContext): Promise<boolean> {
  const controls = context.document.body.contentControls.getByTag(INDEX_TAG);
  controls.load("items");
  await context.sync();
  for (const control of controls.items) {
    control.cannotDelete = false;
    control.cannotEdit = false;
    control.delete(false);
  }
  await context.sync();
  return controls.items.length > 0;
}

/** Headings "Şahıslar Dizini" left by the first version, right after which the index stood. */
async function deleteOldHeadings(context: Word.RequestContext): Promise<void> {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("items/text,items/styleBuiltIn");
  await context.sync();
  const items = paragraphs.items;
  for (let i = items.length - 1; i >= 0; i--) {
    const p = items[i] as Word.Paragraph;
    if (p.text.trim() !== INDEX_HEADING || p.styleBuiltIn !== "Heading1") continue;
    // Only a heading with nothing but empty paragraphs after it.
    if (items.slice(i + 1).some((q) => q.text.trim().length > 0)) continue;
    for (const q of items.slice(i + 1)) q.delete();
    p.delete();
  }
  await context.sync();
}

/**
 * Turns Vefeyât's INDEX field into plain paragraphs (after updating it), so
 * the index survives once its marks are gone. Page numbers no longer follow
 * later edits. Returns whether an index was found.
 */
async function freezeIndex(context: Word.RequestContext): Promise<boolean> {
  const fields = context.document.body.fields;
  fields.load("items/code");
  await context.sync();
  const indexes = fields.items.filter((f) => /^\s*INDEX\b/i.test(f.code) && isVefeyatIndexField(f.code));
  if (indexes.length === 0) return false;
  for (const field of indexes) field.updateResult();
  await context.sync();
  // Newer Word turns a field into its result text in place, columns and styles kept.
  if (Office.context.requirements.isSetSupported("WordApiDesktop", "1.4")) {
    for (const field of indexes) field.unlink();
    await context.sync();
    return true;
  }
  const results = indexes.map((field) => {
    const paragraphs = field.result.paragraphs;
    paragraphs.load("items/text,items/style");
    return { field, paragraphs, first: field.result.paragraphs.getFirst() };
  });
  await context.sync();
  for (const { field, paragraphs, first } of results) {
    for (const p of paragraphs.items) {
      if (p.text.trim().length === 0) continue;
      const copy = first.insertParagraph(p.text, "Before");
      copy.style = p.style;
    }
    field.delete();
  }
  await context.sync();
  return true;
}

/** Dizin işaretlerini kaldır: the index stays as text, the hidden XE marks go. */
export async function removeIndexMarks(): Promise<{ marks: number; index: boolean }> {
  return Word.run(async (context) => {
    const index = await freezeIndex(context);
    const fields = await deleteIndexFields(context);
    return { marks: fields.marks, index };
  });
}

/** Dizini tamamen kaldır: marks, index, its heading and page break. */
export async function removeIndex(): Promise<{ marks: number; index: boolean }> {
  return Word.run(async (context) => {
    const block = await deleteIndexBlock(context);
    const fields = await deleteIndexFields(context);
    if (fields.indexes > 0) await deleteOldHeadings(context);
    return { marks: fields.marks, index: block || fields.indexes > 0 };
  });
}

/** Error of one step of a long operation, named for the message. */
export class StepError extends Error {
  constructor(
    readonly step: string,
    readonly cause: unknown,
  ) {
    super(`${step}: ${String(cause)}`);
  }
}

const PLACE_TAG = "vefeyat:index-place";

/** Step 1: removes the old index and its marks; leaves a placeholder where it stood. */
async function clearOldIndex(): Promise<void> {
  await Word.run(async (context) => {
    const blocks = context.document.body.contentControls.getByTag(INDEX_TAG);
    blocks.load("items");
    const fields = context.document.body.fields;
    fields.load("items/code");
    await context.sync();
    const previous = blocks.items[0] ?? fields.items.find((f) => /^\s*INDEX\b/i.test(f.code) && isVefeyatIndexField(f.code));
    if (previous) {
      const where = "result" in previous ? previous.result : previous.getRange("Whole");
      const place = where.paragraphs.getFirst().insertParagraph("", "Before").insertContentControl();
      place.tag = PLACE_TAG;
      place.appearance = "Hidden";
    }
    await context.sync();
    await deleteIndexBlock(context);
    const removed = await deleteIndexFields(context);
    if (removed.indexes > 0 && blocks.items.length === 0) await deleteOldHeadings(context);
  });
}

/** Step 2: an XE mark before every certain mention, in chunks; a failing chunk is retried one by one. */
async function insertMarks(segments: readonly Segment[], rows: readonly InventoryRow[], progress: Progress): Promise<number> {
  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items");
    await context.sync();
    const noteLists = new Map<number, { foot: Word.NoteItemCollection; end: Word.NoteItemCollection }>();
    const anchors = [...new Set(rows.filter((r) => segments[r.segment]?.part === "note").map((r) => (segments[r.segment] as Segment).anchor))];
    for (let from = 0; from < anchors.length; from += CHUNK) {
      for (const anchor of anchors.slice(from, from + CHUNK)) {
        const host = paragraphs.items[anchor];
        if (!host) continue;
        const foot = host.getRange("Whole").footnotes;
        const end = host.getRange("Whole").endnotes;
        foot.load("items");
        end.load("items");
        noteLists.set(anchor, { foot, end });
      }
      await context.sync();
    }
    // One search per paragraph and spelling serves all its mentions.
    const cache = new Map<string, Word.RangeCollection>();
    const find = (row: InventoryRow) => {
      const segment = segments[row.segment];
      if (!segment || row.text.length > 250 || !row.person) return null;
      const host = paragraphs.items[segment.anchor];
      if (!host) return null;
      const scope: Word.Body | Word.Range | undefined =
        segment.part === "note" ? (() => {
          const lists = noteLists.get(segment.anchor);
          return lists ? noteOf(lists, segment.order)?.body : undefined;
        })() : host.getRange("Whole");
      if (!scope) return null;
      const key = `${row.segment}|${row.text}`;
      let results = cache.get(key);
      if (!results) {
        results = scope.search(searchText(row.text), { matchCase: true, matchWholeWord: false });
        results.load("items");
        cache.set(key, results);
      }
      return { row, results, nth: occurrenceBefore(segment.text, row.text, row.start) };
    };
    // From the end backwards: a mark never shifts the mentions still to be found before it.
    const ordered = [...rows].reverse();
    // Chunks never split a paragraph: all its mentions are found before any mark goes in.
    const chunks: InventoryRow[][] = [];
    for (const row of ordered) {
      const last = chunks[chunks.length - 1];
      if (last && (last.length < MARK_CHUNK || (last[last.length - 1] as InventoryRow).segment === row.segment)) last.push(row);
      else chunks.push([row]);
    }
    let marks = 0;
    let done = 0;
    for (const chunk of chunks) {
      const searches = chunk.map(find).filter((x) => x !== null);
      await context.sync();
      const targets = searches.map(({ row, results, nth }) => ({ row, target: results.items[nth] })).filter((t) => t.target);
      try {
        for (const { row, target } of targets) (target as Word.Range).insertField("Before", "XE", xeCode(row.person as Person));
        await context.sync();
        marks += targets.length;
      } catch {
        // One mark Word refuses (e.g. inside a field or a protected area) must not stop the rest.
        for (const { row, target } of targets) {
          try {
            (target as Word.Range).insertField("Before", "XE", xeCode(row.person as Person));
            await context.sync();
            marks++;
          } catch {
            /* skipped; counted as missing */
          }
        }
      }
      cache.clear(); // ranges of this chunk are done with
      done += chunk.length;
      progress(done, ordered.length);
      await breathe();
    }
    return marks;
  });
}

/**
 * Step 3: heading and INDEX field. Tries the tagged block with its page
 * break inside; if Word refuses, without the inner page break; last, plain
 * paragraphs without a block.
 */
async function insertIndexBlock(lines: readonly IndexLine[] | null = null): Promise<void> {
  const attempt = (mode: "block" | "block-outer-break" | "plain") =>
    Word.run(async (context) => {
      const places = context.document.body.contentControls.getByTag(PLACE_TAG);
      places.load("items");
      await context.sync();
      const place = places.items[0];
      const body = context.document.body;
      if (!place && mode !== "block") body.insertBreak("Page", "End");
      const first = place ? place.getRange("Whole").paragraphs.getFirst() : body.insertParagraph("", "End");
      if (place) place.delete(true);
      first.insertText(INDEX_HEADING, "Replace");
      first.styleBuiltIn = "Heading1";
      let holder = first.insertParagraph("", "After");
      holder.styleBuiltIn = "Normal";
      if (lines) {
        // Static index (page lookup): plain paragraphs, letters bold.
        let last: Word.Paragraph | null = null;
        for (const line of lines) {
          const paragraph: Word.Paragraph = last ? last.insertParagraph(line.text, "After") : holder;
          if (!last) paragraph.insertText(line.text, "Replace");
          paragraph.styleBuiltIn = "Normal";
          paragraph.font.bold = line.kind === "letter";
          if (line.kind === "letter") paragraph.spaceBefore = 8;
          last = paragraph;
        }
        if (last) holder = last;
      }
      if (mode !== "plain") {
        const control = first.getRange("Whole").expandTo(holder.getRange("Whole")).insertContentControl();
        control.tag = INDEX_TAG;
        control.title = "Vefeyât: şahıslar dizini";
        control.appearance = "BoundingBox";
        if (mode === "block") control.insertBreak("Page", "Start");
      }
      if (!lines) {
        const index = holder.getRange("Start").insertField("Start", "Index", INDEX_CODE);
        await context.sync();
        index.updateResult();
      }
      await context.sync();
    });
  try {
    await attempt("block");
  } catch {
    try {
      await attempt("block-outer-break");
    } catch {
      await attempt("plain");
    }
  }
}

/**
 * Şahıslar dizini: replaces Vefeyât's index marks with one XE field before
 * every certain mention and puts the index at the end of the document (or
 * where the previous index stood). Each step reports its own failure.
 */
export async function buildIndex(segments: readonly Segment[], rows: readonly InventoryRow[], progress: Progress = () => {}): Promise<{ marks: number; missing: number }> {
  try {
    await clearOldIndex();
  } catch (error) {
    throw new StepError("Eski dizin kaldırılamadı", error);
  }
  let marks: number;
  try {
    marks = await insertMarks(segments, rows, progress);
  } catch (error) {
    throw new StepError("Dizin işaretleri eklenemedi", error);
  }
  try {
    await insertIndexBlock();
  } catch (error) {
    throw new StepError(`${marks} dizin işareti eklendi, ama dizin sayfası oluşturulamadı`, error);
  }
  return { marks, missing: rows.length - marks };
}

// ---- Belge sağlığı ve Tarihin kaynağı ----

/** Content-control lists of the body and (WordApi 1.5) the footnotes, loaded. */
async function controlLists(context: Word.RequestContext): Promise<Array<{ part: "main" | "note"; list: Word.ContentControlCollection }>> {
  const lists: Array<{ part: "main" | "note"; list: Word.ContentControlCollection }> = [{ part: "main", list: context.document.body.contentControls }];
  for (const body of await noteBodies(context)) lists.push({ part: "note", list: body.contentControls });
  for (let from = 0; from < lists.length; from += CHUNK) {
    for (const { list } of lists.slice(from, from + CHUNK)) list.load("items/id,items/tag,items/text");
    await context.sync();
  }
  return lists;
}

async function findControl(context: Word.RequestContext, id: number): Promise<Word.ContentControl | null> {
  for (const { list } of await controlLists(context)) {
    const control = list.items.find((c) => c.id === id);
    if (control) return control;
  }
  return null;
}

/** Every Vefeyât date control, in reading order of body and footnotes. */
export async function readDateControls(): Promise<DateControl[]> {
  return Word.run(async (context) =>
    (await controlLists(context)).flatMap(({ part, list }) =>
      list.items.filter((c) => c.tag?.startsWith(DATE_TAG_PREFIX)).map((c) => ({ id: c.id, part, tag: c.tag, text: c.text })),
    ),
  );
}

/** Rewrites the text of one date control. Returns false when it is gone. */
export async function writeDateControl(id: number, text: string): Promise<boolean> {
  return Word.run(async (context) => {
    const control = await findControl(context, id);
    if (!control) return false;
    control.cannotEdit = false;
    control.insertText(text, "Replace");
    await context.sync();
    return true;
  });
}

export async function selectDateControl(id: number): Promise<boolean> {
  return Word.run(async (context) => {
    const control = await findControl(context, id);
    if (!control) return false;
    control.select("Select");
    await context.sync();
    return true;
  });
}

/** The Vefeyât date control the selection is in or contains, or null. */
export async function selectedDateControl(): Promise<DateControl | null> {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    const parent = selection.parentContentControlOrNullObject;
    parent.load("id,tag,text");
    const inside = selection.contentControls;
    inside.load("items/id,items/tag,items/text");
    // A name selected right before its date: look just after the selection.
    const after = selection.getRange("End").expandTo(selection.paragraphs.getFirst().getRange("End")).contentControls;
    after.load("items/id,items/tag,items/text");
    await context.sync();
    const candidates = [...(parent.isNullObject ? [] : [parent]), ...inside.items, ...after.items.slice(0, 1)];
    const control = candidates.find((c) => c.tag?.startsWith(DATE_TAG_PREFIX));
    return control ? { id: control.id, part: "main", tag: control.tag, text: control.text } : null;
  });
}

// ---- Denetim kapsamı ----

export interface Coverage {
  readonly paragraphs: number;
  readonly footnotes: number | null;
  readonly endnotes: number | null;
  readonly headers: number;
  readonly comments: number | null;
}

/** What Vefeyât scans (body, footnotes, endnotes) and what it does not (headers, comments, text boxes). */
export async function readCoverage(): Promise<Coverage> {
  return Word.run(async (context) => {
    const body = context.document.body;
    const paragraphs = body.paragraphs;
    paragraphs.load("items");
    const sections = context.document.sections;
    sections.load("items");
    const v15 = Office.context.requirements.isSetSupported("WordApi", "1.5");
    const v14 = Office.context.requirements.isSetSupported("WordApi", "1.4");
    const footnotes = v15 ? body.footnotes : null;
    const endnotes = v15 ? body.endnotes : null;
    footnotes?.load("items");
    endnotes?.load("items");
    const comments = v14 ? body.getComments() : null;
    comments?.load("items");
    await context.sync();
    const parts = sections.items.flatMap((s) => (["Primary", "FirstPage", "EvenPages"] as const).flatMap((k) => [s.getHeader(k), s.getFooter(k)]));
    for (const part of parts) part.load("text");
    await context.sync();
    return {
      paragraphs: paragraphs.items.length,
      footnotes: footnotes ? footnotes.items.length : null,
      endnotes: endnotes ? endnotes.items.length : null,
      headers: parts.filter((p) => p.text.trim().length > 0).length,
      comments: comments ? comments.items.length : null,
    };
  });
}

/**
 * Alan gölgeleri: Vefeyât dates are hidden content controls; this shows them
 * with a frame (to see what Vefeyât added) or hides the frame again.
 * Returns the new state and the number of controls changed.
 */
export async function toggleDateFrames(): Promise<{ visible: boolean; count: number }> {
  return Word.run(async (context) => {
    const lists = await controlLists(context);
    for (const { list } of lists) list.load("items/tag,items/appearance");
    await context.sync();
    const controls = lists.flatMap(({ list }) => list.items.filter((c) => c.tag?.startsWith(DATE_TAG_PREFIX)));
    const visible = controls.some((c) => c.appearance === "Hidden");
    for (const c of controls) c.appearance = visible ? "BoundingBox" : "Hidden";
    await context.sync();
    return { visible, count: controls.length };
  });
}

// ---- Sayfa numaralarından dizin (WordApiDesktop 1.2: Range.pages) ----

/** Whether this Word can tell on which page a range is. */
export function canReadPages(): boolean {
  return Office.context.requirements.isSetSupported("WordApiDesktop", "1.2");
}

/** Physical page (1-based) of the selection: "this page carries number 1". */
export async function selectionPage(): Promise<number | null> {
  return Word.run(async (context) => {
    const pages = context.document.getSelection().pages;
    pages.load("items/index");
    await context.sync();
    return pages.items[0]?.index ?? null;
  });
}

/**
 * Şahıslar dizini without marks: reads the page of every certain mention
 * (paragraph pages; the mention itself when its paragraph spans pages;
 * footnotes by their reference) and writes the index as plain text.
 * The document is only read until the index block is written.
 */
export async function buildPageIndex(
  segments: readonly Segment[],
  rows: readonly InventoryRow[],
  firstArabic: number,
  progress: Progress = () => {},
): Promise<{ people: number; hits: number; missing: number }> {
  try {
    await clearOldIndex();
  } catch (error) {
    throw new StepError("Eski dizin kaldırılamadı", error);
  }
  let hits: Array<{ person: Person; page: number }>;
  try {
    hits = await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      await context.sync();
      const out: Array<{ person: Person; page: number }> = [];
      const mainRows = new Map<number, InventoryRow[]>();
      const noteRows = new Map<number, InventoryRow[]>();
      for (const row of rows) {
        const segment = segments[row.segment];
        if (!segment || !row.person) continue;
        const map = segment.part === "main" ? mainRows : noteRows;
        map.set(segment.anchor, [...(map.get(segment.anchor) ?? []), row]);
      }
      const total = mainRows.size + noteRows.size;
      let done = 0;
      // Main text: pages of each paragraph that names someone.
      const spanning: Array<{ anchor: number; rows: InventoryRow[] }> = [];
      const mainList = [...mainRows.entries()];
      for (let from = 0; from < mainList.length; from += CHUNK) {
        const chunk = mainList.slice(from, from + CHUNK).map(([anchor, list]) => {
          const pages = paragraphs.items[anchor]?.getRange("Whole").pages;
          pages?.load("items/index");
          return { anchor, list, pages };
        });
        await context.sync();
        for (const { anchor, list, pages } of chunk) {
          if (!pages || pages.items.length === 0) continue;
          if (pages.items.length === 1) for (const row of list) out.push({ person: row.person as Person, page: (pages.items[0] as Word.Page).index });
          else spanning.push({ anchor, rows: list });
        }
        done += chunk.length;
        progress(done, total);
      }
      // A paragraph over a page break: the page of each mention itself.
      for (let from = 0; from < spanning.length; from += 20) {
        const lookups = spanning.slice(from, from + 20).flatMap(({ anchor, rows: list }) => {
          const segment = segments[(list[0] as InventoryRow).segment] as Segment;
          const host = paragraphs.items[anchor] as Word.Paragraph;
          return list.map((row) => {
            const results = host.getRange("Whole").search(searchText(row.text), { matchCase: true, matchWholeWord: false });
            results.load("items");
            return { row, results, nth: occurrenceBefore(segment.text, row.text, row.start) };
          });
        });
        await context.sync();
        const pageLists = lookups.map(({ row, results, nth }) => {
          const target = results.items[nth];
          const pages = target?.pages;
          pages?.load("items/index");
          return { row, pages };
        });
        await context.sync();
        for (const { row, pages } of pageLists) {
          const page = pages?.items[0]?.index;
          if (page) out.push({ person: row.person as Person, page });
        }
      }
      // Footnotes: the page of the reference mark in the text.
      const noteList = [...noteRows.entries()];
      for (let from = 0; from < noteList.length; from += CHUNK) {
        const chunk = noteList.slice(from, from + CHUNK).map(([anchor, list]) => {
          const range = paragraphs.items[anchor]?.getRange("Whole");
          const foot = range?.footnotes;
          const end = range?.endnotes;
          foot?.load("items");
          end?.load("items");
          return { list, notes: foot ? { foot, end } : undefined };
        });
        await context.sync();
        const pageLists = chunk.flatMap(({ list, notes }) =>
          list.map((row) => {
            const segment = segments[row.segment] as Segment;
            const pages = notes ? noteOf(notes, segment.order)?.reference.pages : undefined;
            pages?.load("items/index");
            return { row, pages };
          }),
        );
        await context.sync();
        for (const { row, pages } of pageLists) {
          const page = pages?.items[0]?.index;
          if (page) out.push({ person: row.person as Person, page });
        }
        done += chunk.length;
        progress(done, total);
      }
      return out;
    });
  } catch (error) {
    throw new StepError("Sayfa numaraları okunamadı", error);
  }
  const lines = staticIndex(hits, firstArabic);
  try {
    await insertIndexBlock(lines);
  } catch (error) {
    throw new StepError("Sayfalar okundu, ama dizin sayfası yazılamadı", error);
  }
  return { people: lines.filter((l) => l.kind === "entry").length, hits: hits.length, missing: rows.length - hits.length };
}

// ---- VBA belgeleri (ADDIN Vefeyat.Date) → içerik denetimleri ----

async function legacyFields(context: Word.RequestContext): Promise<Word.Field[]> {
  const bodies: Word.Body[] = [context.document.body, ...(await noteBodies(context))];
  const found: Word.Field[] = [];
  for (let from = 0; from < bodies.length; from += CHUNK) {
    const lists = bodies.slice(from, from + CHUNK).map((b) => {
      const fields = b.fields;
      fields.load("items/code");
      return fields;
    });
    await context.sync();
    for (const list of lists) for (const field of list.items) if (isLegacyDateField(field.code)) found.push(field);
  }
  return found;
}

/** Number of dates the VBA add-in wrote into this document (Word fields). */
export async function countLegacyDates(): Promise<number> {
  if (!Office.context.requirements.isSetSupported("WordApi", "1.4")) return 0;
  return Word.run(async (context) => (await legacyFields(context)).length);
}

/**
 * Replaces each VBA date field with the same text inside a Vefeyât date
 * control, bound to the person whose name stands right before it. A field
 * whose name cannot be found is left untouched.
 */
export async function convertLegacyDates(
  index: NameIndex,
  tagFor: (person: Person | null) => string,
  progress: Progress = () => {},
): Promise<{ converted: number; unresolved: number; skipped: number }> {
  return Word.run(async (context) => {
    const fields = await legacyFields(context);
    const items = fields.map((field) => {
      const paragraph = field.result.paragraphs.getFirst();
      paragraph.load("text");
      field.result.load("text");
      return { field, paragraph };
    });
    await context.sync();
    // Several dates in one paragraph: the k-th field has the k-th occurrence of its text.
    const seen = new Map<string, number>();
    const plans = items.map(({ field, paragraph }) => {
      const dateText = field.result.text;
      const key = `${paragraph.text}\u0000${dateText}`;
      const k = seen.get(key) ?? 0;
      seen.set(key, k + 1);
      let at = -1;
      for (let n = 0; n <= k; n++) at = paragraph.text.indexOf(dateText, at + 1);
      const resolution = at > 0 && dateText.trim().length > 0 ? legacyPerson(paragraph.text.slice(0, at), dateText, index) : null;
      return { field, paragraph, dateText, before: at > 0 ? paragraph.text.slice(0, at) : "", resolution };
    });
    const usable = plans.filter((p) => p.resolution);
    let converted = 0;
    let unresolved = 0;
    for (let from = 0; from < usable.length; from += 40) {
      const chunk = usable.slice(from, from + 40);
      // New date first, then the old field: a name that cannot be found loses nothing.
      const searches = chunk.map((plan) => {
        const literal = (plan.resolution as LegacyResolution).literal;
        const results = plan.paragraph.search(searchText(literal), { matchCase: true, matchWholeWord: false });
        results.load("items");
        return { plan, results, nth: occurrenceBefore(plan.before, literal, (plan.resolution as LegacyResolution).start) };
      });
      await context.sync();
      for (const { plan, results, nth } of searches) {
        const target = results.items[nth];
        if (!target) continue; // field kept as it is; reported as skipped
        const inserted = target.insertText(plan.dateText, "After");
        const control = inserted.insertContentControl();
        const person = (plan.resolution as LegacyResolution).person;
        control.tag = tagFor(person);
        control.title = person ? `Vefeyât: ${person.shortName}` : "Vefeyât: kimliği belirsiz";
        control.appearance = "Hidden";
        plan.field.delete();
        converted++;
        if (!person) unresolved++;
      }
      await context.sync();
      progress(Math.min(from + 40, usable.length), usable.length);
    }
    return { converted, unresolved, skipped: plans.length - converted };
  });
}

/**
 * The first of `rows` (mentions of one person, in reading order) that lies
 * before the selection in the document, asked of Word by location rather
 * than matched by text. A footnote or endnote counts at its reference mark.
 */
export async function firstMentionBefore(segments: readonly Segment[], rows: readonly InventoryRow[]): Promise<InventoryRow | null> {
  if (rows.length === 0) return null;
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items");
    await context.sync();
    const probes = rows.slice(0, 40).map((row) => {
      const segment = segments[row.segment];
      const host = segment ? paragraphs.items[segment.anchor] : undefined;
      if (!segment || !host || row.text.length > 250) return null;
      if (segment.part === "main") {
        const results = host.getRange("Whole").search(searchText(row.text), { matchCase: true, matchWholeWord: false });
        results.load("items");
        return { row, segment, results, notes: null };
      }
      const foot = host.getRange("Whole").footnotes;
      const end = host.getRange("Whole").endnotes;
      foot.load("items");
      end.load("items");
      return { row, segment, results: null, notes: { foot, end } };
    });
    await context.sync();
    const relations = probes.map((probe) => {
      if (!probe) return null;
      const range = probe.results
        ? probe.results.items[occurrenceBefore(probe.segment.text, probe.row.text, probe.row.start)]
        : probe.notes
          ? noteOf(probe.notes, probe.segment.order)?.reference
          : undefined;
      return range ? { row: probe.row, relation: range.compareLocationWith(selection) } : null;
    });
    await context.sync();
    return relations.find((r) => r && BEFORE.has(r.relation.value))?.row ?? null;
  });
}

// ---- Yazım denetimi (WordApiDesktop 1.3: Range.hasNoProofing) ----

export function canSkipProofing(): boolean {
  return Office.context.requirements.isSetSupported("WordApiDesktop", "1.3");
}

/**
 * Marks every listed mention as exempt from spelling and grammar checking.
 * Word then shows its ordinary text menu on a right click there (with the
 * Vefeyât commands), not the spelling menu that add-ins cannot extend.
 */
export async function skipProofing(segments: readonly Segment[], rows: readonly InventoryRow[], progress: Progress = () => {}): Promise<number> {
  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items");
    await context.sync();
    let done = 0;
    let marked = 0;
    for (let from = 0; from < rows.length; from += CHUNK) {
      const chunk = rows.slice(from, from + CHUNK);
      const probes = chunk.map((row) => {
        const segment = segments[row.segment];
        const host = segment ? paragraphs.items[segment.anchor] : undefined;
        if (!segment || !host || row.text.length > 250) return null;
        if (segment.part === "main") {
          const results = host.getRange("Whole").search(searchText(row.text), { matchCase: true, matchWholeWord: false });
          results.load("items");
          return { row, segment, results, notes: null };
        }
        const foot = host.getRange("Whole").footnotes;
        const end = host.getRange("Whole").endnotes;
        foot.load("items");
        end.load("items");
        return { row, segment, results: null, notes: { foot, end } };
      });
      await context.sync();
      const noteSearches = probes.map((probe) => {
        if (!probe?.notes) return null;
        const note = noteOf(probe.notes, probe.segment.order);
        if (!note) return null;
        const results = note.body.search(searchText(probe.row.text), { matchCase: true, matchWholeWord: false });
        results.load("items");
        return results;
      });
      await context.sync();
      probes.forEach((probe, k) => {
        if (!probe) return;
        const results = probe.results ?? noteSearches[k];
        const range = results?.items[occurrenceBefore(probe.segment.text, probe.row.text, probe.row.start)];
        if (!range) return;
        range.hasNoProofing = true;
        marked++;
      });
      await context.sync();
      done += chunk.length;
      progress(done, rows.length);
    }
    return marked;
  });
}

// ---- Belge okumasının saklanması (WordApi 1.6: paragraf olayları) ----

interface ReadCache {
  segments: Segment[];
  footnotes: boolean;
  /** uniqueLocalId of each main-text paragraph, by anchor. */
  ids: string[];
}

let cache: ReadCache | null = null;
let readAt = 0;
/** A full read at least this often, in case a change raised no event (notes, other stories). */
const MAX_AGE = 10 * 60_000;
let watching = false;
const dirty = new Set<string>();
let structural = true;
/** Bumped whenever the reading changes, so derived results (the inventory) know to recompute. */
let readVersion = 0;

export function documentVersion(): number {
  return readVersion;
}

/** Forget the stored reading (the next read is a full one). */
export function forgetDocument(): void {
  structural = true;
}

/**
 * Starts listening to paragraph changes. Returns false where Word has no
 * such events: every read is then a full one.
 */
export async function watchDocument(): Promise<boolean> {
  if (watching) return true;
  if (!Office.context.requirements.isSetSupported("WordApi", "1.6")) return false;
  try {
    await Word.run(async (context) => {
      context.document.onParagraphChanged.add(async (event) => {
        for (const id of event.uniqueLocalIds) dirty.add(id);
      });
      context.document.onParagraphAdded.add(async () => {
        structural = true;
      });
      context.document.onParagraphDeleted.add(async () => {
        structural = true;
      });
      await context.sync();
    });
    watching = true;
  } catch {
    watching = false;
  }
  return watching;
}

async function fullRead(progress: Progress): Promise<ReadCache> {
  const read = await readSegments(progress);
  let ids: string[] = [];
  if (watching) {
    ids = await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items/uniqueLocalId");
      await context.sync();
      return paragraphs.items.map((p) => p.uniqueLocalId);
    });
  }
  return { segments: read.segments, footnotes: read.footnotes, ids };
}

/**
 * The document's segments, read once and then kept up to date: only changed
 * paragraphs (with their footnotes and endnotes) are read again; added or
 * deleted paragraphs, or a change Vefeyât cannot place, mean a full read.
 */
export async function readDocument(progress: Progress = () => {}): Promise<{ segments: Segment[]; footnotes: boolean }> {
  const live = await watchDocument();
  if (!live || !cache || structural || cache.ids.length === 0 || Date.now() - readAt > MAX_AGE) {
    dirty.clear();
    structural = false;
    cache = await fullRead(progress);
    readAt = Date.now();
    readVersion++;
    return cache;
  }
  if (dirty.size === 0) return cache;
  const changed = [...dirty];
  dirty.clear();
  const anchors = changed.map((id) => (cache as ReadCache).ids.indexOf(id));
  if (anchors.some((a) => a < 0)) {
    // A footnote paragraph or one Vefeyât has not seen: read everything again.
    cache = await fullRead(progress);
    readVersion++;
    return cache;
  }
  const fresh = await Word.run(async (context) => {
    const items = changed.map((id) => {
      const paragraph = context.document.getParagraphByUniqueLocalId(id);
      paragraph.load("text");
      const foot = paragraph.getRange("Whole").footnotes;
      const end = paragraph.getRange("Whole").endnotes;
      if (cache?.footnotes) {
        foot.load("items/body/text");
        end.load("items/body/text");
      }
      return { paragraph, foot, end };
    });
    await context.sync();
    return items.map(({ paragraph, foot, end }) => ({
      text: paragraph.text,
      foot: cache?.footnotes ? foot.items.map((n) => n.body.text) : [],
      end: cache?.footnotes ? end.items.map((n) => n.body.text) : [],
    }));
  });
  const replaced = new Set(anchors);
  const kept = cache.segments.filter((s) => !replaced.has(s.anchor));
  anchors.forEach((anchor, k) => {
    const f = fresh[k] as { text: string; foot: string[]; end: string[] };
    kept.push({ part: "main", anchor, order: 0, text: f.text });
    f.foot.forEach((text, i) => kept.push({ part: "note", anchor, order: i + 1, text }));
    f.end.forEach((text, i) => kept.push({ part: "note", anchor, order: ENDNOTE_BASE + i + 1, text }));
  });
  kept.sort((a, b) => a.anchor - b.anchor || a.order - b.order);
  cache = { ...cache, segments: kept };
  readVersion++;
  return cache;
}

/**
 * Toplu düzeltme: replaces the date written right after the selected name
 * (`old`, as the inventory read it) with `value`. The selection must be the
 * name. Returns false when the date is not found right after it.
 */
export async function replaceDateAfterSelection(old: string, value: string): Promise<boolean> {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    const rest = selection.getRange("End").expandTo(selection.paragraphs.getFirst().getRange("End"));
    const found = rest.search(searchText(old), { matchCase: true, matchWholeWord: false });
    found.load("items/text");
    await context.sync();
    const target = found.items[0];
    if (!target) return false;
    // Only the date directly after the name, not one later in the paragraph.
    const gap = selection.getRange("End").expandTo(target.getRange("Start"));
    gap.load("text");
    await context.sync();
    if (gap.text.trim().length > 0) return false;
    target.insertText(value, "Replace");
    await context.sync();
    return true;
  });
}

/** Whether the selection is inside a footnote or an endnote (WordApi 1.3: Body.type). */
export async function selectionInNote(): Promise<boolean> {
  return Word.run(async (context) => {
    const body = context.document.getSelection().parentBody;
    body.load("type");
    await context.sync();
    return body.type === "Footnote" || body.type === "Endnote";
  });
}

/**
 * İlk geçiş bakımı: removes the date written right after one mention — a
 * Vefeyât date with its control, a hand-written one with the space before it.
 * Nothing is removed unless that exact date stands directly after the name.
 */
export async function removeDateAfter(segment: Segment, start: number, literal: string, dateText: string): Promise<boolean> {
  if (!dateText || literal.length > 250) return false;
  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items");
    await context.sync();
    const host = paragraphs.items[segment.anchor];
    if (!host) return false;
    let scope: Word.Body | Word.Range = host.getRange("Whole");
    if (segment.part === "note") {
      const foot = host.getRange("Whole").footnotes;
      const end = host.getRange("Whole").endnotes;
      foot.load("items");
      end.load("items");
      await context.sync();
      const note = noteOf({ foot, end }, segment.order);
      if (!note) return false;
      scope = note.body;
    }
    const names = scope.search(searchText(literal), { matchCase: true, matchWholeWord: false });
    names.load("items");
    await context.sync();
    const name = names.items[occurrenceBefore(segment.text, literal, start)];
    if (!name) return false;
    const rest = name.getRange("End").expandTo(scope.getRange("End"));
    const dates = rest.search(searchText(dateText), { matchCase: true, matchWholeWord: false });
    dates.load("items");
    await context.sync();
    const date = dates.items[0];
    if (!date) return false;
    const gap = name.getRange("End").expandTo(date.getRange("Start"));
    gap.load("text");
    const control = date.parentContentControlOrNullObject;
    control.load("tag");
    await context.sync();
    if (gap.text.trim().length > 0) return false; // not directly after the name
    if (!control.isNullObject && control.tag?.startsWith("vefeyat:date:")) {
      control.cannotDelete = false;
      control.delete(false);
    } else {
      gap.expandTo(date).delete();
    }
    await context.sync();
    return true;
  });
}

/** Whether the document changed since it was last read (only known with paragraph events). */
export function hasPendingChanges(): boolean {
  return watching && cache !== null && (dirty.size > 0 || structural);
}

/** Whether the document has been read once since the pane opened (with paragraph events). */
export function documentRead(): boolean {
  return watching && cache !== null;
}
