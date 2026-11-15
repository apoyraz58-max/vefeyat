/**
 * Thin Word layer (Office.js, WordApi 1.3). Every decision is made in
 * logic.ts; these functions only read or write the document.
 */
import type { InventoryRow, Segment } from "./inventory.js";
import { parseSpellings } from "./prefs.js";
import { DATE_TAG_PREFIX, type DateControl } from "./health.js";
import { INDEX_CODE, isVefeyatIndexField, xeCode } from "./submit.js";
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
    const notes = context.document.body.footnotes;
    notes.load("items/body/text");
    await context.sync();
    const lowered = spellings.map((value) => value.toLocaleLowerCase("tr"));
    const containing = notes.items.filter((note) => {
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

/** Main-text paragraphs and, where WordApi 1.5 is available, footnotes. */
/** Word requests queued per sync in long operations. */
const CHUNK = 200;

export type Progress = (done: number, total: number) => void;

export async function readSegments(): Promise<{ segments: Segment[]; footnotes: boolean }> {
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
          const collection = p.getRange("Whole").footnotes;
          collection.load("items/body/text");
          return collection;
        });
        await context.sync();
        notes.forEach((collection, k) => {
          collection.items.forEach((note, i) => segments.push({ part: "note", anchor: from + k, order: i + 1, text: note.body.text }));
        });
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
      const notes = host.getRange("Whole").footnotes;
      notes.load("items");
      await context.sync();
      const note = notes.items[segment.order - 1];
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
    const lists = [context.document.body.contentControls];
    if (Office.context.requirements.isSetSupported("WordApi", "1.5")) {
      const notes = context.document.body.footnotes;
      notes.load("items");
      await context.sync();
      for (const note of notes.items) lists.push(note.body.contentControls);
    }
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
  for (const key of [CONFIRMATIONS_KEY, AUDIT_KEY, EXCLUSIONS_KEY, SPELLINGS_KEY, ...Object.values(SETTING_KEYS)]) {
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
  const notes = context.document.body.footnotes;
  notes.load("items");
  await context.sync();
  const lists = [context.document.body.fields, ...notes.items.map((n) => n.body.fields)];
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

/**
 * Şahıslar dizini: replaces Vefeyât's index marks with one XE field before
 * every certain mention and puts the INDEX field at the end of the document
 * (or where the previous index stood). Returns the number of marks.
 */
export async function buildIndex(segments: readonly Segment[], rows: readonly InventoryRow[], progress: Progress = () => {}): Promise<{ marks: number; missing: number }> {
  return Word.run(async (context) => {
    // Where the old index was, so the new one goes back there.
    const blocks = context.document.body.contentControls.getByTag(INDEX_TAG);
    blocks.load("items");
    const oldIndex = context.document.body.fields;
    oldIndex.load("items/code");
    await context.sync();
    const previous = blocks.items[0] ?? oldIndex.items.find((f) => /^\s*INDEX\b/i.test(f.code) && isVefeyatIndexField(f.code));
    const where = previous ? ("result" in previous ? previous.result : previous.getRange("Whole")) : null;
    const anchor = where ? where.paragraphs.getFirst().insertParagraph("", "Before") : null;
    await deleteIndexBlock(context);
    const removed = await deleteIndexFields(context);
    if (removed.indexes > 0 && blocks.items.length === 0) await deleteOldHeadings(context);

    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items");
    await context.sync();
    const noteLists = new Map<number, Word.NoteItemCollection>();
    const anchors = [...new Set(rows.filter((r) => segments[r.segment]?.part === "note").map((r) => (segments[r.segment] as Segment).anchor))];
    for (let from = 0; from < anchors.length; from += CHUNK) {
      for (const anchor of anchors.slice(from, from + CHUNK)) {
        const host = paragraphs.items[anchor];
        if (!host) continue;
        const list = host.getRange("Whole").footnotes;
        list.load("items");
        noteLists.set(anchor, list);
      }
      await context.sync();
    }
    // From the end of the document backwards: a mark inserted before a name
    // never shifts the mentions still to be found before it.
    const ordered = [...rows].reverse();
    let marks = 0;
    for (let from = 0; from < ordered.length; from += CHUNK) {
      const searches: Array<{ row: InventoryRow; results: Word.RangeCollection; nth: number }> = [];
      for (const row of ordered.slice(from, from + CHUNK)) {
        const segment = segments[row.segment];
        if (!segment || row.text.length > 250) continue;
        const host = paragraphs.items[segment.anchor];
        if (!host) continue;
        const scope: Word.Body | Word.Range | undefined =
          segment.part === "note" ? noteLists.get(segment.anchor)?.items[segment.order - 1]?.body : host.getRange("Whole");
        if (!scope) continue;
        const results = scope.search(searchText(row.text), { matchCase: true, matchWholeWord: false });
        results.load("items");
        searches.push({ row, results, nth: occurrenceBefore(segment.text, row.text, row.start) });
      }
      await context.sync();
      for (const { row, results, nth } of searches) {
        const target = results.items[nth];
        if (!target || !row.person) continue;
        target.insertField("Before", "XE", xeCode(row.person));
        marks++;
      }
      await context.sync();
      progress(Math.min(from + CHUNK, ordered.length), ordered.length);
    }

    // The block: heading and INDEX field in one tagged control.
    let first: Word.Paragraph;
    first = anchor ?? context.document.body.insertParagraph("", "End");
    first.insertText(INDEX_HEADING, "Replace");
    first.styleBuiltIn = "Heading1";
    const holder = first.insertParagraph("", "After");
    holder.styleBuiltIn = "Normal";
    const control = first.getRange("Whole").expandTo(holder.getRange("Whole")).insertContentControl();
    control.tag = INDEX_TAG;
    control.title = "Vefeyât: şahıslar dizini";
    control.appearance = "BoundingBox";
    control.insertBreak("Page", "Start"); // inside the block: removing it leaves no blank page
    const index = holder.getRange("Start").insertField("Start", "Index", INDEX_CODE);
    await context.sync();
    index.updateResult();
    await context.sync();
    return { marks, missing: rows.length - marks };
  });
}

// ---- Belge sağlığı ve Tarihin kaynağı ----

/** Content-control lists of the body and (WordApi 1.5) the footnotes, loaded. */
async function controlLists(context: Word.RequestContext): Promise<Array<{ part: "main" | "note"; list: Word.ContentControlCollection }>> {
  const lists: Array<{ part: "main" | "note"; list: Word.ContentControlCollection }> = [{ part: "main", list: context.document.body.contentControls }];
  if (Office.context.requirements.isSetSupported("WordApi", "1.5")) {
    const notes = context.document.body.footnotes;
    notes.load("items");
    await context.sync();
    for (const note of notes.items) lists.push({ part: "note", list: note.body.contentControls });
  }
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

/** What Vefeyât scans (body, footnotes) and what it does not (endnotes, headers, comments). */
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
