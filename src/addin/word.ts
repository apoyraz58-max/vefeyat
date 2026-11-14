/**
 * Thin Word layer (Office.js, WordApi 1.3). Every decision is made in
 * logic.ts; these functions only read or write the document.
 */
import { type SettingName, type Settings, settingValue } from "./logic.js";

export interface SelectionText {
  /** Selected text as Word returns it. */
  readonly text: string;
  /** Rest of the paragraph after the selection. */
  readonly following: string;
}

export async function readSelection(): Promise<SelectionText> {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    const paragraphEnd = selection.paragraphs.getLast().getRange("End");
    const following = selection.getRange("End").expandTo(paragraphEnd);
    following.load("text");
    await context.sync();
    return { text: selection.text, following: following.text };
  });
}

/** Writes " <date>" right after the selection inside a hidden, tagged content control. */
export async function insertDate(dateText: string, tag: string, title: string): Promise<void> {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const inserted = selection.getRange("End").insertText(` ${dateText}`, "After");
    const control = inserted.insertContentControl();
    control.tag = tag;
    control.title = title;
    control.appearance = "Hidden";
    // Continue typing after the date, outside the control.
    control.getRange("After").select("Start");
    await context.sync();
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
    if (earlier && select) {
      earlier.range.select("Select");
      await context.sync();
    }
    return earlier !== undefined;
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
