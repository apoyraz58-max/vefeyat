import { describe, expect, it } from "vitest";
import { addHistory, applyOrder, bulkSuggestions, parseBookmark, parseHistory, resumeIndex, rowKey } from "../src/addin/bulk.js";
import { buildInventory, type Segment } from "../src/addin/inventory.js";
import { DEFAULT_SETTINGS } from "../src/addin/logic.js";
import { Catalog, type Person } from "../src/core/index.js";

const p = (shortName: string, fullName: string, deathDate: string, slug: string): Person => ({ shortName, fullName, originalName: "", deathDate, description: "", source: `https://islamansiklopedisi.org.tr/${slug}` });
const catalog = new Catalog([
  p("Gazzâlî", "Ebû Hâmid Muhammed b. Muhammed el-Gazzâlî", "(ö. 505/1111)", "gazzali"),
  p("Cüveynî", "Ebü'l-Meâlî Abdülmelik el-Cüveynî", "(ö. 478/1085)", "cuveyni"),
]);
const segments: Segment[] = [
  { part: "main", anchor: 0, order: 0, text: "Gazzâlî der ki. Cüveynî (ö. 478/1085) hocasıdır." },
  { part: "main", anchor: 1, order: 0, text: "Sonra Gazzâlî'nin talebeleri gelir." },
];
const rows = buildInventory(segments, catalog);

describe("bulkSuggestions", () => {
  it("first-mention dates only, with the prefix setting", () => {
    const list = bulkSuggestions(rows, { ...DEFAULT_SETTINGS, prefix: "v" });
    expect(list.map((s) => [s.kind, s.row.text, s.value])).toEqual([["date", "Gazzâlî", "(v. 505/1111)"]]);
  });
  it("all mentions with Bütün kullanımlar, none with Yalnız denetle", () => {
    expect(bulkSuggestions(rows, { ...DEFAULT_SETTINGS, scope: "all" }).map((s) => s.row.text)).toEqual(["Gazzâlî", "Gazzâlî'nin"]);
    expect(bulkSuggestions(rows, { ...DEFAULT_SETTINGS, scope: "none" })).toEqual([]);
  });
  it("applies from the end of the document", () => {
    const list = bulkSuggestions(rows, { ...DEFAULT_SETTINGS, scope: "all" });
    expect(applyOrder(list).map((s) => s.row.text)).toEqual(["Gazzâlî'nin", "Gazzâlî"]);
  });
});

describe("audit bookmark", () => {
  it("resumes at the same mention, else the same position", () => {
    const key = rowKey(rows[2]!, segments);
    expect(resumeIndex(rows, segments, { at: 0, key, total: 3 })).toBe(2);
    expect(resumeIndex(rows, segments, { at: 9, key: "x", total: 3 })).toBe(2);
    expect(resumeIndex([], segments, { at: 0, key, total: 3 })).toBe(-1);
  });
  it("parses only valid bookmarks", () => {
    expect(parseBookmark('{"at":1,"key":"k","total":2}')).toEqual({ at: 1, key: "k", total: 2 });
    expect(parseBookmark("{bad")).toBeNull();
    expect(parseBookmark(null)).toBeNull();
  });
});

describe("history", () => {
  it("newest first, survives a round trip", () => {
    let list = addHistory([], { kind: "date", text: "a" }, new Date(2026, 0, 1, 9, 5));
    list = addHistory(list, { kind: "spelling", text: "b", control: 3 });
    expect(list.map((e) => e.text)).toEqual(["b", "a"]);
    expect(list[1]!.time).toBe("09:05");
    expect(parseHistory(JSON.stringify(list))).toEqual(list);
    expect(parseHistory("nope")).toEqual([]);
  });
});

describe("wrong written dates", () => {
  it("offered unchecked, applied before a respelling of the same name", () => {
    const rows2 = buildInventory([{ part: "main", anchor: 0, order: 0, text: "Gazzâlî (ö. 506/1112) der ki." }], catalog);
    const list = bulkSuggestions(rows2, { ...DEFAULT_SETTINGS, prefix: "o" });
    expect(list.map((s) => [s.kind, s.old ?? null, s.value, s.preselected])).toEqual([["fixdate", "(ö. 506/1112)", "(ö. 505/1111)", false]]);
  });
});
