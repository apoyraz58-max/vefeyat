import { describe, expect, it } from "vitest";
import { addHistory, applyOrder, bulkSuggestions, firstUseSuggestions, parseBookmark, parseHistory, resumeIndex, rowKey } from "../src/addin/bulk.js";
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

describe("first-use upkeep (main text and notes apart)", () => {
  const seg = (anchor: number, text: string, part: "main" | "note" = "main", order = 0) => ({ part, anchor, order, text });
  it("a date on a later mention moves to the new first one; extra dates are dropped", () => {
    const rows3 = buildInventory([seg(0, "Gazzâlî burada anılır."), seg(1, "Sonra Gazzâlî (ö. 505/1111) der ki."), seg(2, "Yine Gazzâlî (v. 505/1111) der.")], catalog);
    const list = firstUseSuggestions(rows3, DEFAULT_SETTINGS);
    expect(list.map((s) => [s.kind, s.row.segment, s.from?.segment, s.value])).toEqual([
      ["move", 0, 1, "(ö. 505/1111)"],
      ["drop", 2, 2, ""],
    ]);
    // The receiving first mention gets no separate "add a date".
    expect(bulkSuggestions(rows3, DEFAULT_SETTINGS).filter((s) => s.kind === "date")).toEqual([]);
  });
  it("notes keep their own first mention; dates the catalog contradicts are left alone", () => {
    const rows3 = buildInventory([seg(0, "Gazzâlî (ö. 505/1111) der."), seg(0, "Gazzâlî (ö. 505/1111) de böyle der.", "note", 1), seg(1, "Sonra Gazzâlî (ö. 506/1112) der.")], catalog);
    expect(firstUseSuggestions(rows3, DEFAULT_SETTINGS)).toEqual([]);
  });
  it("only with Yalnız ilk kullanım", () => {
    const rows3 = buildInventory([seg(0, "Gazzâlî der."), seg(1, "Gazzâlî (ö. 505/1111) der.")], catalog);
    expect(firstUseSuggestions(rows3, { ...DEFAULT_SETTINGS, scope: "all" })).toEqual([]);
  });
});
