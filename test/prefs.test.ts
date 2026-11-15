import { describe, expect, it } from "vitest";
import { bulkSuggestions } from "../src/addin/bulk.js";
import { buildInventory } from "../src/addin/inventory.js";
import { DEFAULT_SETTINGS } from "../src/addin/logic.js";
import { exportPrefs, mergePrefs, parsePrefs, preferenceOf, withPreferred } from "../src/addin/prefs.js";
import { Catalog, type Person, personKey } from "../src/core/index.js";

const gazzali: Person = { shortName: "Gazzâlî", fullName: "Ebû Hâmid Muhammed b. Muhammed el-Gazzâlî", originalName: "", deathDate: "(ö. 505/1111)", description: "", source: "https://islamansiklopedisi.org.tr/gazzali" };
const catalog = new Catalog([gazzali]);
const seg = (text: string) => [{ part: "main" as const, anchor: 0, order: 0, text }];

describe("preferred spellings", () => {
  const prefs = { [personKey(gazzali)]: "İmam Gazzâlî" };
  it("preference of a selection drops the ending", () => {
    expect(preferenceOf("Gazzâlî'nin")).toBe("Gazzâlî");
  });
  it("the catalog spelling becomes a spelling warning, the preferred one a match", () => {
    const rows = buildInventory(seg("Gazzâlî (ö. 505/1111) der."), catalog, undefined, {}, [], prefs);
    expect(rows[0]!.status).toBe("spelling");
    const suggestion = bulkSuggestions(rows, DEFAULT_SETTINGS, prefs);
    expect(suggestion.map((s) => s.value)).toEqual(["İmam Gazzâlî"]);
  });
  it("options put the preference first", () => {
    expect(withPreferred([{ value: "a", label: "x" }, { value: "b", label: "y" }], "b").map((o) => o.value)).toEqual(["b", "a"]);
  });
});

describe("export and import", () => {
  it("round trip with validation", () => {
    const text = exportPrefs({ ...DEFAULT_SETTINGS, prefix: "v" }, { k: "Gazzâlî" }, ["kuran"], { m: "p" });
    const parsed = parsePrefs(text);
    expect(parsed.ok && parsed.file.settings.prefix).toBe("v");
    expect(parsed.ok && parsed.file.exclusions).toEqual(["kuran"]);
  });
  it("rejects other files and bad values", () => {
    expect(parsePrefs("{").ok).toBe(false);
    expect(parsePrefs('{"format":"x"}').ok).toBe(false);
    const bad = parsePrefs('{"format":"vefeyat-tercihler","version":1,"settings":{"prefix":"zz","scope":"all"}}');
    expect(bad.ok && bad.file.settings).toEqual({ scope: "all" });
  });
  it("merges lists, imported values win", () => {
    const file = parsePrefs(exportPrefs(DEFAULT_SETTINGS, { a: "Yeni" }, ["x"], {}));
    if (!file.ok) throw new Error();
    expect(mergePrefs({ spellings: { a: "Eski", b: "B" }, exclusions: ["x", "y"], confirmations: {} }, file.file)).toEqual({
      spellings: { a: "Yeni", b: "B" },
      exclusions: ["x", "y"],
      confirmations: {},
    });
  });
});
