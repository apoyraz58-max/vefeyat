import { describe, expect, it } from "vitest";
import { candidateLabel, dateTag, earlierSpellings, lookup, matchHint, planDate } from "../src/addin/logic.js";
import { Catalog, nameOptions, personKey, sourceName, withEnding, type Person } from "../src/core/index.js";

const DIA = "https://islamansiklopedisi.org.tr/";
const p = (shortName: string, fullName: string, deathDate: string, slug: string, originalName = "", description = ""): Person => ({
  shortName,
  fullName,
  originalName,
  deathDate,
  description,
  source: DIA + slug,
});

const gazzali = p("Gazzâlî", "Ebû Hâmid Muhammed b. Muhammed el-Gazzâlî", "(ö. 505/1111)", "gazzali", "الغزالي");
const fahreddin = p("Fahreddin er-Râzî", "Ebû Abdillâh Fahreddîn Muhammed b. Ömer er-Râzî", "(ö. 606/1210)", "fahreddin-er-razi");
const ebuBekir = p("Ebû Bekir er-Râzî", "Ebû Bekir Muhammed b. Zekeriyyâ er-Râzî", "(ö. 313/925)", "razi-ebu-bekir");
const kivam = p("Kıvâmüssünne", "Ebü'l-Kāsım Kıvâmü's-sünne İsmâil b. Muhammed", "(ö. 535/1140)", "kivamussunne");
const catalog = new Catalog([
  gazzali,
  fahreddin,
  ebuBekir,
  kivam,
  p("Zeyd", "Zeyd b. Ali", "(ö. 122/740)", "zeyd", "", "Râzî ile karıştırılır"),
]);

describe("lookup", () => {
  it("one strong candidate is chosen automatically", () => {
    const result = lookup(catalog, "Gazzâlî'nin");
    expect(result.query).toBe("Gazzâlî");
    expect(result.automatic).toEqual(gazzali);
  });
  it("homonyms need a choice; well-known bearer first; descriptions ignored", () => {
    const result = lookup(catalog, "Râzî");
    expect(result.automatic).toBeNull();
    expect(result.candidates.map((x) => x.shortName)).toEqual(["Fahreddin er-Râzî", "Ebû Bekir er-Râzî"]);
    expect(result.strong.size).toBe(0);
  });
  it("explains why a candidate is not exact", () => {
    const razi = lookup(catalog, "Râzî");
    expect(razi.candidates.map((x) => razi.scores.get(personKey(x)))).toEqual([2, 2]);
    expect(matchHint(2)).toBe("adın bir parçası");
    const weak = lookup(catalog, "Gazali");
    expect([...weak.scores.values()]).toEqual([1]);
    expect(matchHint(1)).toBe("benzer yazım, doğrulayın");
    expect(matchHint(4)).toBe("");
  });
  it("a weak spelling is never automatic", () => {
    const result = lookup(catalog, "Gazali");
    expect(result.automatic).toBeNull();
    expect(result.candidates).toEqual([gazzali]);
  });
  it("unknown name", () => {
    expect(lookup(catalog, "Mâlik").candidates).toEqual([]);
  });
});

describe("planDate", () => {
  it("formats the catalog date with the chosen marker", () => {
    expect(planDate(gazzali, "v", " bu görüşü savunur.")).toEqual({ ok: true, text: "(v. 505/1111)" });
    expect(planDate(gazzali, "original", "")).toEqual({ ok: true, text: "(ö. 505/1111)" });
  });
  it("refuses when a date already follows the name", () => {
    expect(planDate(gazzali, "o", " (ö. 505/1111) bu").ok).toBe(false);
    expect(planDate(gazzali, "o", " (v. 505/1111)").ok).toBe(false);
    expect(planDate(gazzali, "o", " (505/1111)").ok).toBe(false);
  });
  it("a numbered remark is not a date", () => {
    expect(planDate(gazzali, "o", " (12) bu").ok).toBe(true);
  });
  it("refuses a person without a usable date", () => {
    expect(planDate(p("X", "", "", "x"), "o", "").ok).toBe(false);
  });
});

describe("spelling", () => {
  it("recovers the catalog spelling of the selected part", () => {
    expect(sourceName(gazzali, "Gazali")).toBe("Gazzâlî");
    expect(sourceName(kivam, "Kivamussunne")).toBe("Kıvâmüssünne");
    expect(sourceName(fahreddin, "Razi")).toBe("Râzî");
  });
  it("offers distinct spellings, closest first", () => {
    expect(nameOptions(gazzali, "Gazali'nin")).toEqual([
      { value: "Gazzâlî", label: "Seçili ada yakın yazım" },
      { value: "Ebû Hâmid Muhammed b. Muhammed el-Gazzâlî", label: "Tam / uzun ad" },
      { value: "الغزالي", label: "Orijinal ad" },
    ]);
  });
  it("keeps the Turkish ending of the selection", () => {
    expect(withEnding("Gazali'nin", "Gazzâlî")).toBe("Gazzâlî'nin");
    expect(withEnding("Gazali", "Gazzâlî")).toBe("Gazzâlî");
  });
});

describe("earlier mention search", () => {
  it("searches the selected spelling, catalog names and the family name", () => {
    expect(earlierSpellings(fahreddin, "Fahreddin Râzî'nin")).toEqual([
      "Fahreddin Râzî",
      "Fahreddin er-Râzî",
      "Ebû Abdillâh Fahreddîn Muhammed b. Ömer er-Râzî",
      "Râzî",
    ]);
  });
  it("skips very short values", () => {
    expect(earlierSpellings(p("Ali", "", "", "ali"), "Ali")).toEqual(["Ali"]);
    expect(earlierSpellings(p("Ab", "", "", "ab"), "Ab")).toEqual([]);
  });
});

describe("tags and labels", () => {
  it("date tag is stable per person and short", () => {
    expect(dateTag(gazzali)).toBe(dateTag({ ...gazzali, deathDate: "(ö. 506/1112)" }));
    expect(dateTag(gazzali)).not.toBe(dateTag(fahreddin));
    expect(dateTag(gazzali)).toMatch(/^vefeyat:date:[0-9a-f]{8}$/);
  });
  it("candidate label", () => {
    expect(candidateLabel(gazzali)).toBe("Gazzâlî (ö. 505/1111)");
    expect(candidateLabel(p("Râzî, Fahreddin", "", "", "r"))).toBe("Fahreddin Râzî");
  });
});
