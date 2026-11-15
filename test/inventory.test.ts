import { describe, expect, it } from "vitest";
import { buildInventory, catalogLabel, followingDate, spellingMatches, statusLabel, summarize, type Segment } from "../src/addin/inventory.js";
import { mentionKey } from "../src/addin/logic.js";
import { Catalog, type Person, personKey } from "../src/core/index.js";

const DIA = "https://islamansiklopedisi.org.tr/";
const p = (shortName: string, fullName: string, deathDate: string, slug: string): Person => ({
  shortName,
  fullName,
  originalName: "",
  deathDate,
  description: "",
  source: DIA + slug,
});
const gazzali = p("Gazzâlî", "Ebû Hâmid Muhammed b. Muhammed el-Gazzâlî", "(ö. 505/1111)", "gazzali");
const catalog = new Catalog([
  gazzali,
  p("Fahreddin er-Râzî", "Ebû Abdillâh Fahreddîn Muhammed b. Ömer er-Râzî", "(ö. 606/1210)", "fahreddin-er-razi"),
  p("Ebû Bekir er-Râzî", "Ebû Bekir Muhammed b. Zekeriyyâ er-Râzî", "(ö. 313/925)", "razi-ebu-bekir"),
  p("Cüveynî", "Ebü'l-Meâlî Abdülmelik el-Cüveynî", "", "cuveyni"),
]);

const main = (anchor: number, text: string): Segment => ({ part: "main", anchor, order: 0, text });
const note = (anchor: number, order: number, text: string): Segment => ({ part: "note", anchor, order, text });

describe("followingDate", () => {
  it.each([
    [" (ö. 505/1111) bu", "(ö. 505/1111)"],
    [" (v. 606/1210).", "(v. 606/1210)"],
    [" (505/1111) der", "(505/1111)"],
    [" (12) bu", ""],
    [" bu görüşü", ""],
  ])("%j", (after, expected) => {
    expect(followingDate(after)).toBe(expected);
  });
});

describe("buildInventory", () => {
  const segments = [
    main(0, "Gazali bu görüşü savunur."),
    main(1, "Gazzâlî (ö. 505/1111) ise farklı düşünür. Râzî de katılır."),
    note(1, 1, "Gazzâlî, el-Müstasfâ, 2/34."),
    note(1, 2, "Bu konuda Cüveynî de durur."),
    main(2, "Sonra Gazzâlî'nin talebeleri gelir."),
  ];
  const rows = buildInventory(segments, catalog);

  it("lists mentions in reading order, footnotes after their paragraph", () => {
    expect(rows.map((r) => [r.text, r.part])).toEqual([
      ["Gazali", "main"],
      ["Gazzâlî", "main"],
      ["Râzî", "main"],
      ["Cüveynî", "note"],
      ["Gazzâlî'nin", "main"],
    ]);
  });
  it("skips a name inside a footnote citation", () => {
    expect(rows.some((r) => r.part === "note" && r.text === "Gazzâlî")).toBe(false);
  });
  it("weak spelling needs verification and is not counted as the first mention", () => {
    expect(rows[0]).toMatchObject({ status: "near", person: null, first: false });
    expect(catalogLabel(rows[0]!)).toBe("Gazzâlî (?)");
  });
  it("first certain mention, its date and later mentions", () => {
    expect(rows[1]).toMatchObject({ status: "match", first: true, date: "(ö. 505/1111)" });
    expect(rows[4]).toMatchObject({ status: "match", first: false, date: "" });
  });
  it("homonyms need a choice", () => {
    expect(rows[2]).toMatchObject({ status: "choose", person: null });
    expect(catalogLabel(rows[2]!)).toBe("2 aday; seçim gerekli");
  });
  it("status hints", () => {
    expect(statusLabel(rows[1]!)).toBe("Kaynak yazımıyla uyumlu");
    expect(statusLabel(rows[3]!)).toBe("Kaynak yazımıyla uyumlu · Kaynakta tarih yok");
  });
  it("summary", () => {
    expect(summarize(rows)).toEqual({ mentions: 5, people: 2, firstWithoutDate: 0, spelling: 0, uncertain: 2 });
  });
  it("first mention without date is reported", () => {
    const only = buildInventory([main(0, "Gazzâlî der ki.")], catalog);
    expect(statusLabel(only[0]!)).toBe("Kaynak yazımıyla uyumlu · İlk geçişte tarih yok");
    expect(summarize(only).firstWithoutDate).toBe(1);
  });
});

describe("confirmations in the inventory", () => {
  it("a confirmed homonym counts as a certain person", () => {
    const text = "Bu konuda Râzî de katılır.";
    const ebuBekir = catalog.people[2]!;
    const key = mentionKey("Râzî", "Bu konuda ", " de katılır.");
    const rows = buildInventory([main(0, text)], catalog, undefined, { [key]: personKey(ebuBekir) });
    expect(rows[0]).toMatchObject({ person: ebuBekir, confirmed: true, first: true });
    expect(statusLabel(rows[0]!)).toContain("onaylı kimlik");
  });
});

describe("spellingMatches", () => {
  it("catalog spellings and their parts", () => {
    expect(spellingMatches(gazzali, "Gazzâlî'nin")).toBe(true);
    expect(spellingMatches(gazzali, "Gazali")).toBe(false);
  });
});
