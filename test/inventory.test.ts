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
  it("a variant spelling is settled by the certain mention later in the document", () => {
    expect(rows[0]).toMatchObject({ status: "spelling", person: gazzali, first: true, resolved: "context" });
    expect(statusLabel(rows[0]!)).toContain("kimlik belgedeki diğer geçişlerden");
    const alone = buildInventory([main(0, "Gazali bu görüşü savunur.")], catalog);
    expect(alone[0]).toMatchObject({ status: "near", person: null, first: false });
    expect(catalogLabel(alone[0]!)).toBe("Gazzâlî (?)");
  });
  it("certain mention, its date and later mentions", () => {
    expect(rows[1]).toMatchObject({ status: "match", first: false, date: "(ö. 505/1111)", resolved: "catalog" });
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
    expect(summarize(rows)).toEqual({ mentions: 5, people: 2, firstWithoutDate: 1, spelling: 1, uncertain: 1, dateIssues: 0 });
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

describe("document exclusions", () => {
  it("a phrase marked Şahıs değil is skipped with its endings", async () => {
    const { exclusionKey } = await import("../src/addin/logic.js");
    const rows = buildInventory([main(0, "Gazzâlî der ki. Gazzâlî'nin talebeleri.")], catalog, undefined, {}, [exclusionKey("Gazzâlî")]);
    expect(rows).toEqual([]);
  });
});

describe("identity from the document", () => {
  const ebuBekir = catalog.people[2]!;
  const fahreddin = catalog.people[1]!;
  it("the date written after the name picks the bearer", () => {
    const rows = buildInventory([main(0, "Râzî (ö. 313/925) bu görüşü savunur.")], catalog);
    expect(rows[0]).toMatchObject({ person: ebuBekir, resolved: "date" });
  });
  it("two bearers in the document: the last one named comes first, the choice stays open", () => {
    const rows = buildInventory(
      [main(0, "Ebû Bekir er-Râzî tıpta meşhurdur."), main(1, "Fahreddin er-Râzî tefsir yazdı."), main(2, "Râzî burada farklı düşünür.")],
      catalog,
    );
    expect(rows.map((r) => r.person?.shortName ?? null)).toEqual(["Ebû Bekir er-Râzî", "Fahreddin er-Râzî", null]);
    expect(rows[2]!.candidates[0]).toBe(fahreddin);
  });
});

describe("endnotes", () => {
  it("are read like footnotes, after them, and labelled Sonnot", async () => {
    const { ENDNOTE_BASE, placeOf } = await import("../src/addin/inventory.js");
    const segments: Segment[] = [main(0, "Giriş."), note(0, ENDNOTE_BASE + 1, "Cüveynî de bu görüştedir."), note(0, 1, "Gazzâlî de bu görüştedir.")];
    const rows = buildInventory(segments, catalog);
    expect(rows.map((r) => [r.text, placeOf(segments[r.segment])])).toEqual([
      ["Gazzâlî", "Dipnot"],
      ["Cüveynî", "Sonnot"],
    ]);
  });
});

describe("lookups ignore the Arabic article", () => {
  it("Ebû Bekir er-Râzî is the catalog's Râzî, Ebû Bekir", async () => {
    const { lookup } = await import("../src/addin/logic.js");
    const inverted = new Catalog([
      { shortName: "Râzî, Ebû Bekir", fullName: "Ebû Bekr Muhammed b. Zekeriyyâ er-Râzî", originalName: "", deathDate: "(ö. 313/925)", description: "", source: DIA + "razi-ebu-bekir" },
      { shortName: "Fahreddin er-Râzî", fullName: "", originalName: "", deathDate: "(ö. 606/1210)", description: "", source: DIA + "fahreddin" },
      { shortName: "İbn Sînâ", fullName: "Ebû Alî el-Hüseyn b. Abdillâh b. Sînâ", originalName: "", deathDate: "(ö. 428/1037)", description: "", source: DIA + "ibn-sina" },
      { shortName: "İbn Sinân Hafâcî", fullName: "", originalName: "", deathDate: "(ö. 466/1074)", description: "", source: DIA + "ibn-sinan" },
    ]);
    expect(lookup(inverted, "Ebû Bekir er-Râzî").automatic?.source).toBe(DIA + "razi-ebu-bekir");
    expect(lookup(inverted, "İbn Sînâ'nın").automatic?.source).toBe(DIA + "ibn-sina");
  });
});

describe("written dates in the text", () => {
  it("a date that differs from the catalog, or whose calendars disagree, is reported and audited", async () => {
    const { hasDateIssue } = await import("../src/addin/inventory.js");
    const { auditRows, DEFAULT_SETTINGS } = await import("../src/addin/logic.js");
    const rows = buildInventory([main(0, "Gazzâlî (ö. 505/1111) der. Sonra Gazzâlî (ö. 506/1112) yine der."), main(1, "Ebû Bekir er-Râzî (ö. 313/1025) hekimdi.")], catalog);
    expect(rows.map((r) => r.dateCheck)).toEqual(["ok", "different", "inconsistent"]);
    expect(statusLabel(rows[1]!)).toContain("metindeki tarih katalogdan farklı (katalog: (ö. 505/1111))");
    expect(statusLabel(rows[2]!)).toContain("hicrî ve milâdî yıl birbirini tutmuyor");
    expect(summarize(rows).dateIssues).toBe(2);
    expect(auditRows(rows, DEFAULT_SETTINGS).map((r) => r.text)).toEqual(["Gazzâlî", "Gazzâlî", "Ebû Bekir er-Râzî"]);
    expect(hasDateIssue(rows[0]!)).toBe(false);
  });
});

describe("homonyms by the subject of the sentence", () => {
  it("both bearers named: tefsir and tıp tell them apart; no subject keeps the choice", () => {
    const rows = buildInventory(
      [
        main(0, "Fahreddin er-Râzî tefsir yazdı. Ebû Bekir er-Râzî hekimdi."),
        main(1, "Râzî tefsirinde bu meseleyi tartışır."),
        main(2, "Râzî tıpta deneysel yöntemi kullanmıştır."),
        main(3, "Râzî burada farklı düşünür."),
      ],
      new Catalog([
        p("Fahreddin er-Râzî", "", "(ö. 606/1210)", "fahreddin-er-razi"),
        { ...p("Ebû Bekir er-Râzî", "", "(ö. 313/925)", "razi-ebu-bekir"), description: "Ünlü hekim ve filozof." },
      ].map((x, k) => (k === 0 ? { ...x, description: "Kelâm, felsefe, tefsir ve usûl-i fıkıh alanlarında Eş‘arî âlimi." } : x))),
    );
    expect(rows.slice(2).map((r) => [r.person?.shortName ?? null, r.resolved, r.topic ?? null])).toEqual([
      ["Fahreddin er-Râzî", "topic", "tefsir"],
      ["Ebû Bekir er-Râzî", "topic", "tıp"],
      [null, null, null],
    ]);
  });
});
