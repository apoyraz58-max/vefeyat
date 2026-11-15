import { describe, expect, it } from "vitest";
import { DEFAULT_SETTINGS, addConfirmation, auditRows, candidateLabel, confirmedPerson, mentionKey, parseConfirmations, dateTag, earlierSpellings, lookup, matchHint, orderByNameForm, parseLaunch, planDate, quickInsertBlocker, scopeRule, settingValue, settingsSummary } from "../src/addin/logic.js";
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

describe("ribbon launch", () => {
  it.each([
    ["?action=insert", { view: "secili", action: "insert", tool: null }],
    ["?action=spelling", { view: "secili", action: "spelling", tool: null }],
    ["?view=dokum", { view: "dokum", action: null, tool: null }],
    ["?view=denetim", { view: "secili", action: "audit", tool: null }],
    ["?tool=dizin", { view: "teslim", action: null, tool: "dizin" }],
    ["?view=katalog", { view: "katalog", action: null, tool: null }],
    ["?view=yardim", { view: "yardim", action: null, tool: null }],
    ["?view=ayarlar", { view: "katalog", action: null, tool: null }],
    ["?view=ara", { view: "ara", action: null, tool: null }],
    ["?action=dia", { view: "secili", action: "dia", tool: null }],
    ["?soon=shading", { view: "yakinda", action: null, tool: null, soon: "shading" }],
    ["?view=kimlik", { view: "kimlik", action: null, tool: null }],
    ["?view=devam", { view: "secili", action: "resume", tool: null }],
    ["?view=toplu", { view: "toplu", action: null, tool: null }],
    ["?view=saglik", { view: "saglik", action: null, tool: null }],
    ["?view=kaynak", { view: "kaynak", action: null, tool: null }],
    ["?soon=bilinmeyen", { view: "secili", action: null, tool: null }],
    ["", { view: "secili", action: null, tool: null }],
    ["?view=bilinmeyen&action=sil", { view: "secili", action: null, tool: null }],
  ])("%j", (search, expected) => {
    expect(parseLaunch(search)).toEqual(expected);
  });

  it("quick insert only for an unambiguous selection", () => {
    expect(quickInsertBlocker(false, "Gazzâlî", null)).toMatch(/Katalog yüklenemedi/);
    expect(quickInsertBlocker(true, null, null)).toMatch(/şahıs adı seçip/);
    expect(quickInsertBlocker(true, "Mâlik", lookup(catalog, "Mâlik"))).toMatch(/eşleşen şahıs yok/);
    expect(quickInsertBlocker(true, "Râzî", lookup(catalog, "Râzî"))).toMatch(/Birden fazla aday/);
    expect(quickInsertBlocker(true, "Gazali", lookup(catalog, "Gazali"))).toMatch(/benzer yazım/);
    expect(quickInsertBlocker(true, "Gazzâlî", lookup(catalog, "Gazzâlî"))).toBeNull();
  });
});

describe("settings", () => {
  it("defaults and unknown values", () => {
    expect(settingValue("prefix", undefined)).toBe("o");
    expect(settingValue("prefix", "v")).toBe("v");
    expect(settingValue("scope", "sil")).toBe("first");
    expect(settingValue("nameForm", "full")).toBe("full");
    expect(settingValue("auditMode", 3)).toBe("all");
  });
  it("scope rules", () => {
    expect(scopeRule("first")).toEqual({ allowed: true, checkEarlier: true });
    expect(scopeRule("all")).toEqual({ allowed: true, checkEarlier: false });
    expect(scopeRule("none").allowed).toBe(false);
  });
  it("summary", () => {
    expect(settingsSummary(DEFAULT_SETTINGS)).toBe("Kısaltma: ö. · Yalnız ilk kullanım · Ad biçimi: Her seferinde sor");
  });
  it("preferred spelling first", () => {
    const options = nameOptions(gazzali, "Gazali");
    expect(orderByNameForm(options, "ask").map((o) => o.label)[0]).toBe("Seçili ada yakın yazım");
    expect(orderByNameForm(options, "full").map((o) => o.label)[0]).toBe("Tam / uzun ad");
    expect(orderByNameForm(options, "original").map((o) => o.value)[0]).toBe("الغزالي");
  });
});

describe("audit rows", () => {
  const rows = [
    { person: {}, first: true, date: "(ö. 1)" },
    { person: {}, first: false, date: "" },
    { person: null, first: false, date: "" },
  ];
  it("first-use scope skips verified repeats", () => {
    expect(auditRows(rows, DEFAULT_SETTINGS)).toEqual([rows[0], rows[2]]);
  });
  it("undated mode skips dated mentions", () => {
    expect(auditRows(rows, { ...DEFAULT_SETTINGS, scope: "all", auditMode: "undated" })).toEqual([rows[1], rows[2]]);
  });
});

describe("identity confirmations", () => {
  it("mention key uses nearby normalized words", () => {
    expect(mentionKey("Râzî", "Bu konuda İmam ", " de der ki")).toBe("bu konuda imam|Râzî|de der ki");
    expect(mentionKey("Râzî", "uzun bir giriş cümlesi var. Bu konuda İmam ", " de")).toBe("giris cumlesi var bu konuda imam|Râzî|de");
    // The date inserted after the name does not change the key.
    expect(mentionKey("Râzî", "Bu konuda İmam ", " (ö. 313/925) de der ki")).toBe(mentionKey("Râzî", "Bu konuda İmam ", " de der ki"));
  });
  it("round trip, limit and lookup", () => {
    let map = addConfirmation({}, "a|Râzî|b", "P1");
    map = addConfirmation(map, "c|Râzî|d", "P2");
    expect(parseConfirmations(JSON.stringify(map))).toEqual(map);
    expect(parseConfirmations("bozuk")).toEqual({});
    const fahreddin = lookup(catalog, "Râzî").candidates[0]!;
    map = addConfirmation(map, "x|Râzî|y", personKey(fahreddin));
    expect(confirmedPerson(map, "x|Râzî|y", lookup(catalog, "Râzî").candidates)).toEqual(fahreddin);
    expect(confirmedPerson(map, "a|Râzî|b", lookup(catalog, "Râzî").candidates)).toBeNull();
  });
});
