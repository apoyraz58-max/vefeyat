/**
 * Behaviour of the ported core beyond the VBA self-checks. Expected values
 * follow the VBA 0.3.41 logic step by step (see comments where not obvious).
 */
import { describe, expect, it } from "vitest";
import {
  Catalog,
  CatalogError,
  alreadyHasDate,
  articleUrl,
  chronoYear,
  collapseLetters,
  dateKind,
  dateNote,
  identity,
  nameKey,
  nameQuery,
  nameScore,
  nearWord,
  normalize,
  parseCatalogCsv,
  requirePrefix,
  singleSafe,
  startsWithDate,
  textName,
  type Person,
} from "../src/core/index.js";

const HEADER = "Kısa Ad;Tam Ad;Orijinal Ad;Vefat Tarihi;Tanıtım;Kaynak Link";
const DIA = "https://islamansiklopedisi.org.tr/";

describe("normalize", () => {
  it.each([
    ["Gazzâlî", "gazzali"],
    ["GAZZÂLÎ", "gazzali"],
    ["İbn Sînâ", "ibn sina"],
    ["IŞIK", "isik"],
    ["Fahreddin er-Râzî", "fahreddin er razi"],
    ["İbni Haldûn", "ibn haldun"],
    ["Kıvâmü’s-sünne", "kivamu s sunne"],
    ["Ebû'l-Kâsım", "ebu l kasim"],
    ["Ḥasan", "hasan"],
    ["Ṣalāḥ", "salah"],
    ["  Ali  b. Hasan ", "ali b hasan"],
    ["Æthelred", "aethelred"],
    ["الغَزَالِي", "الغزالي"],
    ["", ""],
  ])("%s -> %s", (input, expected) => {
    expect(normalize(input)).toBe(expected);
  });

  it("collapses repeated letters for spelling keys", () => {
    expect(collapseLetters("gazzali")).toBe("gazali");
    expect(nameKey("Gazzâlî")).toBe("gazali");
    expect(nameKey("Ebû Hanîfe")).toBe("ebuhanife");
  });
});

describe("nameQuery: Turkish case endings after an apostrophe", () => {
  it.each([
    ["Gazzâlî'nin", "Gazzâlî"],
    ["Râzî’ye", "Râzî"],
    ["Ali'nin", "Ali"],
    ["İbn Sînâ'ya", "İbn Sînâ"],
    ["Mâverdî'den", "Mâverdî"],
    ["Şâfiî'ninki", "Şâfiî"],
    ["Ebu'l-Kasım", "Ebu'l-Kasım"], // "l-Kasım" is not an ending
    ["Kıvâmü's-sünne", "Kıvâmü's-sünne"],
    ["Mes'ûdî", "Mes'ûdî"], // circumflex: part of the name
    ["Ca'fer", "Ca'fer"],
    ["Gazzâlî", "Gazzâlî"],
  ])("%s -> %s", (input, expected) => {
    expect(nameQuery(input)).toBe(expected);
  });
});

describe("textName", () => {
  it("turns catalog order into reading order", () => {
    expect(textName("Râzî, Fahreddin")).toBe("Fahreddin Râzî");
  });
  it("leaves other values unchanged", () => {
    expect(textName("Gazzâlî")).toBe("Gazzâlî");
    expect(textName("A, B, C")).toBe("A, B, C");
    expect(textName("Ali (Hasan), Zeyd")).toBe("Ali (Hasan), Zeyd");
    expect(textName(", Ali")).toBe(", Ali");
  });
});

describe("name scores", () => {
  it("edit distance", () => {
    expect(nearWord("gazali", "gazzali", 1)).toBe(true);
    expect(nearWord("razi", "gazi", 1)).toBe(true);
    expect(nearWord("abcdef", "ghijkl", 1)).toBe(false);
    expect(nearWord("hamid", "hamit", 1)).toBe(true);
  });
  it("short spelling of a doubled-letter name is a weak match", () => {
    // "gazali" is not "gazzali", but its collapsed form is: score 1 (chooser).
    expect(nameScore("Gazzâlî", "Gazali")).toBe(1);
  });
  it("suffix is ignored, accents are folded", () => {
    expect(nameScore("Gazzâlî", "Gazzâlî'nin")).toBe(4);
    expect(nameScore("Gazzâlî", "GAZZALI")).toBe(4);
  });
  it("every query token inside a longer name", () => {
    expect(nameScore("Fahreddin er-Râzî", "Râzî")).toBe(2);
    expect(nameScore("Muhammed b. İsmâil el-Buhârî", "Buhârî Muhammed")).toBe(2);
  });
  it("words in another order with one near spelling", () => {
    expect(nameScore("Ebû Hâmid Muhammed el-Gazzâlî", "Gazzali Hamit")).toBe(1);
  });
  it("unrelated names", () => {
    expect(nameScore("Ebû Hanîfe", "Mâlik")).toBe(0);
  });
});

const p = (shortName: string, fullName: string, deathDate: string, source: string, description = ""): Person => ({
  shortName,
  fullName,
  originalName: "",
  deathDate,
  description,
  source,
});

describe("Catalog.find", () => {
  const catalog = new Catalog([
    p("Ebû Bekir er-Râzî", "Ebû Bekir Muhammed b. Zekeriyyâ er-Râzî", "(ö. 313/925)", `${DIA}razi-ebu-bekir`),
    p("Fahreddin er-Râzî", "Ebû Abdillâh Fahreddîn Muhammed b. Ömer er-Râzî", "(ö. 606/1210)", `${DIA}fahreddin-er-razi`),
    p("Ebû Hâtim er-Râzî", "Ebû Hâtim Ahmed b. Hamdân er-Râzî", "(ö. 322/934)", `${DIA}ebu-hatim-er-razi`),
    p("Zeyd", "Zeyd b. Ali", "(ö. 122/740)", `${DIA}zeyd-b-ali`, "Râzî ile karıştırılır"),
    p("Gazzâlî", "Ebû Hâmid Muhammed b. Muhammed el-Gazzâlî", "(ö. 505/1111)", `${DIA}gazzali`),
  ]);

  it("well-known bearer first, then catalog order, description matches last", () => {
    expect(catalog.find("Râzî").map((x) => x.shortName)).toEqual([
      "Fahreddin er-Râzî",
      "Ebû Bekir er-Râzî",
      "Ebû Hâtim er-Râzî",
      "Zeyd",
    ]);
  });
  it("Turkish ending is removed before the lookup", () => {
    expect(catalog.find("Râzî'nin")[0]?.shortName).toBe("Fahreddin er-Râzî");
  });
  it("description matches can be excluded", () => {
    expect(catalog.find("Râzî", false).map((x) => x.shortName)).not.toContain("Zeyd");
  });
  it("falls back to words in any order", () => {
    expect(catalog.find("Hamit Gazzali").map((x) => x.shortName)).toEqual(["Gazzâlî"]);
  });
  it("empty query finds nothing", () => {
    expect(catalog.find("  ")).toEqual([]);
  });
  it("single strong candidate needs no chooser", () => {
    const gazzali = catalog.find("Gazzâlî");
    expect(singleSafe(gazzali, "Gazzâlî")).toBe(true);
    expect(singleSafe(gazzali, "Gazali")).toBe(false);
  });
});

describe("DİA identity", () => {
  it("canonical article address", () => {
    expect(articleUrl(" https://www.islamansiklopedisi.org.tr/gazzali ")).toBe(`${DIA}gazzali`);
    expect(articleUrl("https://example.org/gazzali")).toBe("");
    expect(articleUrl(`${DIA}gazz ali`)).toBe("");
  });
  it("identity uses the article, else the person key", () => {
    expect(identity(p("Gazzâlî", "", "", `${DIA}Gazzali/#bolum`))).toBe(`DIA:${DIA}gazzali`);
    expect(identity(p("Gazzâlî", "", "", ""))).toMatch(/^Vfp1:/);
  });
});

describe("parseCatalogCsv", () => {
  it("reads BOM, CRLF, quoted separators, doubled quotes and line breaks", () => {
    const csv = [
      `﻿${HEADER}`,
      `Gazzâlî;"Ebû Hâmid; el-Gazzâlî";الغزالي;(ö. 505/1111);"Eş'arî ""kelâmcı""\nve sûfî";${DIA}gazzali`,
      `  Râzî  ;---;;(ö. 606/1210);;${DIA}fahreddin-er-razi`,
      "",
    ].join("\r\n");
    const people = parseCatalogCsv(csv);
    expect(people).toHaveLength(2);
    expect(people[0]).toEqual({
      shortName: "Gazzâlî",
      fullName: "Ebû Hâmid; el-Gazzâlî",
      originalName: "الغزالي",
      deathDate: "(ö. 505/1111)",
      description: 'Eş\'arî "kelâmcı"\nve sûfî',
      source: `${DIA}gazzali`,
    });
    expect(people[1]?.shortName).toBe("Râzî");
    expect(people[1]?.fullName).toBe("");
  });

  const fails = (csv: string, message: RegExp): void => {
    expect(() => parseCatalogCsv(csv)).toThrow(CatalogError);
    expect(() => parseCatalogCsv(csv)).toThrow(message);
  };
  it("rejects wrong headers", () => fails("Ad;Tam Ad;Orijinal Ad;Vefat Tarihi;Tanıtım;Kaynak Link\nA;;;;;", /line 1: The six CSV headers/));
  it("rejects a wrong column count with its line", () => fails(`${HEADER}\nA;;;;;\nB;;;;`, /line 3: Expected exactly six/));
  it("counts lines inside quoted fields", () => fails(`${HEADER}\nA;;;;"x\ny";\nB;;`, /line 4: Expected exactly six/));
  it("rejects a missing short name", () => fails(`${HEADER}\n;B;;;;`, /short name is required/));
  it("rejects a quote inside an unquoted field", () => fails(`${HEADER}\nA";;;;;`, /Quote inside an unquoted field/));
  it("rejects text after a closing quote", () => fails(`${HEADER}\n"A"x;;;;;`, /Unexpected text after a closing quote/));
  it("rejects an unclosed quote", () => fails(`${HEADER}\n"A;;;;;`, /Unclosed quoted field/));
  it("rejects an empty catalog", () => fails(HEADER, /contains no people/));
  it("rejects a file without header", () => fails("", /no header/));
  it("rejects two dates for one person", () =>
    fails(`${HEADER}\nA;B;;(ö. 1);;u\nA;B;;(ö. 2);x;u`, /conflicting dates/));
  it("allows the same person twice with the same date", () => {
    expect(parseCatalogCsv(`${HEADER}\nA;B;;(ö. 1);;u\nA;B;;(ö. 1);x;u`)).toHaveLength(2);
  });
});

describe("dates", () => {
  it("prefix setting", () => {
    expect(requirePrefix(" V ")).toBe("v");
  });
  it.each([
    [" (ö. 505/1111) diyor", true],
    ["(öl. 505)", true],
    ["(Ö. 505)", true],
    ["(v. ?)", true],
    ["(505/1111)", true],
    ["(1111)", true],
    ["( 1100 - 1111 )", true],
    ["(12)", false], // numbered comment, not a date
    ["(1111 yılında)", false],
    ["ve sonra", false],
  ])("startsWithDate(%j) = %s", (text, expected) => {
    expect(startsWithDate(text)).toBe(expected);
  });

  it.each([
    [" (ö. 505/1111) ve", "(ö. 505/1111)", true],
    [" (v. 505/1111).", "(ö. 505/1111)", true],
    [" (Ö.  505/1111)", "(ö. 505/1111)", true],
    [" (ö. 505/1112)", "(ö. 505/1111)", false],
    [" (ö. 505/1111)x", "(ö. 505/1111)", false],
    ["\r(ö. 505/1111)", "(ö. 505/1111)", false],
  ])("alreadyHasDate(%j)", (after, raw, expected) => {
    expect(alreadyHasDate(after, raw)).toBe(expected);
  });

  it.each([
    ["505/1111", "hijriGregorian"],
    ["500/1106'dan sonra", "hijriGregorian"],
    ["505/1111; kaynak: 506", "hijriGregorian"],
    ["1111", "single"],
    ["1100-1111", "range"],
    ["1111?", "other"],
    ["ilk yarısı", null],
    ["", null],
  ])("dateKind(%j)", (body, expected) => {
    expect(dateKind(body)).toBe(expected);
  });

  it("chronology year", () => {
    expect(chronoYear("(ö. 505/1111)")).toEqual({ year: 1111, certain: true });
    expect(chronoYear("(v. 500/1106'dan sonra)")).toEqual({ year: 1106, certain: true });
    expect(chronoYear("(ö. 1111)")).toEqual({ year: 1111, certain: false });
    expect(chronoYear("(m.ö. 384-322)")).toEqual({ year: -384, certain: true });
    expect(chronoYear("(tarih bilinmiyor)")).toBeNull();
  });

  it("date notes", () => {
    expect(dateNote("(v. 500/1106'dan sonra)")).toMatch(/^Belirsizlik/);
    expect(dateNote("(ö. ?)")).toMatch(/^Tarih belirsizlik/);
    expect(dateNote("(ö. 505/1111)")).toBe("\r\nKaynak takvimleri arasında yeni bir dönüşüm hesaplanmadı.");
    expect(dateNote("(ö. 1111)")).toBe("");
  });
});
