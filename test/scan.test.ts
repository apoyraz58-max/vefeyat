/**
 * Scanner behaviour on the challenge categories of the VBA academic gold set
 * (variants, suffixes, schools, citations, ordinary words), with a small
 * synthetic catalog. Expected values follow the VBA 0.3.41 rules.
 */
import { describe, expect, it } from "vitest";
import {
  NameIndex,
  citationSpans,
  excludeMention,
  findMatches,
  locators,
  tokenize,
  units,
  type Person,
} from "../src/core/index.js";

const DIA = "https://islamansiklopedisi.org.tr/";
const p = (shortName: string, fullName: string, deathDate: string, slug: string): Person => ({
  shortName,
  fullName,
  originalName: "",
  deathDate,
  description: "",
  source: DIA + slug,
});

const index = new NameIndex([
  p("Gazzâlî", "Ebû Hâmid Muhammed b. Muhammed el-Gazzâlî", "(ö. 505/1111)", "gazzali"),
  p("Fahreddin er-Râzî", "Ebû Abdillâh Fahreddîn Muhammed b. Ömer er-Râzî", "(ö. 606/1210)", "fahreddin-er-razi"),
  p("Ebû Bekir er-Râzî", "Ebû Bekir Muhammed b. Zekeriyyâ er-Râzî", "(ö. 313/925)", "razi-ebu-bekir"),
  p("Şâfiî", "Ebû Abdillâh Muhammed b. İdrîs eş-Şâfiî", "(ö. 204/820)", "safii"),
  p("Aynî", "Ebû Muhammed Bedreddîn Mahmûd b. Ahmed el-Aynî", "(ö. 855/1451)", "ayni-bedreddin"),
  p("Cüveynî", "Ebü'l-Meâlî Abdülmelik b. Abdillâh el-Cüveynî", "(ö. 478/1085)", "cuveyni-imamul-harameyn"),
]);

const found = (text: string, strict = true): string[] => findMatches(text, index, strict).map((m) => m.text);

describe("tokenize", () => {
  it("keeps hyphens and apostrophes inside a token", () => {
    expect(tokenize("Fahreddin er-Râzî'nin tefsiri.").map((t) => t.key)).toEqual(["fahredin", "erazinin", "tefsiri"]);
  });
  it("offsets are UTF-16 positions", () => {
    const t = tokenize("  Gazzâlî der");
    expect(t[0]).toMatchObject({ start: 2, length: 7 });
  });
});

describe("findMatches: variants and suffixes", () => {
  it.each([
    ["Gazzâlî bu görüşü savunur.", ["Gazzâlî"]],
    ["Gazali bu görüşü savunur.", ["Gazali"]],
    ["GAZZÂLÎ bu görüşü savunur.", ["GAZZÂLÎ"]],
    ["Gazzâlî'nin eserinde geçer.", ["Gazzâlî'nin"]],
    ["Bu görüş Gazzâlî’ye aittir.", ["Gazzâlî’ye"]],
    ["Fahreddin er-Râzî tefsirinde bunu açıklar.", ["Fahreddin er-Râzî"]],
    ["Râzî ve Cüveynî farklı düşünür.", ["Râzî", "Cüveynî"]],
  ])("%s", (text, expected) => {
    expect(found(text)).toEqual(expected);
  });

  it("does not take common words or lowercase single words as names in strict mode", () => {
    expect(found("Tarih ve fıkıh usulü üzerine.")).toEqual([]);
    expect(found("gazali kelimesi")).toEqual([]);
  });
  it("lenient mode finds lowercase spellings too", () => {
    expect(found("gazali kelimesi", false)).toEqual(["gazali"]);
  });
  it("a longer capitalised name run that is not in the catalog is not split", () => {
    // "Ahmed Gazzâlî" (the brother) is one name run; the short name alone must not match.
    expect(found("Ahmed Gazzâlî bunu yazdı.")).toEqual([]);
  });
});

describe("context rules", () => {
  it("school adjective is not the person", () => {
    expect(found("Şâfiî mezhebinde bu mesele tartışılır.")).toEqual([]);
    expect(found("Hanefî ve Şâfiî fakihleri ayrılır.")).toEqual([]);
  });
  it("a person predicate keeps Şâfiî as a person", () => {
    expect(found("Şâfiî der ki bu hüküm sabittir.")).toEqual(["Şâfiî"]);
  });
  it("ordinary 'aynı' is not Aynî", () => {
    expect(found("Aynı eserde şöyle denir.")).toEqual([]);
    expect(found("Aynî der ki hadis sahihtir.")).toEqual(["Aynî"]);
  });
  it("'… adlı eserinde' refers to a work", () => {
    expect(excludeMention("Gazzâlî adlı eserinde", 0, 7)).toBe(true);
  });
  it("author–title citation is not a person mention", () => {
    expect(found("Gazzâlî, el-Müstasfâ, 2/34.")).toEqual([]);
    expect(found("Râzî, Mefâtîhu'l-gayb, Beyrut 1990, s. 12.")).toEqual([]);
  });
  it("a narrative sentence with a comma is still a mention", () => {
    expect(found("Gazzâlî, el-Müstasfâ adlı eserinde der ki usul önemlidir.")).toEqual(["Gazzâlî"]);
  });
  it("citation check can be switched off", () => {
    expect(findMatches("Gazzâlî, el-Müstasfâ, 2/34.", index, true, false).map((m) => m.text)).toEqual(["Gazzâlî"]);
  });
});

describe("ranking and shapes", () => {
  it("well-known bearer first for a shared short name", () => {
    expect(index.rank("razi", "Râzî").map((x) => x.shortName)).toEqual(["Fahreddin er-Râzî", "Ebû Bekir er-Râzî"]);
  });
  it("strict mode accepts a key only with a word count seen in the catalog", () => {
    expect(index.shapeAllowed("fahredinerazi", 2)).toBe(true);
    expect(index.shapeAllowed("fahredinerazi", 3)).toBe(false);
  });
});

describe("citation islands", () => {
  it("sentence units ignore abbreviations and initials", () => {
    const text = "Bk. s. 12 vd. Gazzâlî der. İkinci cümle.";
    // "Bk.", "s." and "vd." are abbreviations; only the period after "der" ends a unit.
    expect(units(text).map(([a, b]) => text.slice(a, b))).toEqual(["Bk. s. 12 vd. Gazzâlî der", " İkinci cümle"]);
  });
  it("locators: volume/page, page marker, bare page and year", () => {
    expect(locators("el-Müstasfâ, 2/34").map((l) => l[2])).toEqual(["page"]);
    expect(locators("Mefâtîh, s. 12-14").map((l) => l[2])).toEqual(["marked"]);
    expect(locators("Mefâtîh, 45.").map((l) => l[2])).toEqual(["bare"]);
    expect(locators("Beyrut: Dâru'l-kütüb, 1990").map((l) => l[2])).toContain("year");
  });
  it("covers the author and title of a footnote citation", () => {
    const text = "Bk. Râzî, Mefâtîhu'l-gayb, 3/45.";
    const start = text.indexOf("Râzî");
    const spans = citationSpans(text, [[start, 4, "Râzî"]]);
    expect(spans.length).toBe(1);
    const [a, b] = spans[0]!;
    expect(a).toBeLessThanOrEqual(start);
    expect(text.slice(a, b)).toContain("3/45");
  });
  it("author–year references", () => {
    const text = "Gazzâlî (1990: 12) bunu tartışır.";
    const spans = citationSpans(text, [[0, 7, "Gazzâlî"]]);
    expect(spans).toEqual([[0, 18]]);
  });
  it("plain narrative has no citation span", () => {
    const text = "Gazzâlî bu görüşü savunur.";
    expect(citationSpans(text, [[0, 7, "Gazzâlî"]])).toEqual([]);
  });
});
