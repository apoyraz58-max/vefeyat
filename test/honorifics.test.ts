import { describe, expect, it } from "vitest";
import { Catalog, NameIndex, fuseLinks, honorificAliases, type Person } from "../src/core/index.js";
import { lookup } from "../src/addin/logic.js";
import { buildInventory } from "../src/addin/inventory.js";

const p = (shortName: string, fullName: string, deathDate: string): Person => ({ shortName, fullName, originalName: "", deathDate, description: "", source: "x" + shortName });
const catalog = new Catalog([
  p("Pezdevî, Ebü’l-Usr", "Ebü’l-Hasen Ebü’l-Usr Fahrü’l-İslâm Alî b. Muhammed b. el-Hüseyn b. Abdilkerîm el-Pezdevî", "(ö. 482/1089)"),
  p("Pezdevî, Ebü’l-Yüsr", "Sadrü’l-İslâm Ebü’l-Yüsr Muhammed b. Muhammed b. el-Hüseyn b. Abdilkerîm el-Pezdevî", "(ö. 493/1100)"),
  p("Gazzâlî", "Hüccetü’l-İslâm Ebû Hâmid Muhammed b. Muhammed b. Muhammed b. Ahmed el-Gazzâlî et-Tûsî", "(ö. 505/1111)"),
]);

describe("lakap + nisba, joined or apart", () => {
  it("helpers", () => {
    expect(fuseLinks("Fahrü’l-İslâm el-Pezdevî")).toBe("Fahrülİslâm el-Pezdevî");
    expect(honorificAliases(catalog.people[0]!.fullName)).toContain("Fahrü’l-İslâm el-Pezdevî");
    expect(honorificAliases(catalog.people[0]!.fullName)).toContain("Fahrü’l-İslâm Ebü’l-Usr el-Pezdevî");
  });
  it("selected name: the right Pezdevî is found and chosen", () => {
    for (const alias of ["Fahrülislâm el-Pezdevî", "Fahrülislam Pezdevî", "Fahrü'l-İslâm Pezdevî", "Fahrülislâm Pezdevî"]) {
      const found = lookup(catalog, alias);
      expect(found.automatic?.shortName, alias).toBe("Pezdevî, Ebü’l-Usr");
    }
    expect(lookup(catalog, "Hüccetülislâm Gazzâlî").automatic?.shortName).toBe("Gazzâlî");
    expect(lookup(catalog, "Sadrülislâm Ebü'l-Yüsr el-Pezdevî").candidates[0]?.shortName).toBe("Pezdevî, Ebü’l-Yüsr");
  });
  it("running text", () => {
    const text = "Fahrülislâm Pezdevî bu konuda ayrışır. Sadrülislâm Ebü'l-Yüsr el-Pezdevî de katılır. Fahrülislâm Ebü'l-Usr el-Pezdevî ise böyle der. Hüccetülislâm Gazzâlî de.";
    const rows = buildInventory([{ part: "main", anchor: 0, order: 0, text }], catalog, new NameIndex(catalog.people));
    expect(rows.map((r) => [r.text, r.person?.shortName])).toEqual([
      ["Fahrülislâm Pezdevî", "Pezdevî, Ebü’l-Usr"],
      ["Sadrülislâm Ebü'l-Yüsr el-Pezdevî", "Pezdevî, Ebü’l-Yüsr"],
      ["Fahrülislâm Ebü'l-Usr el-Pezdevî", "Pezdevî, Ebü’l-Usr"],
      ["Hüccetülislâm Gazzâlî", "Gazzâlî"],
    ]);
  });
});

describe("running text: ayn, eş- and article-less nisbas", () => {
  const cat = new Catalog([
    p("Ebû Mutî‘ el-Belhî", "Ebû Mutî‘ el-Hakem b. Abdillâh b. Mesleme el-Belhî", "(ö. 199/814)"),
    p("A‘lem eş-Şentemerî", "Ebü’l-Haccâc Yûsuf b. Süleymân b. Îsâ el-A‘lem eş-Şentemerî", "(ö. 476/1083)"),
    p("Ebû Ca‘fer en-Nüfeylî", "Ebû Ca‘fer Abdullāh b. Muhammed b. Alî en-Nüfeylî", "(ö. 234/848)"),
    p("Ebü’l-Abbâs es-Seffâh", "Ebü’l-Abbâs Abdullāh b. Muhammed b. Alî es-Seffâh", "(ö. 136/754)"),
  ]);
  const index = new NameIndex(cat.people);
  it.each([
    ["Ebû Mutî‘ el-Belhî", "Ebû Mutî‘ el-Belhî"],
    ["A‘lem eş-Şentemerî", "A‘lem eş-Şentemerî"],
    ["A‘lem Şentemerî", "A‘lem eş-Şentemerî"],
    ["Ebû Ca‘fer Nüfeylî", "Ebû Ca‘fer en-Nüfeylî"],
    ["Ebü’l-Abbas Seffâh", "Ebü’l-Abbâs es-Seffâh"],
  ])("%s", (written, expected) => {
    const rows = buildInventory([{ part: "main", anchor: 0, order: 0, text: `Bu konuda ${written} şöyle der.` }], cat, index);
    expect(rows.map((r) => [r.text, r.person?.shortName])).toEqual([[written, expected]]);
  });
});

describe("ordinary words are not certain persons on their own", () => {
  const cat = new Catalog([
    p("Hâlî", "Hâlî Ahmed Efendi", "(ö. 1100/1688)"),
    p("Bahâr", "Bahâr Mehmed", "(ö. 1200/1785)"),
    p("Gazzâlî", "Ebû Hâmid Muhammed el-Gazzâlî", "(ö. 505/1111)"),
  ]);
  const index = new NameIndex(cat.people);
  const first = (text: string) => buildInventory([{ part: "main", anchor: 0, order: 0, text }], cat, index)[0];
  it("another Turkish letter than the catalog's: a candidate, not certain", () => {
    expect(first("Bu konuda Halı önemlidir.")?.person).toBeNull();
    expect(first("Bu konuda Hali önemlidir.")?.person?.shortName).toBe("Hâlî");
  });
  it("an ordinary word at the start of a sentence: a candidate; inside a sentence: certain", () => {
    expect(first("Bahar bu konuda önemlidir.")?.person).toBeNull();
    expect(first("Bu konuda Bahar şöyle der.")?.person?.shortName).toBe("Bahâr");
    expect(first("Bahar (ö. 1200/1785) bu konuda şöyle der.")?.person?.shortName).toBe("Bahâr");
    expect(first("Gazzâlî bu konuda şöyle der.")?.person?.shortName).toBe("Gazzâlî");
  });
});

describe("school eponyms and works after a name", () => {
  const cat = new Catalog([
    p("Şâfiî", "Ebû Abdillâh Muhammed b. İdrîs b. Abbâs eş-Şâfiî", "(ö. 204/820)"),
    p("Eş‘arî, Ebü’l-Hasan", "Ebü’l-Hasen Alî b. İsmâîl b. Ebî Bişr İshâk b. Sâlim el-Eş‘arî el-Basrî", "(ö. 324/935-36)"),
    p("Eş‘arî, Ebû Bürde", "Ebû Bürde Âmir b. Ebî Mûsâ Abdillâh b. Kays el-Eş‘arî", "(ö. 104/722)"),
    p("Gazzâlî", "Ebû Hâmid Muhammed el-Gazzâlî", "(ö. 505/1111)"),
  ]);
  const index = new NameIndex(cat.people);
  const who = (text: string) => buildInventory([{ part: "main", anchor: 0, order: 0, text }], cat, index).map((r) => [r.text, r.person?.shortName ?? null]);
  it("\"şöyle\" is not a lineage word after a school name", () => {
    expect(who("Bu konuda Şâfiî şöyle der.")).toEqual([["Şâfiî", "Şâfiî"]]);
    expect(who("Şâfiî fakihler böyle düşünür.")).toEqual([]);
  });
  it("Eş‘arî alone is Ebü’l-Hasan; as an adjective it is no one", () => {
    expect(who("Eş‘arî (ö. 324/935) bunu reddeder.")).toEqual([["Eş‘arî", "Eş‘arî, Ebü’l-Hasan"]]);
    expect(who("Eş‘arî kelâmcılar buna itiraz eder.")).toEqual([]);
  });
  it("a work after the name is not part of it", () => {
    expect(who("İmam Şâfiî er-Risâle'yi yazdı.")).toEqual([["Şâfiî", "Şâfiî"]]);
    expect(who("Bu görüşü Gazzâlî el-Mustasfâ'da savunur.")).toEqual([["Gazzâlî", "Gazzâlî"]]);
    expect(who("Gazzâlî el-Mustasfâ adlı eserinde bunu savunur.")).toEqual([["Gazzâlî", "Gazzâlî"]]);
  });
});

describe("epithets and short name + nisba", () => {
  const cat = new Catalog([
    p("Muhammed", "---", "(ö. 11/632)"),
    p("Ebû Hanîfe", "Ebû Hanîfe Nu‘mân b. Sâbit b. Zûtâ b. Mâh", "(ö. 150/767)"),
    p("Ahund Molla", "Muhammed Kâzım-ı Horasânî", "(1839-1911)"),
    p("Gazzâlî", "Hüccetü’l-İslâm Ebû Hâmid Muhammed b. Muhammed b. Ahmed el-Gazzâlî et-Tûsî", "(ö. 505/1111)"),
  ]);
  const index = new NameIndex(cat.people);
  const who = (text: string) => buildInventory([{ part: "main", anchor: 0, order: 0, text }], cat, index).map((r) => [r.text, r.person?.shortName ?? null]);
  it("Hz. Peygamber, Resûlullah, İmâm-ı Âzam: the person, in text and in the selected name", () => {
    expect(who("Bu konuda Hz. Peygamber şöyle buyurur.")).toEqual([["Peygamber", "Muhammed"]]);
    expect(who("Resûlullah'ın sünneti böyledir.")).toEqual([["Resûlullah'ın", "Muhammed"]]);
    expect(who("İmâm-ı Âzam bu görüştedir.")).toEqual([["İmâm-ı Âzam", "Ebû Hanîfe"]]);
    expect(lookup(cat, "Resûl-i Ekrem").automatic?.shortName).toBe("Muhammed");
    expect(who("Bu peygamber kıssası bilinir.")).toEqual([]);
  });
  it("a word of the short name with a nisba of the full name", () => {
    expect(who("Bu konuda Ahund el-Horasani şöyle der.")).toEqual([["Ahund el-Horasani", "Ahund Molla"]]);
    expect(lookup(cat, "Ahund el-Horasani").automatic?.shortName).toBe("Ahund Molla");
    expect(who("Gazzâlî et-Tûsî bunu savunur.")).toEqual([["Gazzâlî et-Tûsî", "Gazzâlî"]]);
  });
});

describe("Wikidata entries: initials and other spellings", () => {
  const w = (shortName: string, fullName: string, deathDate: string, id: string): Person => ({ shortName, fullName, originalName: "---", deathDate, description: "", source: `https://www.wikidata.org/wiki/${id}` });
  const cat = new Catalog([
    w("Émile Durkheim", "---", "(1858-1917)", "Q15948"),
    w("Max Weber", "---", "(1864-1920)", "Q9387"),
    w("Wilhelm Eduard Weber", "---", "(1804-1891)", "Q76780"),
    w("Aristo", "Aristoteles | Aristotales", "(m.ö. 384-322)", "Q868"),
  ]);
  const index = new NameIndex(cat.people);
  const who = (text: string) => buildInventory([{ part: "main", anchor: 0, order: 0, text }], cat, index).map((r) => [r.text, r.person?.shortName ?? null]);
  it("initials before a surname", () => {
    expect(lookup(cat, "É. Durkheim").automatic?.shortName).toBe("Émile Durkheim");
    expect(lookup(cat, "M. Weber").automatic?.shortName).toBe("Max Weber");
    expect(lookup(cat, "Weber").candidates.length).toBe(2);
  });
  it("each other spelling is a name of its own", () => {
    expect(lookup(cat, "Aristoteles").automatic?.shortName).toBe("Aristo");
    expect(who("Bu konuda Aristoteles şöyle der.")).toEqual([["Aristoteles", "Aristo"]]);
    expect(who("Durkheim de katılır.")).toEqual([["Durkheim", "Émile Durkheim"]]);
  });
});

describe("notability among Wikidata namesakes", () => {
  const w = (shortName: string, id: string, links: number): Person => ({ shortName, fullName: "---", originalName: "---", deathDate: "(1724-1804)", description: "", source: `https://www.wikidata.org/wiki/${id}#${links}` });
  it("one far better known bearer is chosen; close ones stay a choice; DİA is never outranked", () => {
    const cat = new Catalog([w("Immanuel Kant", "Q9312", 200), w("Edgar Kant", "Q123", 9)]);
    expect(lookup(cat, "Kant").automatic?.shortName).toBe("Immanuel Kant");
    const close = new Catalog([w("Max Weber", "Q9387", 120), w("Alfred Weber", "Q61", 45)]);
    expect(lookup(close, "Weber").automatic).toBeNull();
    const withDia = new Catalog([w("Immanuel Kant", "Q9312", 200), { ...w("Kant", "Q1", 0), source: "https://islamansiklopedisi.org.tr/kant" }]);
    expect(lookup(withDia, "Kant").automatic?.source).toBe("https://islamansiklopedisi.org.tr/kant");
  });
});
