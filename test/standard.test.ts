import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { Catalog, arabicSkeleton, buildSpellingStandard, onlyAdds, parseCatalogCsv, spellingKey, standardName, standardPerson, wantsStandard } from "../src/core/index.js";

const dia = (short: string, full = "---") => ({ shortName: short, fullName: full, originalName: "", deathDate: "(ö. 1900)", description: "", source: `https://islamansiklopedisi.org.tr/${spellingKey(short)}` });
const small = buildSpellingStandard([
  dia("Yûsuf b. Ali"),
  dia("Yûsuf el-Hemedânî"),
  dia("Yûsuf Ağa"),
  dia("Yusuf Has Hâcib"), // DİA writes Yûsuf far more often
  dia("Ali Suâvî"),
  dia("Alî b. Ebû Tâlib"), // Ali and Alî: written both ways, left alone
  dia("Mâlik b. Enes"),
  dia("Mâlik b. Nebî"),
  dia("Kürt İsmâil Paşa"),
  dia("Smıth, Adam"), // a slip in the list must not reach "Smith"
  dia("Mîhâîl Nuayme"),
  { ...dia("Râzî"), source: "https://www.wikidata.org/wiki/Q1" }, // not DİA: not counted
]);

describe("standart yazım (DİA esas)", () => {
  it("keys words without marks and Turkish letters", () => {
    expect(spellingKey("Türâbî")).toBe(spellingKey("Turabi"));
    expect(spellingKey("Ka‘nebî")).toBe("kanebi");
  });

  it("takes DİA's clear majority and leaves a split word alone", () => {
    expect(standardName("Yusuf Ali", small)).toBe("Yûsuf Ali");
    expect(standardName("Alî Yusuf", small)).toBe("Alî Yûsuf");
  });

  it("only adds marks: never drops or changes a letter", () => {
    expect(onlyAdds("Malik", "Mâlik")).toBe(true);
    expect(onlyAdds("Abdulaziz", "Abdülazîz")).toBe(true);
    expect(onlyAdds("Kürt", "Kurt")).toBe(false);
    expect(onlyAdds("Smith", "Smıth")).toBe(false);
    expect(onlyAdds("Mean", "Ma‘n")).toBe(false);
    expect(standardName("Smith Yusuf", small)).toBe("Smith Yûsuf");
  });

  it("assimilates the article and writes b. as DİA does", () => {
    expect(standardName("Malik bin Nebi", small)).toBe("Mâlik b. Nebî");
    expect(standardName("Hasan el-Turabi", small)).toBe("Hasan et-Turabi");
    expect(standardName("Ahmad al-Razi", small)).toBe("Ahmad er-Razi");
    expect(standardName("Halid al-Şafii", small)).toBe("Halid eş-Şafii");
    expect(standardName("Ali al-Cezeri", small)).toBe("Ali el-Cezeri");
    expect(standardName("İbn Haldun", small)).toBe("İbn Haldun"); // "İbn" at the start is a name, not "b."
  });

  it("leaves a name most of whose words DİA does not have", () => {
    expect(standardName("George Kennan Malik", small)).toBe("George Kennan Malik");
    expect(standardName("Mihail Bakunin", small)).toBe("Mihail Bakunin");
  });

  it("keeps the spelling it had among the names", () => {
    const person = standardPerson({ ...dia("Malik bin Nebi"), fullName: "Malik Bennabi", source: "https://www.wikidata.org/wiki/Q2" }, small);
    expect(person.shortName).toBe("Mâlik b. Nebî");
    expect(person.fullName).toBe("Malik bin Nebi | Malik Bennabi");
    expect(new Catalog([person]).find("Malik bin Nebi", false)).toHaveLength(1);
  });

  it("is for the Muslim world and Ottoman subjects before 1928", () => {
    expect(wantsStandard(["Q1049"], 2016)).toBe(true); // Sudan
    expect(wantsStandard(["Q12560"], 1890)).toBe(true);
    expect(wantsStandard(["Q12560", "Q43"], 1950)).toBe(false); // a Turk of the Republic
    expect(wantsStandard(["Q183"], 1900)).toBe(false); // Germany
    expect(wantsStandard(["Q668"], 1900)).toBe(false); // India: many faiths
  });

  it("writes well-known names as the DİA catalog does", () => {
    const real = buildSpellingStandard(parseCatalogCsv(readFileSync(new URL("../data/catalog.csv", import.meta.url), "utf8")));
    expect(standardName("Malik bin Nebi", real)).toBe("Mâlik b. Nebî");
    expect(standardName("Ali Şeriati", real)).toBe("Ali Şerîatî");
    expect(standardName("Abdulaziz bin Baz", real)).toBe("Abdülazîz b. Bâz");
    expect(standardName("Hasan el-Turabi", real)).toBe("Hasan et-Turabi");
  });

  it("tells apart names written alike in Latin letters by their Arabic", () => {
    const withArabic = (short: string, arabic: string) => ({ ...dia(short), originalName: arabic });
    const table = buildSpellingStandard([
      withArabic("Hâmid b. Abbas", "حامد بن عباس"),
      withArabic("Hâmid el-Âmidî", "حامد الآمدي"),
      withArabic("Hâmid Efendi", "حامد افندي"),
      withArabic("Hamîd ed-Dîn", "حميد الدين"),
    ]);
    expect(arabicSkeleton("الحَمِيد")).toBe("حميد");
    expect(standardName("Hamid Abbas", table, 0.5, "حميد عباس")).toBe("Hamîd Abbas");
    expect(standardName("Hamid Abbas", table, 0.5, "حامد عباس")).toBe("Hâmid Abbas");
    // Without the Arabic word, the two names are not guessed between.
    expect(standardName("Hamid Abbas", table)).toBe("Hamid Abbas");
  });
});
