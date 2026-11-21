import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { Catalog, arabicSkeleton, badCase, nameOptions, westernSpelling, buildSpellingStandard, longVowels, pairWords, wantsLongVowels, onlyAdds, parseCatalogCsv, spellingKey, standardName, standardPerson, wantsStandard } from "../src/core/index.js";

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
    expect(standardName("İbn Haldun", small)).toBe("İbn Haldun");
    expect(standardName("Ad-Dardir Malik", small)).toBe("Dardir Mâlik"); // no short name begins with the article // "İbn" at the start is a name, not "b."
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

  it("reads long vowels from the Arabic spelling of a word DİA does not have", () => {
    expect(longVowels("Turabi", "الترابي")).toBe("Turâbî");
    expect(longVowels("Kardavi", "القرضاوي")).toBe("Kardâvî");
    expect(longVowels("Suveydan", "السويدان")).toBe("Suveydân");
    expect(longVowels("Turabi", "محمد")).toBeNull(); // not the same word
    const real = buildSpellingStandard(parseCatalogCsv(readFileSync(new URL("../data/catalog.csv", import.meta.url), "utf8")));
    expect(standardName("Hasan el-Turabi", real, 0.5, "حسن الترابي", true)).toBe("Hasan et-Turâbî");
    expect(standardName("Hasan el-Turabi", real, 0.5, "حسن الترابي")).toBe("Hasan et-Turabi"); // only where asked
    expect(wantsLongVowels(["Q1049"])).toBe(true);
    expect(wantsLongVowels(["Q227"])).toBe(false); // Azerbaijan: Turkic names
  });

  it("writes rare Arabic words of DİA's own names as DİA does (at least 90%)", () => {
    const people = parseCatalogCsv(readFileSync(new URL("../data/catalog.csv", import.meta.url), "utf8"));
    const freq = new Map<string, number>();
    for (const p of people) for (const w of `${p.shortName} ${p.fullName}`.split(/\s+/)) freq.set(spellingKey(w.replace(/^[a-zşü]{2}-/, "")), (freq.get(spellingKey(w.replace(/^[a-zşü]{2}-/, ""))) ?? 0) + 1);
    let decided = 0;
    let right = 0;
    for (const p of people) {
      if (p.shortName.includes(",")) continue;
      for (const [latin, arabic] of pairWords(p.shortName, p.originalName) ?? []) {
        const word = latin.replace(/^[a-zşü]{2}-/, "");
        if (!/^\p{Lu}/u.test(word) || (freq.get(spellingKey(word)) ?? 0) > 2) continue;
        const dia = word.normalize("NFC").replace(/ā/g, "â").replace(/ī/g, "î").replace(/ū/g, "û").replace(/[‘’ʿʾ]/g, "");
        const plain = dia.normalize("NFD").replace(/\u0302/g, "").normalize("NFC");
        const got = longVowels(plain, arabic);
        if (got === null) continue;
        decided++;
        if (got === dia) right++;
      }
    }
    expect(decided).toBeGreaterThan(500);
    expect(right / decided).toBeGreaterThan(0.9);
  });

  it("joins Abdul as DİA does and lets the Arabic overrule a look-alike", () => {
    const real = buildSpellingStandard(parseCatalogCsv(readFileSync(new URL("../data/catalog.csv", import.meta.url), "utf8")));
    expect(standardName("Abdul Aziz al-Hakim", real, 0.5, "عبد العزيز الحكيم", true)).toBe("Abdülazîz el-Hakîm");
    expect(standardName("Abd al-Rahman Salih", real)).toMatch(/^Abdurrahm[âa]n /);
    expect(standardName("Hassan Nasrallah", real, 0.5, "حسن نصر الله", true)).not.toContain("Hassân");
    expect(standardName("Selim Hassan", real, 0.5, "سليم حسن", true)).toBe("Selîm Hassan");
  });

  it("leaves Abdul alone when it cannot be joined, and a yâ between vowels short", () => {
    const real = buildSpellingStandard(parseCatalogCsv(readFileSync(new URL("../data/catalog.csv", import.meta.url), "utf8")));
    expect(standardName("Abdul Qadir al-Badri", real, 0.5, "عبد القادر البدري", true)).toMatch(/^Abdul Qadir /);
    expect(standardName("Abdul Halim", real, 0.5, "عبد الحليم", true)).toBe("Abdülhalîm");
    expect(longVowels("Eslaiah", "اصليح") ?? "").not.toContain("î");
    expect(longVowels("Said", "سعيد")).toBe("Saîd");
  });

  it("mends a slip in a Western name's Turkish label, and keeps Turkish renderings", () => {
    const labels = ["Friedrich Schleiermacher", "Friedrich Schleiermacher", "Friedrich Schleiermacher", "Friedrich Daniel Ernst Schleiermacher"];
    expect(westernSpelling("Friedrich Scleirmacher", labels, ["Q27306"])).toBe("Friedrich Schleiermacher");
    expect(westernSpelling("Friedrich Scleirmacher", labels, ["Q159"])).toBeNull(); // not a Latin-alphabet country
    expect(westernSpelling("Sogomon Tehliryan", ["Soghomon Tehlirian", "Soghomon Tehlirian"], ["Q12560", "Q142"])).toBeNull(); // also Ottoman
    expect(westernSpelling("Mihail Bakunin", ["Mikhail Bakunin", "Mikhail Bakunin"], ["Q34266", "Q39"])).toBeNull(); // also Russian
    expect(westernSpelling("Emanuel Swedenborg", ["Emanuel Swedenborg", "Emmanuel Swedenborg", "Emmanuel Swedenborg"], ["Q34"])).toBeNull(); // Swedish writes it so
    expect(westernSpelling("Jerome Kağan", ["Jerome Kagan", "Jerome Kagan"], ["Q30"])).toBeNull(); // a Turkish letter: a rendering
    expect(westernSpelling("Platon", ["Plato", "Platon", "Platone"], ["Q30"])).toBeNull();
    expect(westernSpelling("Alejo Carpentier", ["Alejo Carpentier", "Alejo Carpanter"], ["Q241"])).toBeNull();
  });

  it("never offers a spelling with wrong capitals", () => {
    expect(badCase("Friedrich daniel ernst schleiermacher")).toBe(true);
    expect(badCase("Levis carroll")).toBe(true);
    expect(badCase("Ludwig van Beethoven")).toBe(false);
    expect(badCase("Alî b. Ebû Tâlib")).toBe(false);
    expect(badCase("José Ortega y Gasset")).toBe(false);
    expect(badCase("Hasan et-Turâbî")).toBe(false);
    const person = { shortName: "Friedrich Schleiermacher", fullName: "Friedrich daniel ernst schleiermacher", originalName: "", deathDate: "(1768-1834)", description: "", source: "x" };
    expect(nameOptions(person, "Friedrich daniel ernst schleiermacher").map((o) => o.value)).toEqual(["Friedrich Schleiermacher"]);
  });

  it("leaves Armenian, Slavic, Greek and Romanian surnames' names as their people write them", () => {
    const real = buildSpellingStandard(parseCatalogCsv(readFileSync(new URL("../data/catalog.csv", import.meta.url), "utf8")));
    expect(standardName("Serkis Minasyan", real)).toBe("Serkis Minasyan");
    expect(standardName("Kuzman Şapkarev", real)).toBe("Kuzman Şapkarev");
    expect(standardName("Tahir Salahov", real)).toBe("Tahir Salahov");
    expect(standardName("Yusuf Ziya Paşa", real)).toBe("Yûsuf Ziya Paşa"); // others as before
    expect(standardName("Tahir Efendi Jakova", real)).toBe("Tâhir Efendi Jakova"); // an Albanian place, not a Slavic surname
  });
});
