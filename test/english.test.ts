import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { Catalog, NameIndex, assimilated, englishName, findMatches } from "../src/core/index.js";
import { lookup } from "../src/addin/logic.js";

describe("English transliteration of DİA's names", () => {
  it("takes the vowels from the Turkish and the consonants from the Arabic", () => {
    expect(englishName("Tâceddin Sübkî", "تاج الدين السبكي")).toBe("Tāj al-Dīn al-Subkī");
    expect(englishName("Fahreddin er-Râzî", "فخر الدين الرازي")).toBe("Fakhr al-Dīn al-Rāzī");
    expect(englishName("İbn Haldûn", "ابن خلدون")).toBe("Ibn Khaldūn");
    expect(englishName("Ebû Hanîfe", "أبو حنيفة")).toBe("Abū Ḥanīfa");
    expect(englishName("Süfyân es-Sevrî", "سفيان الثوري")).toBe("Sufyān al-Thawrī");
    expect(englishName("Mâlik b. Enes", "مالك بن أنس")).toBe("Mālik b. Anas");
    expect(englishName("Takıyyüddin Sübkî", "تقيّ الدين السبكي")).toBe("Taqī al-Dīn al-Subkī");
    expect(englishName("Şâfiî", "الشافعي")).toBe("al-Shāfiʿī");
    expect(englishName("Âişe", "عائشة")).toBe("ʿĀʾisha");
    expect(englishName("Gazzâlî", "محمد")).toBeNull(); // not the same name
  });

  it("writes the assimilated article too", () => {
    expect(assimilated("Tāj al-Dīn al-Subkī")).toBe("Tāj ad-Dīn as-Subkī");
    expect(assimilated("al-Shāfiʿī")).toBe("ash-Shāfiʿī");
    expect(assimilated("al-Ghazzālī")).toBe("al-Ghazzālī");
  });

  it("finds DİA's people written in English, in a selection and in running text", async () => {
    const catalog = Catalog.fromCsv(readFileSync(new URL("../data/catalog.csv", import.meta.url), "utf8"));
    expect(lookup(catalog, "Tāj al-Dīn al-Subkī").automatic?.shortName).toBe("Sübkî, Tâceddin");
    expect(lookup(catalog, "Taj al-Din al-Subki").automatic?.shortName).toBe("Sübkî, Tâceddin");
    expect(lookup(catalog, "Fakhr al-Din al-Razi").automatic?.shortName).toBe("Fahreddin er-Râzî");
    const index = await NameIndex.build(catalog.people);
    const text = "According to Tāj al-Dīn al-Subkī, al-Ghazālī and Fakhr al-Dīn al-Rāzī disagreed with Ibn Taymiyya. Ibn Khaldūn wrote later.";
    expect(findMatches(text, index, true).map((m) => m.text)).toEqual(["Tāj al-Dīn al-Subkī", "al-Ghazālī", "Fakhr al-Dīn al-Rāzī", "Ibn Taymiyya", "Ibn Khaldūn"]);
  }, 60000);
});
