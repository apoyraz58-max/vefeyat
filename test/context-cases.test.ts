/**
 * Contextual detection on the published catalog: sect, order and dynasty
 * adjectives are not people; citations and the bibliography are left out,
 * the narrative around them is still read.
 */
import { existsSync, readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { buildInventory, type Segment } from "../src/addin/inventory.js";
import { Catalog, NameIndex, bibliographyFlags, isBibliographyEntry } from "../src/core/index.js";

const path = new URL("../data/catalog.csv", import.meta.url);
const have = existsSync(path);

describe.runIf(have)("contextual detection", () => {
  const catalog = Catalog.fromCsv(have ? readFileSync(path, "utf8") : "");
  const index = new NameIndex(catalog.people);
  const found = (text: string, part: "main" | "note" = "main") =>
    buildInventory([{ part, anchor: 0, order: 0, text }], catalog, index).map((r) => r.text);

  it.each([
    "Nakşibendî tarikatı Osmanlı'da yaygınlaşmıştır.",
    "Kādirî ve Rifâî tarikatları da etkiliydi.",
    "Mevlevî dervişleri semâ ederdi.",
    "Bektaşî geleneğinde bu anlayış görülür.",
    "Halvetî şeyhleri bu bölgede faaldi.",
    "Sühreverdî tarikatı Hindistan'a yayıldı.",
    "Çiştî silsilesi Hindistan'da kuruldu.",
    "Sünnî ve Şiî müellifler ayrışır.",
    "Ehl-i hadîs ile ehl-i re'y arasında fark vardır.",
    "Abbâsî halifesi bu dönemde hüküm sürdü.",
    "Fâtımî devleti Mısır'da kuruldu.",
    "Hanefî fıkhında bu mesele farklıdır.",
    "Şâfiî mezhebinde ise durum farklıdır.",
  ])("group adjective: %s", (text) => {
    expect(found(text)).toEqual([]);
  });

  it.each([
    ["Şâfiî bu görüşü er-Risâle'de savunur.", ["Şâfiî"]],
    ["İmam Şâfiî kıyası sistemleştirmiştir.", ["Şâfiî"]],
    ["Şâfiî, Hanefî fakihlerinden farklı olarak bu görüşü reddeder.", ["Şâfiî"]],
    ["Şeyh Şâzelî bu tarikatın kurucusudur.", ["Şâzelî"]],
    ["Gazzâlî, Râzî ve Cüveynî bu görüştedir.", ["Gazzâlî", "Râzî", "Cüveynî"]],
    ["Gazzâlî (1993: 45) aklın sınırlarını vurgular.", ["Gazzâlî"]],
    ["Cüveynî, Nizâmiye Medresesi'nde ders vermiştir; Gazzâlî de onun talebesidir.", ["Cüveynî", "Gazzâlî"]],
  ] as const)("person kept: %s", (text, names) => {
    expect(found(text)).toEqual(names);
  });

  it.each([
    "Bu görüş için bk. Gazzâlî, el-Mustasfâ, I, 45.",
    "Bu tartışma (Gazzâlî, 1993: 45) yaygındır.",
    "Bu tartışma (bk. Gazzâlî, 1993; Râzî, 1997) yaygındır.",
  ])("main-text citation left out: %s", (text) => {
    expect(found(text)).toEqual([]);
  });

  it("footnote: citation left out, narrative read", () => {
    expect(found("Bu konuda Gazzâlî farklı düşünür; bk. Gazzâlî, el-Munkız, s. 34.", "note")).toEqual(["Gazzâlî"]);
    expect(found("Gazzâlî'nin bu görüşü için bk. Cüveynî, el-Burhân, I, 23.", "note")).toEqual(["Gazzâlî'nin"]);
    expect(found("Gazzâlî, age, II, 45.", "note")).toEqual([]);
    expect(found("Gazzâlî bu eserinde İbn Sînâ'yı eleştirir.", "note")).toEqual(["Gazzâlî", "İbn Sînâ'yı"]);
  });

  it("bibliography section and entry-shaped paragraphs are left out", () => {
    const texts = [
      "Gazzâlî bu meseleyi tartışır.",
      "KAYNAKÇA",
      "İbn Haldûn. Mukaddime. Çev. Süleyman Uludağ. İstanbul: Dergâh Yayınları, 2004.",
      "Cüveynî, el-Burhân fî usûli'l-fıkh, Katar 1399.",
      "Ekler",
      "Gazzâlî'nin hayatı hakkında bir not.",
    ];
    const segments: Segment[] = texts.map((text, anchor) => ({ part: "main", anchor, order: 0, text }));
    expect(buildInventory(segments, catalog, index).map((r) => [r.segment, r.text])).toEqual([
      [0, "Gazzâlî"],
      [5, "Gazzâlî'nin"],
    ]);
  });
});

describe("bibliography shapes", () => {
  it("entries", () => {
    expect(isBibliographyEntry("Gazzâlî, Ebû Hâmid. el-Mustasfâ. Beyrut: Dâru'l-kütübi'l-ilmiyye, 1993.")).toBe(true);
    expect(isBibliographyEntry("Çağrıcı, Mustafa. \"Gazzâlî\". Türkiye Diyanet Vakfı İslâm Ansiklopedisi. İstanbul 1996, 13: 489-505.")).toBe(true);
    expect(isBibliographyEntry("Gazzâlî, 1111 yılında Tûs'ta vefat etmiştir.")).toBe(false);
    expect(isBibliographyEntry("Gazzâlî. Bu isim İslâm düşüncesinde önemli bir yer tutar.")).toBe(false);
  });
  it("section flags", () => {
    expect(bibliographyFlags(["Giriş", "5. Kaynakça", "x", "EKLER", "y"])).toEqual([false, true, true, false, false]);
  });
});
