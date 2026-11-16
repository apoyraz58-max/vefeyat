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
    "Usûlî açıdan bakıldığında bu delil zayıftır.",
    "Usûlî bir yaklaşım benimsenmiştir.",
    "Lügavî anlamı farklıdır.",
    "Nahvî bir tahlil yapılmıştır.",
    "Hakîm bir kişi böyle davranmaz.",
    "Irak'taki medreseler gelişmişti.",
    "Arap dili ve edebiyatı önemlidir.",
    "Bil ki bu yol zordur.",
    "Minhac'ı şerh eden âlimler çoktur.",
    "Ehl-i Sünnet bu konuda farklıdır.",
    "Cem-i esmâ bu bağlamda ele alınır.",
    "Muhaddis, Şâfiî fakihi.",
    "Özbek hanı bu bölgeyi yönetti.",
    "Suûdî ailesine mensup bir emîrdi.",
    "Fars edebiyatında sebk-i Hindî'nin kurucusu sayılır.",
    "İsmâilî filozof ve şair.",
    "Kadı, edip ve şair.",
    "Tunuslu bir âlim olarak tanınır.",
    "İngiliz şarkiyatçısı bu eseri neşretti.",
    "Bu konu, Usûlî açıdan ele alınmalıdır.",
    "Osmanlı âlimleri Mâtürîdî idi.",
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
    ["Gazzâlî kelâm ilmine yöneldi.", ["Gazzâlî"]],
    ["Gazzâlî fıkıh usulünde el-Mustasfâ'yı yazdı.", ["Gazzâlî"]],
    ["Râzî tefsir alanında Mefâtîhu'l-gayb'ı yazmıştır.", ["Râzî"]],
    ["Bu dönemde Gazzâlî medrese hocalığı yaptı.", ["Gazzâlî"]],
    ["Gazzâlî devlet adamlarıyla yakın ilişki kurdu.", ["Gazzâlî"]],
    ["Süyûtî kitap yazmayı çok severdi.", ["Süyûtî"]],
    ["Teftâzânî görüşlerini Şerhu'l-Makāsıd'da açıklar.", ["Teftâzânî"]],
    ["Cem Sultan'ın sürgün yılları uzun sürdü.", ["Cem Sultan'ın"]],
    ["İmam Mâlik'in talebesi.", ["Mâlik'in"]],
    ["Hayâtî Ahmed Efendi bu tekkede şeyhti.", ["Hayâtî Ahmed Efendi"]],
    ["Hocası Cüveynî (ö. 478/1085) Nizâmiye'de ders vermiştir.", ["Cüveynî"]],
    ["Talebesi Gazzâlî bu görüşü geliştirdi.", ["Gazzâlî"]],
    ["Nitekim Fahreddin er-Râzî de bunu kabul eder.", ["Fahreddin er-Râzî"]],
    ["F. er-Râzî bu görüşü savunur.", ["F. er-Râzî"]],
    ["Kitapta ‘Gazzâlî’ adı geçer.", ["Gazzâlî"]],
    ["“Râzî” burada farklı düşünür.", ["Râzî"]],
    ["Bu konuda Gazzâlî, Cüveynî ve Şâfiî aynı görüştedir.", ["Gazzâlî", "Cüveynî", "Şâfiî"]],
    ["Mâtürîdî, Kitâbü't-Tevhîd'de aklı öne çıkarır.", ["Mâtürîdî"]],
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

  it("a word written in lower case elsewhere is an adjective at a sentence start", () => {
    const segments: Segment[] = [
      { part: "main", anchor: 0, order: 0, text: "Bu meseleyi usûlî bakışla ele alalım." },
      { part: "main", anchor: 1, order: 0, text: "Usûlî gelenek burada belirleyicidir." },
    ];
    expect(buildInventory(segments, catalog, index)).toEqual([]);
  });

  it("headings: a one-word match the document writes in lower case is that word", () => {
    const segments: Segment[] = [
      { part: "main", anchor: 0, order: 0, text: "HAYATI VE ESERLERİ" },
      { part: "main", anchor: 1, order: 0, text: "Gazzâlî'nin Hayatı ve Eserleri" },
      { part: "main", anchor: 2, order: 0, text: "Gazzâlî'nin hayatı hakkında çok şey yazılmıştır." },
      { part: "main", anchor: 3, order: 0, text: "Cem sultan sürgüne gitti." },
    ];
    expect(buildInventory(segments, catalog, index).map((r) => r.text)).toEqual(["Gazzâlî'nin", "Gazzâlî'nin"]);
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

describe("chapter headings", async () => {
  const { isChapterHeading } = await import("../src/core/index.js");
  it.each([
    ["BİRİNCİ BÖLÜM", "", true],
    ["2. Bölüm: Gazzâlî'nin Hayatı", "", true],
    ["İkinci Bölüm", "", true],
    ["GİRİŞ", "", true],
    ["Sonuç", "", true],
    ["Gazzâlî'nin Hayatı", "Heading1", true],
    ["Gazzâlî'nin Hayatı", "Heading2", false],
    ["Bu bölümde Gazzâlî ele alınır.", "", false],
  ] as const)("%s (%s) → %s", (text, style, expected) => {
    expect(isChapterHeading(text, style)).toBe(expected);
  });
});
