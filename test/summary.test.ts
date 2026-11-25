import { describe, expect, it } from "vitest";
import { Catalog, summaryFlags } from "../src/core/index.js";
import { buildInventory } from "../src/addin/inventory.js";

const article = [
  "Gazzâlî'de Bilgi",
  "ÖZ",
  "Bu çalışmada Gazzâlî ve Fahreddin er-Râzî karşılaştırılmaktadır.",
  "Anahtar Kelimeler: Gazzâlî, Râzî, bilgi.",
  "Knowledge in al-Ghazālī",
  "Abstract",
  "This study compares Gazzâlî and Fahreddin er-Râzî.",
  "Keywords: Ghazali, Razi.",
  "GİRİŞ",
  "Gazzâlî bu konuyu ele alır.",
];

describe("Öz / Abstract", () => {
  it("marks the summaries and their keywords lines, not the title or the text", () => {
    expect(summaryFlags(article)).toEqual([false, true, true, true, false, true, true, true, false, false]);
    expect(summaryFlags(["Öz: Bu çalışmada Gazzâlî ele alınıyor.", "Gazzâlî der."])).toEqual([true, false]);
    // Without a keywords line the next heading closes it.
    expect(summaryFlags(["Özet", "Gazzâlî ve Râzî.", "Birinci Bölüm", "Gazzâlî der."])).toEqual([true, true, false, false]);
  });

  it("makes the text's own mention the first one", () => {
    const catalog = Catalog.fromCsv("Kısa Ad;Tam Ad;Orijinal Ad;Vefat Tarihi;Tanıtım;Kaynak Link\r\nGazzâlî;---;---;(ö. 505/1111);;https://islamansiklopedisi.org.tr/gazzali");
    const rows = buildInventory(article.map((text, anchor) => ({ part: "main" as const, anchor, order: 0, text })), catalog);
    const gazzali = rows.filter((r) => r.person?.shortName === "Gazzâlî");
    expect(gazzali.map((r) => r.segment)).toEqual([0, 9]); // the title and the text; the summaries are left out
    expect(gazzali.find((r) => r.segment === 0)?.first).toBe(false); // a title takes no date
    expect(gazzali.find((r) => r.segment === 9)?.first).toBe(true);
  });
});
