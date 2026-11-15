import { unzipSync, strFromU8 } from "fflate";
import { describe, expect, it } from "vitest";
import { buildInventory, type Segment } from "../src/addin/inventory.js";
import { buildDocx, chronology, consistencyReport, personReport, personReportDocx } from "../src/addin/submit.js";
import { Catalog, type Person } from "../src/core/index.js";

const DIA = "https://islamansiklopedisi.org.tr/";
const p = (shortName: string, fullName: string, deathDate: string, slug: string): Person => ({
  shortName,
  fullName,
  originalName: "",
  deathDate,
  description: "",
  source: DIA + slug,
});
const catalog = new Catalog([
  p("Gazzâlî", "Ebû Hâmid Muhammed b. Muhammed el-Gazzâlî", "(ö. 505/1111)", "gazzali"),
  p("Cüveynî", "Ebü'l-Meâlî Abdülmelik el-Cüveynî", "(ö. 478/1085)", "cuveyni"),
  p("Aristo", "Aristoteles", "(m.ö. 384-322)", "aristo"),
]);
const main = (anchor: number, text: string): Segment => ({ part: "main", anchor, order: 0, text });

describe("consistencyReport", () => {
  it("counts markers and forms, gives minority examples", () => {
    const report = consistencyReport([
      main(0, "Gazzâlî (ö. 505/1111) ve Cüveynî (ö. 478/1085) der."),
      { part: "note", anchor: 0, order: 1, text: "Aristo (v. 322) de." },
    ]);
    expect(report.total).toBe(3);
    expect(report.markers).toEqual([
      { label: "ö.", count: 2, examples: [] },
      { label: "v.", count: 1, examples: ["(v. 322)  [Dipnot]"] },
    ]);
    expect(report.mixedMarkers).toBe(true);
    expect(report.kinds.map((k) => [k.label, k.count])).toEqual([
      ["hicrî/milâdî (505/1111)", 2],
      ["tek yıl (1111)", 1],
    ]);
  });
  it("consistent document", () => {
    const report = consistencyReport([main(0, "Gazzâlî (ö. 505/1111).")]);
    expect(report.mixedMarkers || report.mixedKinds).toBe(false);
  });
});

describe("person report and chronology", () => {
  const rows = buildInventory(
    [main(0, "Gazzâlî der ki. Cüveynî (ö. 478/1085) hocasıdır."), main(1, "Aristo eski bir filozoftur. Gazzâlî (ö. 505/1111) yine der.")],
    catalog,
  );
  it("people in first-mention order with notes", () => {
    const report = personReport(rows);
    expect(report.map((r) => [r.person.shortName, r.mentions])).toEqual([
      ["Gazzâlî", 2],
      ["Cüveynî", 1],
      ["Aristo", 1],
    ]);
    expect(report[0]!.notes).toBe("İlk geçişte kısa ad; tam ad önerisi: Ebû Hâmid Muhammed b. Muhammed el-Gazzâlî. Tarih ilk geçişte değil, 2. geçişte.");
  });
  it("chronology by Gregorian year, B.C. first", () => {
    expect(chronology(personReport(rows)).map((c) => c.year)).toEqual(["m.ö. 384", "1085", "1111"]);
  });
  it("report docx contains both tables", () => {
    const files = unzipSync(personReportDocx("Tez.docx", rows));
    const xml = strFromU8(files["word/document.xml"]!);
    expect(Object.keys(files).sort()).toEqual(["[Content_Types].xml", "_rels/.rels", "word/document.xml"]);
    expect(xml).toContain("Vefeyât şahıs raporu: Tez.docx");
    expect(xml).toContain("Kronoloji");
    expect((xml.match(/<w:tbl>/g) ?? []).length).toBe(2);
  });
  it("docx text is escaped", () => {
    const xml = strFromU8(unzipSync(buildDocx("A & <B>", []))["word/document.xml"]!);
    expect(xml).toContain("A &amp; &lt;B&gt;");
  });
});

describe("Şahıslar dizini", () => {
  it("field codes", async () => {
    const { xeCode, INDEX_CODE, isVefeyatIndexField } = await import("../src/addin/submit.js");
    const person = p('Ebû "Ali": İbn Sînâ', "", "", "x");
    expect(xeCode(person)).toBe(`"Ebû 'Ali'\\: İbn Sînâ" \\f "v"`);
    expect(isVefeyatIndexField(` XE ${xeCode(person)} `)).toBe(true);
    expect(isVefeyatIndexField(` INDEX ${INDEX_CODE}`)).toBe(true);
    expect(isVefeyatIndexField(` XE "Başka" `)).toBe(false);
    expect(isVefeyatIndexField(` INDEX \\f "k"`)).toBe(false);
    expect(isVefeyatIndexField(` TOC \\f "v"`)).toBe(false);
  });
});

describe("indexRows", () => {
  it("one mark per person and paragraph", async () => {
    const { indexRows } = await import("../src/addin/submit.js");
    const rows = buildInventory([main(0, "Gazzâlî der. Gazzâlî yine der. Cüveynî de."), main(1, "Gazzâlî sonra.")], catalog);
    expect(indexRows(rows).map((r) => [r.segment, r.text])).toEqual([
      [0, "Gazzâlî"],
      [0, "Cüveynî"],
      [1, "Gazzâlî"],
    ]);
  });
});

describe("static index from pages", () => {
  it("pages, letters, Roman front matter, article ignored", async () => {
    const { staticIndex, pageLabel } = await import("../src/addin/submit.js");
    const g = p("Gazzâlî", "", "", "g");
    const c = p("Cüveynî", "", "", "c");
    const z = p("ez-Zemahşerî", "", "", "z");
    const lines = staticIndex(
      [
        { person: g, page: 14 },
        { person: c, page: 3 },
        { person: g, page: 12 },
        { person: g, page: 14 },
        { person: z, page: 20 },
      ],
      11,
    );
    expect(lines.map((l) => l.text)).toEqual(["C", "Cüveynî, iii", "G", "Gazzâlî, 2, 4", "Z", "ez-Zemahşerî, 10"]);
    expect(pageLabel(9, 11)).toBe("ix");
  });
});
