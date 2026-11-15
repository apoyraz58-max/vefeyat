/**
 * Teslim ve dizin: consistency report, people report with chronology and a
 * minimal .docx writer. Word-independent; ported from VBA 0.3.41
 * `VfcScanDates`/`VfcLines` (Tutarlılık raporu) and `VfdNotes`/`VfdChronoText`
 * (Şahıs raporu).
 */
import { strToU8, zipSync } from "fflate";
import { type DateKind, type Person, chronoYear, dateKind, isValidDate, nameQuery, normalize, personKey, textName } from "../core/index.js";
import type { InventoryRow, Segment } from "./inventory.js";

// ---- Tutarlılık raporu ----

export interface ConsistencyEntry {
  readonly label: string;
  readonly count: number;
  /** Up to three samples with their location, for minority forms. */
  readonly examples: readonly string[];
}

export interface ConsistencyReport {
  readonly total: number;
  readonly markers: readonly ConsistencyEntry[];
  readonly kinds: readonly ConsistencyEntry[];
  readonly mixedMarkers: boolean;
  readonly mixedKinds: boolean;
}

const MARKERS: ReadonlyArray<readonly [string, string]> = [
  ["(ö.", "ö."],
  ["(öl.", "öl."],
  ["(v.", "v."],
];
const KIND_LABELS: Record<DateKind, string> = {
  hijriGregorian: "hicrî/milâdî (505/1111)",
  single: "tek yıl (1111)",
  range: "aralık (1100-1111)",
  other: "diğer",
};

export function consistencyReport(segments: readonly Segment[]): ConsistencyReport {
  const markerCount = new Map<string, string[]>();
  const kindCount = new Map<DateKind, string[]>();
  for (const segment of segments) {
    const place = segment.part === "note" ? "Dipnot" : "Metin";
    for (const [marker, label] of MARKERS) {
      for (let at = segment.text.indexOf(marker); at >= 0; at = segment.text.indexOf(marker, at + 1)) {
        const closing = segment.text.indexOf(")", at);
        if (closing < 0 || closing - at > 59) continue;
        const kind = dateKind(segment.text.slice(at + marker.length, closing));
        if (!kind) continue;
        const sample = `${segment.text.slice(at, closing + 1)}  [${place}]`;
        markerCount.set(label, [...(markerCount.get(label) ?? []), sample]);
        kindCount.set(kind, [...(kindCount.get(kind) ?? []), sample]);
      }
    }
  }
  const entries = <K extends string>(map: Map<K, string[]>, order: readonly K[], label: (k: K) => string): ConsistencyEntry[] => {
    const used = order.filter((k) => map.has(k));
    const max = Math.max(0, ...used.map((k) => (map.get(k) as string[]).length));
    return used.map((k) => {
      const samples = map.get(k) as string[];
      return { label: label(k), count: samples.length, examples: used.length > 1 && samples.length < max ? samples.slice(0, 3) : [] };
    });
  };
  const markers = entries(markerCount, MARKERS.map(([, l]) => l), (k) => k);
  const kinds = entries(kindCount, ["hijriGregorian", "single", "range", "other"] as const, (k) => KIND_LABELS[k]);
  return {
    total: markers.reduce((sum, m) => sum + m.count, 0),
    markers,
    kinds,
    mixedMarkers: markers.length > 1,
    mixedKinds: kinds.length > 1,
  };
}

// ---- Şahıs raporu ----

export interface PersonReportRow {
  readonly person: Person;
  readonly firstAlias: string;
  readonly location: "Metin" | "Dipnot";
  readonly mentions: number;
  readonly notes: string;
}

function tokenCount(value: string): number {
  const key = normalize(value);
  return key.length === 0 ? 0 : key.split(" ").length;
}

/** `VfdNotes`: short name at first mention, missing date, date only later. */
export function personNotes(person: Person, mentions: readonly InventoryRow[]): string {
  const notes: string[] = [];
  const first = mentions[0];
  if (!first) return "";
  const fullName = textName(person.fullName);
  if (fullName.length > 0 && tokenCount(nameQuery(first.text)) < tokenCount(fullName)) notes.push(`İlk geçişte kısa ad; tam ad önerisi: ${fullName}.`);
  if (!first.date && isValidDate(person.deathDate)) {
    const later = mentions.findIndex((row, i) => i > 0 && row.date);
    notes.push(later > 0 ? `Tarih ilk geçişte değil, ${later + 1}. geçişte.` : "İlk geçişte tarih yok.");
  }
  return notes.join(" ");
}

/** Certain people in first-mention order (confirmed homonyms included). */
export function personReport(rows: readonly InventoryRow[]): PersonReportRow[] {
  const groups = new Map<string, { person: Person; mentions: InventoryRow[] }>();
  for (const row of rows) {
    if (!row.person) continue;
    const key = personKey(row.person);
    const group = groups.get(key) ?? { person: row.person, mentions: [] };
    group.mentions.push(row);
    groups.set(key, group);
  }
  return [...groups.values()].map(({ person, mentions }) => ({
    person,
    firstAlias: (mentions[0] as InventoryRow).text,
    location: (mentions[0] as InventoryRow).part === "note" ? "Dipnot" : "Metin",
    mentions: mentions.length,
    notes: personNotes(person, mentions),
  }));
}

/** Chronology rows ordered by Gregorian death year; unknown years last. */
export function chronology(report: readonly PersonReportRow[]): Array<{ year: string; name: string; date: string }> {
  const rows = report.map((r) => ({ r, y: chronoYear(r.person.deathDate) }));
  rows.sort((a, b) => {
    if (a.y && !b.y) return -1;
    if (b.y && !a.y) return 1;
    return (a.y?.year ?? 0) - (b.y?.year ?? 0);
  });
  return rows.map(({ r, y }) => ({
    year: y ? `${y.year < 0 ? `m.ö. ${-y.year}` : y.year}${y.certain ? "" : " (?)"}` : "?",
    name: textName(r.person.shortName),
    date: r.person.deathDate,
  }));
}

// ---- Minimal .docx ----

function xml(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function paragraph(text: string, options: { bold?: boolean; size?: number } = {}): string {
  const props = `${options.bold ? "<w:b/>" : ""}${options.size ? `<w:sz w:val="${options.size}"/>` : ""}`;
  return `<w:p><w:r>${props ? `<w:rPr>${props}</w:rPr>` : ""}<w:t xml:space="preserve">${xml(text)}</w:t></w:r></w:p>`;
}

function table(header: readonly string[], rows: readonly (readonly string[])[]): string {
  const cell = (text: string, bold: boolean) =>
    `<w:tc><w:p><w:r><w:rPr>${bold ? "<w:b/>" : ""}<w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">${xml(text)}</w:t></w:r></w:p></w:tc>`;
  const border = '<w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/>';
  return (
    `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>${border}</w:tblBorders></w:tblPr>` +
    `<w:tr><w:trPr><w:tblHeader/></w:trPr>${header.map((h) => cell(h, true)).join("")}</w:tr>` +
    rows.map((row) => `<w:tr>${row.map((c) => cell(c, false)).join("")}</w:tr>`).join("") +
    "</w:tbl>"
  );
}

export interface DocxBlock {
  readonly heading?: string;
  readonly text?: string;
  readonly table?: { header: readonly string[]; rows: readonly (readonly string[])[] };
}

/** A landscape .docx with headings, paragraphs and bordered tables. */
export function buildDocx(title: string, blocks: readonly DocxBlock[]): Uint8Array {
  const body =
    paragraph(title, { bold: true, size: 28 }) +
    blocks
      .map((b) => (b.heading ? paragraph(b.heading, { bold: true, size: 24 }) : "") + (b.text ? paragraph(b.text) : "") + (b.table ? table(b.table.header, b.table.rows) + paragraph("") : ""))
      .join("");
  const document =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
    body +
    '<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr>' +
    "</w:body></w:document>";
  return zipSync({
    "[Content_Types].xml": strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    ),
    "_rels/.rels": strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    ),
    "word/document.xml": strToU8(document),
  });
}

/** Base64 for Word's createDocument. */
export function toBase64(bytes: Uint8Array): string {
  let binary = "";
  for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(binary);
}

/** The people report as a .docx: people table and chronology. */
export function personReportDocx(documentName: string, rows: readonly InventoryRow[]): Uint8Array {
  const report = personReport(rows);
  const flagged = report.filter((r) => r.notes).length;
  const uncertain = rows.filter((r) => !r.person && r.candidates.length > 0).length;
  return buildDocx(`Vefeyât şahıs raporu: ${documentName}`, [
    {
      text:
        `${report.length} şahıs, ${rows.filter((r) => r.person).length} geçiş. ${flagged} şahıs için not var. ` +
        `Yalnız kimliği kesin veya onaylı geçişler listelenir; sıra ilk geçiş sırasıdır.` +
        (uncertain ? ` ${uncertain} geçişin kimliği belirsiz olduğu için dahil edilmedi.` : ""),
    },
    {
      table: {
        header: ["Şahıs (DİA)", "İlk geçişteki yazım", "Yer", "Geçiş", "Ölüm tarihi (DİA)", "Notlar", "DİA bağlantısı"],
        rows: report.map((r) => [textName(r.person.shortName), r.firstAlias, r.location, String(r.mentions), r.person.deathDate, r.notes, r.person.source]),
      },
    },
    {
      heading: "Kronoloji (milâdî ölüm yılına göre; (?) işaretli yıllar tek sayıdan okunmuştur, hicrî olabilir)",
      table: { header: ["Milâdî yıl", "Şahıs (DİA)", "Ölüm tarihi (DİA)"], rows: chronology(report).map((c) => [c.year, c.name, c.date]) },
    },
  ]);
}
