/**
 * Kaynakça: names in the bibliography are authors of works, not person
 * mentions. A heading ("Kaynakça", "Bibliyografya", "References" …) opens the
 * section until a following section heading ("Ekler", "Dizin", "Özgeçmiş" …);
 * outside it, a paragraph shaped like a bibliography entry is skipped too.
 */
import { isNarrative, words } from "./context.js";
import { normalize } from "./normalize.js";

const OPEN = new Set([
  "kaynakca", "kaynakce", "kaynaklar", "kaynakca ve kisaltmalar", "bibliyografya", "bibliografya", "bibliyografi",
  "bibliography", "references", "works cited", "referanslar", "yararlanilan kaynaklar", "kullanilan kaynaklar",
  "basvurulan kaynaklar", "atif yapilan kaynaklar", "secilmis bibliyografya", "kaynaklar ve arastirmalar",
]);
const CLOSE = new Set([
  "ekler", "ek", "dizin", "indeks", "index", "sahislar dizini", "genel dizin", "ozgecmis", "appendix", "appendices",
  "cv", "curriculum vitae", "sozluk", "terimler sozlugu", "ozet", "abstract", "sonuc", "summary",
]);

/** The heading's key: numbering ("5.", "IV.", "B)") and punctuation dropped. */
function headingKey(text: string): string {
  const trimmed = text.trim();
  if (trimmed.length === 0 || trimmed.length > 60) return "";
  return normalize(trimmed)
    .replace(/^((\d+|[ivxlc]+|[a-h])\s+)+/, "")
    .trim();
}

export function isBibliographyHeading(text: string): boolean {
  return OPEN.has(headingKey(text));
}

export function isSectionAfterBibliography(text: string): boolean {
  const key = headingKey(text);
  return CLOSE.has(key) || /^ek \d+$/.test(key) || /^ek [a-h]$/.test(key);
}

const YEAR = /(^|[^0-9])(1[0-9]{3}|20[0-9]{2})([^0-9]|$)|\bt\.\s?y\.|\(ty\)/;
const PUBLISHER =
  /: ?[A-ZÇĞİÖŞÜÂÎÛ]|\b(yay(ın(ları|evi)?|\.)|press|publications?|verlag|matbaa(sı)?|neşriyat|dâru|daru|mektebe(tü)?|müessese(tü)?|kütüphane(si)?|vakfı|üniversitesi|dergisi|journal|ansiklopedisi|thk\.|tahk\.|nşr\.|çev\.|trc\.|haz\.|ed\.|eds\.)/i;

/**
 * A paragraph that reads like one bibliography entry: an author block,
 * a year, a publication sign and no narrative verb.
 *   Gazzâlî, Ebû Hâmid. el-Mustasfâ. Beyrut: Dâru'l-kütübi'l-ilmiyye, 1993.
 *   İbn Haldûn. Mukaddime. Çev. Süleyman Uludağ. İstanbul: Dergâh Yayınları, 2004.
 */
export function isBibliographyEntry(text: string): boolean {
  const value = text.trim().replace(/^(\d+[.)]|[-–•·*])\s*/, "");
  if (value.length < 20 || value.length > 700) return false;
  if (!/^[A-ZÇĞİÖŞÜÂÎÛĀĪŪḤṢṬẒʿʾ'‘’"“]/.test(value)) return false;
  if (!YEAR.test(value) || !PUBLISHER.test(value)) return false;
  // Author block: "Soyad, Ad." or "Ad Soyad." before the title, short.
  const author = /^[^.,:;()"“]{2,70}(, [^.:;()"“]{1,70})?\. /.exec(value) ?? /^[^.,:;()"“]{2,70}, /.exec(value);
  if (!author) return false;
  return !isNarrative(words(value));
}

/**
 * For each main-text paragraph in reading order: whether it belongs to the
 * bibliography (section or entry shape). Headings themselves are excluded too.
 */
export function bibliographyFlags(paragraphs: readonly string[]): boolean[] {
  let inside = false;
  return paragraphs.map((text) => {
    if (isBibliographyHeading(text)) {
      inside = true;
      return true;
    }
    if (inside && isSectionAfterBibliography(text)) inside = false;
    return inside || isBibliographyEntry(text);
  });
}

const ORDINALS = "birinci|ikinci|ucuncu|dorduncu|besinci|altinci|yedinci|sekizinci|dokuzuncu|onuncu";

/**
 * A chapter heading: Word's Heading 1, or a line like "BİRİNCİ BÖLÜM",
 * "2. Bölüm: …", "GİRİŞ", "SONUÇ". Chapters restart first mentions under
 * "Her bölümde ilk kullanım".
 */
export function isChapterHeading(text: string, style = ""): boolean {
  if (/^(heading ?1|başlık ?1)$/i.test(style.trim())) return true;
  const value = text.trim();
  if (value.length === 0 || value.length > 120) return false;
  const key = normalize(value);
  if (new RegExp(`^((\\d+|[ivxlc]+)\\s+)?(${ORDINALS}|\\d+)\\s+bolum(\\s|$)`).test(key)) return true;
  return /^((\d+|[ivxlc]+)\s+)?(giris|sonuc|sonuc ve degerlendirme)$/.test(key);
}
