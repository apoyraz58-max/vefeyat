// Builds docs/vefeyat-deneme.docx: a test thesis for trying every command in
// real Word (main text, footnotes, endnotes, a VBA date field, homonyms,
// sect names, citations, a bibliography), with the checklist at the top.
import { writeFileSync } from "node:fs";
import { strToU8, zipSync } from "fflate";

const x = (t) => t.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const run = (t, props = "") => `<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ""}<w:t xml:space="preserve">${x(t)}</w:t></w:r>`;
const para = (content, style = "") => `<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ""}${content}</w:p>`;
const p = (t, style) => para(run(t), style);
let fn = 0;
let en = 0;
const footnotes = [];
const endnotes = [];
const foot = (t) => {
  fn++;
  footnotes.push(`<w:footnote w:id="${fn + 1}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>${run(" " + t)}</w:p></w:footnote>`);
  return `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="${fn + 1}"/></w:r>`;
};
const end = (t) => {
  en++;
  endnotes.push(`<w:endnote w:id="${en + 1}"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteRef/></w:r>${run(" " + t)}</w:p></w:endnote>`);
  return `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="${en + 1}"/></w:r>`;
};
// A date as the VBA add-in wrote it: ADDIN field with the date as its result.
const vbaDate = (t) =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ADDIN Vefeyat.Date.D0123456789ABCDEF.P1 </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${run(t)}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;

// The checklist names no one: it is part of the document, and a name in it
// would be the first mention of that person (dates, audit, "daha önce geçti").
const check = [
  "KONTROL LİSTESİ (bu bölümde şahıs adı geçmez; denemeden sonra silebilirsiniz)",
  "1. Şahıs dökümü: GİRİŞ'teki 1.–4. paragraflardaki altı kişi, 6. paragrafın son cümlesindeki imam ve 8. paragraftaki adlar listelenmeli; dipnot ve sonnottakiler 'Dipnot'/'Sonnot' etiketiyle. 5. paragraf, 6. paragrafın ilk iki cümlesi, 7. paragraftaki atıflar, başlıklar ve KAYNAKÇA listelenmemeli.",
  "2. 1. paragraftaki kısa adlar 'kimlik belgedeki diğer geçişlerden'; 4. paragraftaki tarihli kısa ad 'kimlik yazılı tarihten' olmalı; 3. paragraftaki baş harfli ad ve konusuz kısa ad tefsir müellifine bağlanmalı (ikincisi 'kimlik konudan (tefsir)').",
  "3. Belge sağlığı: 1 VBA tarihi bulunmalı; 'VBA tarihlerini dönüştür' sonrası 2. paragraftaki tarih metinde aynen kalmalı ve listede sağlam görünmeli.",
  "4. 4. paragraftaki ilk tam adı seçip Ölüm tarihi ekle: 'daha önce geçti' uyarısı çıkmamalı. 4. paragraftaki tarihli kısa adı seçince doğru kişi işaretli gelmeli. Sonra Tarihin kaynağı; sonra Ayarlar → Tarih alanlarını göster/gizle.",
  "5. Şahıslar dizini → oluştur (süre ve [sayfa API] bilgisini not edin) → Dizin işaretlerini kaldır (dizin metin kalmalı) → Dizini tamamen kaldır (hiçbir iz kalmamalı).",
  "6. Tutarlılık raporu: ö. ve v. karışık uyarısı. Şahıs raporu: yeni belge açılmalı. Şahıs dökümü özetinde '1 tarih uyarısı' olmalı: 6. paragraftaki elle yazılmış tarih yanlış (hicrî yıl doğru, milâdî yıl 100 yıl kaymış).",
  "7. Toplu düzeltme önizlemesi, Denetime devam et (bir geçişte Denetimi bitir, sonra devam et), Tüm belgeyi denetle (her kişinin ilk geçişinde durmalı), Teslim öncesi temizle (kopyada!).",
  "8. Sağ tık: bir ada sağ tıklayın; kırmızı çizgiliyse önce Ayarlar → Şahıs adlarını yazım denetiminden çıkar.",
  "9. Klavye kısayolları: bir adı seçip Alt+Shift+V tarih eklemeli, Alt+Shift+G Şahıs dökümünü açmalı. Ayarlar → Klavye kısayolları altı kısayolu listelemeli; fareyle Ölüm tarihi ekle'nin üzerine gelince '(Alt+Shift+V)' yazmalı.",
  "10. 4. paragrafta 1037 tarihli adı seçip Ölüm tarihi ekle: 'Adın arkasında zaten bir tarih var … Katalogdaki tarih: (ö. …)' ve Güncelle çıkmalı (kısaltma farklı). Güncelle'den sonra tarih katalogdaki gibi ve adla tarih, tarihle sonraki kelime arasında tek boşluk olmalı. Aynı ada yeniden basınca 'zaten katalogdakiyle aynı' demeli.",
  "11. Birkaç tarih ekledikten sonra Kısaltma → v.: panel 'çevrilsin mi?' diye sormalı; Çevir'den sonra eklenen tarihler v. ile yazılmalı, 6. paragraftaki elle yazılmış tarih değişmemeli. Eklenen bir tarihi elle silin: yerinde 'Metin girmek için…' yazısı görünmemeli. Belge sağlığı boşluk uyarısı vermemeli.",
  "12. Bir kişiye tek tarih: 1. paragrafta ilk ada tarih ekleyin, sonra aynı paragraftaki kısa ada Ölüm tarihi ekle: tarih eklenmemeli; 'Bu şahsın tarihi belgede zaten var' ve Tarihe git çıkmalı. Tarihe git ilk adı seçmeli.",
  "13. Şahıs dökümü süzgeci: listenin üstünde seçenekler sayılarıyla görünmeli; 'Doğrulanacak' yalnız kimliği kesin olmayanları göstermeli. Arama kutusuna 'razi' yazınca yalnız 3. ve 4. paragraftaki adlar kalmalı.",
  "14. 8. paragraf: lakaplı ad (Fahrülislâm …) katalogdaki Ebü'l-Usr kaydına bağlanmalı. Tek z ile yazılmış adı seçip İsim yazımı'nda doğru yazımı seçin, Tümünü değiştir: kaç geçişin değişeceğini söylemeli; Onayla'dan sonra 8. paragraftaki iki geçiş ekleriyle birlikte düzelmeli.",
  "15. Birden çok adayı olan bir kısa adda (3. veya 4. paragraf) kişiyi elle seçince 'Bu yazımın bütün belirsiz geçişlerine uygula' çıkmalı; basınca Şahıs dökümünde o yazım için doğrulanacak satır kalmamalı.",
  "Sorun görürseniz: paneldeki mesajın ekran görüntüsü, panelin altındaki sürüm numarası ve hangi adımda olduğu.",
];

const body = [
  ...check.map((t, k) => p(t, k === 0 ? "Heading2" : "")),
  `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`,
  p("GİRİŞ", "Heading1"),
  para(run("Ebû Hâmid el-Gazzâlî İslâm düşüncesinin en etkili isimlerinden biridir.") + foot("Gazzâlî, el-Mustasfâ, Beyrut 1993, I, 45.") + run(" Gazzâlî bu eserinde filozofları eleştirir; Gazzâlî'ye göre akıl sınırlıdır.")),
  para(run("Hocası Cüveynî") + vbaDate(" (ö. 478/1085)") + run(" Nizâmiye'de ders vermiştir.") + foot("Bu konuda Gazzâlî farklı düşünür; bk. Cüveynî, el-Burhân, I, 23.")),
  para(run("Fahreddin er-Râzî tefsirinde bu meseleyi tartışır. Râzî burada Zemahşerî'yi eleştirir. F. er-Râzî'nin bu görüşü sonraki müfessirleri etkilemiştir.")),
  para(run("Ebû Bekir er-Râzî ise tıp alanında meşhurdur. Râzî (ö. 313/925) bu konuda farklı düşünür; İbn Sînâ (v. 428/1037) da onu eleştirmiştir.") + end("İbn Sînâ'nın bu eleştirisi için bk. İbn Sînâ, eş-Şifâ, Kahire 1960, s. 12.")),
  p("Hayatı ve Eserleri", "Heading2"),
  para(run("Usûlî açıdan bakıldığında bu delil zayıftır. Bu meseleyi usûlî bakışla ele almak gerekir. Nakşibendî tarikatı Osmanlı'da yaygınlaşmıştır; Sünnî ve Şiî müellifler ayrışır.")),
  para(run("Irak'taki medreseler gelişmişti. Arap dili ve edebiyatı önemlidir. Şâfiî (ö. 204/920) bu görüşü er-Risâle'de savunur; Şâfiî mezhebinde ise durum farklıdır.")),
  para(run("Bu tartışma (Gazzâlî, 1993: 45) yaygındır. Bu görüş için bk. Gazzâlî, el-Mustasfâ, I, 45.") + foot("Gazzâlî, age, II, 45.")),
  para(run("Gazali'nin bu görüşü Fahrülislâm el-Pezdevî tarafından da tartışılmıştır. Gazali'ye göre ise mesele açıktır.")),
  p("KAYNAKÇA", "Heading1"),
  p("Gazzâlî, Ebû Hâmid. el-Mustasfâ min ilmi'l-usûl. Beyrut: Dâru'l-kütübi'l-ilmiyye, 1993."),
  p("İbn Haldûn. Mukaddime. Çev. Süleyman Uludağ. İstanbul: Dergâh Yayınları, 2004."),
  p("Cüveynî, el-Burhân fî usûli'l-fıkh, Katar 1399."),
].join("");

const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const sepNotes = (tag) =>
  `<w:${tag} w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:${tag}><w:${tag} w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${tag}>`;
const files = {
  "[Content_Types].xml":
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>' +
    '<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>',
  "_rels/.rels":
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  "word/_rels/document.xml.rels":
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>' +
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/></Relationships>',
  "word/styles.xml":
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${ns}>` +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:after="120" w:line="360" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="240"/><w:jc w:val="center"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/></w:rPr></w:style></w:styles>',
  "word/footnotes.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes ${ns}>${sepNotes("footnote")}${footnotes.join("")}</w:footnotes>`,
  "word/endnotes.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:endnotes ${ns}>${sepNotes("endnote")}${endnotes.join("")}</w:endnotes>`,
  "word/document.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${ns}><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr></w:body></w:document>`,
};
const out = new URL("../docs/vefeyat-deneme.docx", import.meta.url);
writeFileSync(out, zipSync(Object.fromEntries(Object.entries(files).map(([k, v]) => [k, strToU8(v)]))));
console.log(`docs/vefeyat-deneme.docx: ${fn} dipnot, ${en} sonnot`);
