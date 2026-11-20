/** Vefeyât pane — Teslim ve dizin: Tutarlılık raporu, Şahıs raporu, Şahıslar dizini, Teslim öncesi temizle. */
import { NameIndex } from "../../core/index.js";
import { buildInventoryAsync } from "../inventory.js";
import { consistencyReport, indexRows, personReportDocx, toBase64 } from "../submit.js";
import { buildPageIndex, canReadPages, StepError, readSpellings, readExclusions, buildIndex, removeIndexMarks, canIndex, documentName, finalizeDocument, removeIndex, openNewDocument, readConfirmations, readDocument } from "../word.js";

import { $, logHistory, progress, run, setMessage, state, ensureIndex } from "./core.js";

// ---- Teslim ve dizin ----

export async function showConsistency(): Promise<void> {
  const { segments, footnotes } = await readDocument();
  const report = consistencyReport(segments);
  const out = $("consistency-out");
  out.hidden = false;
  out.replaceChildren();
  const line = (text: string, cls = "") => {
    const p = document.createElement("p");
    p.textContent = text;
    if (cls) p.className = cls;
    out.append(p);
  };
  if (report.total === 0) {
    line("Belgede (ö. …) biçiminde tarih bulunamadı.");
  } else {
    line(`${report.total} tarih incelendi${footnotes ? " (dipnot ve sonnotlar dahil)" : " (dipnot ve sonnotlar bu Word sürümünde okunamadı)"}.`);
    for (const [title, entries, mixed] of [["İşaret", report.markers, report.mixedMarkers], ["Biçim", report.kinds, report.mixedKinds]] as const) {
      line(`${title}: ${entries.map((e) => `${e.label} ${e.count}`).join(" · ")}${mixed ? " — karışık" : " — tutarlı"}`, mixed ? "warning" : "");
      for (const e of entries) for (const x of e.examples) line(`  ${x}`, "muted");
    }
  }
  setMessage(report.mixedMarkers || report.mixedKinds ? "Tarih yazımında karışıklık var; örnekler aşağıda." : "Tutarlılık raporu hazır.", report.mixedMarkers || report.mixedKinds ? "warning" : "success");
}

export async function createPersonReport(): Promise<void> {
  if (!state.catalog) {
    setMessage("Katalog yüklenmedi.", "warning");
    return;
  }
  const read = await readDocument();
  await ensureIndex();
  const rows = await buildInventoryAsync(read.segments, state.catalog, state.index ?? undefined, readConfirmations(), readExclusions(), readSpellings(), progress("Şahıslar taranıyor"));
  await openNewDocument(toBase64(personReportDocx(documentName(), rows)));
  setMessage("Şahıs raporu yeni bir Word belgesi olarak açıldı.", "success");
}

export async function createIndex(): Promise<void> {
  if (!canIndex()) {
    setMessage("Bu Word sürümü alan eklemeyi desteklemiyor (WordApi 1.5 gerekir). Microsoft 365 Word'ü güncelleyin.", "warning");
    return;
  }
  if (!state.catalog) {
    setMessage("Katalog yüklenmedi.", "warning");
    return;
  }
  const read = await readDocument(progress("Belge okunuyor (dipnotlar)")).catch((error: unknown) => {
    throw new StepError("Belge okunamadı", error);
  });
  await ensureIndex();
  const withNotes = $<HTMLInputElement>("index-notes").checked;
  const rows = indexRows(await buildInventoryAsync(read.segments, state.catalog, state.index ?? undefined, readConfirmations(), readExclusions(), readSpellings(), progress("Şahıslar taranıyor"))).filter(
    (r) => withNotes || r.part === "main",
  );
  if (rows.length === 0) {
    setMessage("Dizine girecek kimliği kesin geçiş yok. Belirsiz adları Şahıs dökümünde onaylayın.", "info");
    return;
  }
  const people = new Set(rows.map((r) => r.person?.source)).size;
  setMessage("", "info");
  if (canReadPages()) {
    const firstArabic = Math.max(1, Number($<HTMLInputElement>("index-first").value) || 1);
    try {
      window.localStorage.setItem("vefeyat.indexFirst", String(firstArabic));
    } catch {
      /* ignored */
    }
    const t0 = Date.now();
    const result = await buildPageIndex(read.segments, rows, firstArabic, progress("Sayfa numaraları okunuyor"));
    const seconds = Math.round((Date.now() - t0) / 1000);
    logHistory("index", `${result.people} şahıs, sayfalardan, ${seconds} sn`);
    setMessage(
      `Şahıslar dizini hazır: ${result.people} şahıs, ${seconds} sn. Belgeye işaret konmadı; sayfa numaraları sabittir, metin değişince dizini yeniden oluşturun.` +
        (result.missing ? ` ${result.missing} geçişin sayfası okunamadı.` : ""),
      result.missing ? "warning" : "success",
    );
    return;
  }
  const started = Date.now();
  const result = await buildIndex(read.segments, rows, progress("Dizin işaretleri ekleniyor"));
  const minutes = Math.round((Date.now() - started) / 6000) / 10;
  logHistory("index", `${people} şahıs, ${result.marks} işaret, ${minutes} dk`);
  const pages = " [sayfa API: yok; dizin işaretlerle hazırlandı]";
  setMessage(
    `Şahıslar dizini hazır: ${people} şahıs, ${result.marks} işaret, ${minutes} dk.` +
      (result.missing ? ` ${result.missing} geçiş metinde bulunamadı.` : "") +
      " Dizin işaretleri gizli metindir; ¶ düğmesiyle görünür." +
      pages,
    result.missing ? "warning" : "success",
  );
}

export async function deleteIndexMarks(): Promise<void> {
  if (!canIndex()) {
    setMessage("Bu Word sürümü alanları düzenlemeyi desteklemiyor (WordApi 1.5 gerekir).", "warning");
    return;
  }
  const result = await removeIndexMarks();
  if (!result.index && result.marks === 0) setMessage("Belgede Vefeyât dizini yok.", "info");
  else
    setMessage(
      `${result.marks} dizin işareti kaldırıldı.` + (result.index ? " Şahıslar Dizini güncel hâliyle metin olarak kaldı; sayfa numaraları artık kendiliğinden güncellenmez." : ""),
      "success",
    );
  logHistory("index", `işaretler kaldırıldı, dizin metin oldu (${result.marks} işaret)`);
}

export async function deleteIndex(): Promise<void> {
  if (!canIndex()) {
    setMessage("Bu Word sürümü alanları düzenlemeyi desteklemiyor (WordApi 1.5 gerekir).", "warning");
    return;
  }
  const result = await removeIndex();
  if (!result.index && result.marks === 0) setMessage("Belgede Vefeyât dizini yok.", "info");
  else setMessage(`${result.marks} dizin işareti${result.index ? " ve Şahıslar Dizini" : ""} kaldırıldı.`, "success");
  logHistory("index", `dizin kaldırıldı (${result.marks} işaret)`);
}

export let cleanupArmed = false;

export async function cleanup(): Promise<void> {
  const button = $<HTMLButtonElement>("run-temizlik");
  if (!cleanupArmed) {
    cleanupArmed = true;
    button.textContent = "Eminim, temizle";
    setMessage("Bu işlem geri alınamaz; teslim kopyasında çalıştığınızdan emin olun ve tekrar tıklayın.", "warning");
    return;
  }
  cleanupArmed = false;
  button.textContent = "Temizle";
  const result = await finalizeDocument();
  logHistory("cleanup", `${result.controls} tarih metne dönüştü`);
  setMessage(`Temizlendi: ${result.controls} tarih denetimi metne dönüştürüldü, ${result.marks} dizin işareti kaldırıldı (dizin metin olarak kaldı), ${result.settings} belge ayarı kaldırıldı.`, "success");
}
