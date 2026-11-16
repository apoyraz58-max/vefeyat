/**
 * Vefeyât pane entry: launch and wiring. The views live in ./pane/.
 */
import { type Launch, earlierSpellings, parseLaunch, quickInsertBlocker } from "./logic.js";
import { GO_FUNCTIONS, goQuery, register, registerSettingCommands } from "./actions.js";
import { canReadPages, selectionPage, writeAuditBookmark, findEarlier, onSelectionChanged, selectMention, writeConfirmations } from "./word.js";

import { chooseCatalog, resetCatalog, restoreCatalog, runSearch, addLocal, renderLocals } from "./pane/catalog.js";
import { $, PAGE_SIZE, clearHistory, history, requestCancel, run, setMessage, showView, state } from "./pane/core.js";
import { checkHealth, convertLegacy, repairHealth, repairSource, showSource } from "./pane/health.js";
import { actOnRow, gotoAudit, prepareInventory, renderAudit, renderInventory, resumeAudit, startAudit, copyInventory, exportInventory } from "./pane/inventory.js";
import { addDate, earlierRow, hideEarlier, preferSelection, refresh, renderSpellingChoice, respell, toggleExclusion } from "./pane/selected.js";
import { cleanup, createIndex, createPersonReport, deleteIndex, deleteIndexMarks, showConsistency } from "./pane/submit.js";
import { applyBulk, applyPrefs, clearSpellings, copyPrefs, downloadPrefs, noProofing, prepareBulk, renderExclusions, renderHistory, renderIdentities, renderSpellings, showCoverage, toggleFrames, checkFirstUse } from "./pane/tools.js";

declare const __VERSION__: string;

/** Carries out what the ribbon or context-menu button asked for. */
export async function launch(request: Launch): Promise<void> {
  showView(request.view);
  if (request.tool) document.getElementById(`tool-${request.tool === "dizinkaldir" || request.tool === "dizinsil" ? "dizin" : request.tool}`)?.classList.add("focus");
  if (request.view === "ara") $<HTMLInputElement>("search-query").focus();
  if (request.view === "dokum") await run(prepareInventory, "Şahıs dökümü hazırlanıyor");
  if (request.view === "saglik") await run(checkHealth, "Tarihler denetleniyor");
  if (request.view === "toplu") await run(prepareBulk, "Öneriler hazırlanıyor");
  if (request.view === "gecmis") renderHistory();
  if (request.view === "kimlik") renderIdentities();
  if (request.view === "katalog") renderLocals();
  if (request.view === "tercihler") renderSpellings();
  if (request.action === "prefer") await run(preferSelection);
  if (request.action === "clearSpelling") await run(clearSpellings);
  if (request.action === "shading") await run(toggleFrames);
  if (request.action === "noproof") await run(noProofing, "Şahıs adları işaretleniyor");
  if (request.view === "istisna") renderExclusions();
  if (request.view === "kapsam") await run(showCoverage);
  if (request.action === "resume") await run(resumeAudit, "Belge taranıyor");
  if (request.view === "kaynak") await run(showSource);
  if (request.tool === "tutarlilik") await run(showConsistency, "Tutarlılık raporu");
  if (request.tool === "rapor") await run(createPersonReport, "Şahıs raporu hazırlanıyor");
  if (request.tool === "dizin") await run(createIndex, "Şahıslar dizini oluşturuluyor");
  if (request.tool === "dizinkaldir") await run(deleteIndexMarks, "Dizin işaretleri kaldırılıyor");
  if (request.tool === "dizinsil") await run(deleteIndex, "Dizin kaldırılıyor");
  if (request.action === "audit") await run(startAudit, "Belge taranıyor");
  if (request.action === "dia") {
    const person = state.chosen;
    if (person && person.source.startsWith("https://islamansiklopedisi.org.tr/")) {
      setMessage("DİA maddesini açmak için aşağıdaki bağlantıya tıklayın.", "info");
      $("detail-link").focus();
    } else {
      setMessage(quickInsertBlocker(state.catalog !== null, state.alias, state.found) ?? "Bu şahıs için DİA bağlantısı yok.", "info");
    }
  }
  if (request.action === "insert") {
    const blocker = quickInsertBlocker(state.catalog !== null, state.alias, state.found);
    if (blocker) setMessage(blocker, "info");
    else await run(() => addDate(false));
  } else if (request.action === "spelling") {
    if (state.chosen && !$<HTMLButtonElement>("respell").disabled) $<HTMLSelectElement>("spelling").focus();
    else setMessage(quickInsertBlocker(state.catalog !== null, state.alias, state.found) ?? "Bu ad için farklı bir yazım önerisi yok.", "info");
  }
}

/**
 * Word may open the pane from its own cache long after an update. The
 * published version.json is read past every cache; when it names a newer
 * build, the pane reloads itself once with that version in the address,
 * which also fetches the new script and style (their URLs carry it).
 */
export async function checkForUpdate(): Promise<void> {
  try {
    const response = await fetch(`version.json?t=${Date.now()}`, { cache: "no-store" });
    if (!response.ok) return;
    const { version } = (await response.json()) as { version?: string };
    if (!version || version === __VERSION__) return;
    const url = new URL(window.location.href);
    if (url.searchParams.get("_v") === version) {
      setMessage(`Yeni sürüm (${version}) yayında ama Word eski kopyayı açtı (${__VERSION__}). Word'ü kapatıp açın; düzelmezse Wef klasörünü temizleyin.`, "warning");
      return;
    }
    url.searchParams.set("_v", version);
    window.location.replace(url.toString());
  } catch {
    /* offline or blocked: keep the running copy */
  }
}

export function initialise(): void {
  void checkForUpdate();
  $("version").textContent = __VERSION__;
  $("about-version").textContent = __VERSION__;
  const file = $<HTMLInputElement>("catalog-file");
  $("choose-catalog").addEventListener("click", () => file.click());
  $("choose-catalog-inline").addEventListener("click", () => file.click());
  $("reset-catalog").addEventListener("click", () => void resetCatalog());
  $("run-tutarlilik").addEventListener("click", () => void run(showConsistency, "Tutarlılık raporu"));
  $("run-rapor").addEventListener("click", () => void run(createPersonReport, "Şahıs raporu hazırlanıyor"));
  $("spellings-clear").addEventListener("click", () => void run(clearSpellings));
  $("prefs-download").addEventListener("click", downloadPrefs);
  $("prefs-copy").addEventListener("click", () => void copyPrefs());
  $("prefs-apply").addEventListener("click", () => void run(applyPrefs));
  $("local-add").addEventListener("click", () => void run(addLocal));
  $("frames-toggle").addEventListener("click", () => void run(toggleFrames));
  $("noproof-run").addEventListener("click", () => void run(noProofing, "Şahıs adları işaretleniyor"));
  const prefsFile = $<HTMLInputElement>("prefs-file");
  $("prefs-choose").addEventListener("click", () => prefsFile.click());
  prefsFile.addEventListener("change", () => {
    const chosen = prefsFile.files?.[0];
    if (chosen) void chosen.text().then((text) => {
      $<HTMLTextAreaElement>("prefs-text").value = text;
      prefsFile.value = "";
      void run(applyPrefs);
    });
  });
  $("not-person").addEventListener("click", () => void run(toggleExclusion));
  $("identity-clear").addEventListener("click", () =>
    void run(async () => {
      await writeConfirmations({});
      renderIdentities();
    }),
  );
  $("bulk-refresh").addEventListener("click", () => void run(prepareBulk, "Öneriler hazırlanıyor"));
  $("bulk-apply").addEventListener("click", () => void run(applyBulk, "Öneriler uygulanıyor"));
  $("history-clear").addEventListener("click", () => {
    clearHistory();
    renderHistory();
  });
  $("legacy-convert").addEventListener("click", () => void run(convertLegacy, "VBA tarihleri dönüştürülüyor"));
  $("health-refresh").addEventListener("click", () => void run(checkHealth, "Tarihler denetleniyor"));
  $("health-repair").addEventListener("click", () => void run(repairHealth));
  $("source-repair").addEventListener("click", () => void run(repairSource));
  $("busy-cancel").addEventListener("click", () => {
    requestCancel();
    $("busy-text").textContent = "İptal ediliyor (Word'ün elindeki parça bitince durur)…";
  });
  $("index-pages").hidden = !canReadPages();
  try {
    const saved = window.localStorage.getItem("vefeyat.indexFirst");
    if (saved) $<HTMLInputElement>("index-first").value = saved;
  } catch {
    /* ignored */
  }
  $("index-here").addEventListener("click", () =>
    void run(async () => {
      const page = await selectionPage();
      if (page) {
        $<HTMLInputElement>("index-first").value = String(page);
        setMessage(`İmleç belgenin ${page}. sayfasında; bu sayfa 1 numaralı kabul edilecek.`, "success");
      }
    }),
  );
  $("run-dizin").addEventListener("click", () => void run(createIndex, "Şahıslar dizini oluşturuluyor"));
  $("run-dizinkaldir").addEventListener("click", () => void run(deleteIndexMarks, "Dizin işaretleri kaldırılıyor"));
  $("run-dizinsil").addEventListener("click", () => void run(deleteIndex, "Dizin kaldırılıyor"));
  $("run-temizlik").addEventListener("click", () => void run(cleanup, "Teslim temizliği"));
  file.addEventListener("change", () => {
    const chosen = file.files?.[0];
    if (chosen) void chooseCatalog(chosen);
    file.value = "";
  });
  $("back").addEventListener("click", () => {
    showView("secili");
    state.alias = null;
    void refresh();
  });
  $("insert").addEventListener("click", () => void run(() => addDate(false)));
  $("insert-anyway").addEventListener("click", () => void run(() => addDate(true)));
  $("go-earlier").addEventListener("click", () =>
    void run(async () => {
      const row = earlierRow;
      const segment = row ? state.segments[row.segment] : undefined;
      if (row && segment && !(await selectMention(segment, row.start, row.text))) {
        setMessage("Önceki geçiş metinde bulunamadı; metin değişmiş olabilir.", "warning");
      } else if (!row && state.chosen && state.alias) {
        await findEarlier(earlierSpellings(state.chosen, state.alias), true);
      }
      hideEarlier();
    }),
  );
  $("respell").addEventListener("click", () => void run(respell));
  $("spelling").addEventListener("change", renderSpellingChoice);
  let timer: number | undefined;
  let upkeepTimer: number | undefined;
  onSelectionChanged(() => {
    window.clearTimeout(timer);
    timer = window.setTimeout(() => {
      if (!$("view-kaynak").hidden) void run(showSource);
      else void refresh();
    }, 250);
    // Typing moves the selection: once it rests, look whether a first mention moved.
    window.clearTimeout(upkeepTimer);
    upkeepTimer = window.setTimeout(() => void checkFirstUse().catch(() => {}), 2000);
  });
  $("firstuse-review").addEventListener("click", () => {
    $("firstuse-notice").hidden = true;
    showView("toplu");
    void run(prepareBulk, "Öneriler hazırlanıyor");
  });
  $("inventory-refresh").addEventListener("click", () => void run(prepareInventory, "Şahıs dökümü hazırlanıyor"));
  $("inventory-act").addEventListener("click", () => void actOnRow());
  $("inventory-xlsx").addEventListener("click", exportInventory);
  $("inventory-copy").addEventListener("click", () => void copyInventory());
  $("inventory-prev").addEventListener("click", () => {
    if (state.page > 0) state.page--;
    renderInventory();
  });
  $("inventory-next").addEventListener("click", () => {
    if ((state.page + 1) * PAGE_SIZE < state.rows.length) state.page++;
    renderInventory();
  });
  $("audit-prev").addEventListener("click", () => void run(() => gotoAudit(state.auditAt - 1)));
  $("audit-next").addEventListener("click", () =>
    void run(async () => {
      if (state.audit && state.auditAt + 1 < state.audit.length) await gotoAudit(state.auditAt + 1);
      else {
        state.audit = null;
        void writeAuditBookmark(null);
        renderAudit();
        setMessage("Denetim tamamlandı.", "success");
      }
    }),
  );
  $("audit-stop").addEventListener("click", () => {
    state.audit = null;
    renderAudit();
    setMessage("Denetim duraklatıldı. Kaldığınız yerden sürdürmek için Vefeyât → Denetime devam et.", "info");
  });
  $("search-form").addEventListener("submit", (event) => {
    event.preventDefault();
    runSearch();
  });
  const request = parseLaunch(window.location.search);
  showView(request.view);
  ready = restoreCatalog()
    .then(() => refresh())
    .catch(() => undefined);
  void ready.then(() => launch(request));
  registerRibbonFunctions();
}

let ready: Promise<void> = Promise.resolve();

/**
 * Keyboard-shortcut manifest: ribbon buttons, the context menu and the
 * shortcut run as functions in this (shared) runtime instead of reloading the
 * pane with a new address. Each shows the pane, reads the current selection
 * and does what ?action=…/?view=… would have done. In the regular manifest
 * nothing calls these.
 */
function registerRibbonFunctions(): void {
  for (const name of GO_FUNCTIONS) {
    const query = goQuery(name)!;
    register(name, (event) => {
      const addin = (Office as unknown as { addin?: { showAsTaskpane: () => Promise<void> } }).addin;
      void (addin?.showAsTaskpane() ?? Promise.resolve())
        .catch(() => undefined)
        .then(() => event?.completed());
      void ready
        .then(() => refresh())
        .then(() => launch(parseLaunch(query)))
        .catch(() => undefined);
    });
  }
  registerSettingCommands(() => refresh());
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) initialise();
  else setMessage("Vefeyât yalnız Word içinde çalışır.", "error");
});
