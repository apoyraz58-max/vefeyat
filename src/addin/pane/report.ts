/**
 * Sorun bildir: a short report for the closed beta, copied or mailed by the
 * user. It holds the versions, what this Word supports, the settings, counts
 * and the last messages and errors of this session; never the document text.
 */
import { $, events, noteEvent, setMessage, state } from "./core.js";
import { SETTING_CHOICES } from "../logic.js";
import { summarize } from "../inventory.js";

declare const __VERSION__: string;

export const REPORT_ADDRESS = "apoyraz@cumhuriyet.edu.tr";

/** Page errors go into the report too. */
export function watchErrors(): void {
  window.addEventListener("error", (event) => noteEvent("hata", `${event.message} (${(event.filename ?? "").split("/").pop()}:${event.lineno})`));
  window.addEventListener("unhandledrejection", (event) => {
    const reason = event.reason as { message?: string } | string | undefined;
    noteEvent("hata", typeof reason === "string" ? reason : (reason?.message ?? String(reason)));
  });
}

const SETS: Array<[string, string[]]> = [
  ["WordApi", ["1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "1.9"]],
  ["WordApiDesktop", ["1.1", "1.2", "1.3", "1.4"]],
  ["SharedRuntime", ["1.1"]],
  ["KeyboardShortcuts", ["1.1"]],
];

function supported(): string {
  const requirements = Office.context.requirements;
  return SETS.map(([name, versions]) => {
    const highest = versions.filter((v) => requirements.isSetSupported(name, v)).pop();
    return `${name} ${highest ?? "yok"}`;
  }).join(", ");
}

export function buildReport(note: string): string {
  const diagnostics = (Office.context as unknown as { diagnostics?: { host?: string; platform?: string; version?: string } }).diagnostics;
  const settings = state.settings;
  const rows = state.rows;
  const sum = rows.length ? summarize(rows, settings.scope) : null;
  const notes = state.segments.filter((s) => s.part === "note").length;
  const lines = [
    "Vefeyât sorun bildirimi",
    "",
    `Ne oldu: ${note.trim() || "(yazılmadı)"}`,
    "",
    `Vefeyât: ${__VERSION__}`,
    `Word: ${diagnostics?.host ?? "?"} ${diagnostics?.version ?? "?"} · ${diagnostics?.platform ?? "?"}`,
    `Destek: ${supported()}`,
    `Tarayıcı: ${navigator.userAgent}`,
    `Ayarlar: Kısaltma ${SETTING_CHOICES.prefix[settings.prefix]} · ${SETTING_CHOICES.scope[settings.scope]} · Ad biçimi ${SETTING_CHOICES.nameForm[settings.nameForm]}`,
    `Katalog: ${state.catalog ? `${state.catalog.people.length} şahıs (${state.catalogSource === "user" ? "kendi CSV'si" : "eklentiyle gelen"})` : "yüklenmedi"}`,
    `Belge (son okuma): ${state.segments.length ? `${state.segments.length - notes} paragraf, ${notes} not` : "okunmadı"}`,
    sum ? `Döküm: ${sum.mentions} geçiş, ${sum.people} kesin şahıs, ${sum.uncertain} doğrulanacak, ${sum.firstWithoutDate} ilk geçişte tarih yok` : "Döküm: hazırlanmadı",
    "",
    "Son mesajlar ve hatalar:",
    ...(events.length ? events : ["(yok)"]),
    "",
    "Belge metni bu bildirime eklenmez. Gerekirse sorunu gösteren kısa bir örneği kendiniz ekleyebilirsiniz.",
  ];
  return lines.join("\n");
}

export function renderReport(): void {
  $<HTMLTextAreaElement>("report-text").value = buildReport($<HTMLTextAreaElement>("report-note").value);
}

export async function copyReport(): Promise<void> {
  renderReport();
  try {
    await navigator.clipboard.writeText($<HTMLTextAreaElement>("report-text").value);
    setMessage(`Bildirim panoya kopyalandı. E-postanıza yapıştırıp ${REPORT_ADDRESS} adresine gönderin.`, "success");
  } catch {
    $<HTMLTextAreaElement>("report-text").select();
    setMessage("Panoya kopyalanamadı; metin seçili, Ctrl+C ile kopyalayın.", "warning");
  }
}

/** Opens a new e-mail with the report; mail programs cut long links, so the text is kept short. */
export function mailReport(): void {
  renderReport();
  const body = $<HTMLTextAreaElement>("report-text").value;
  const short = body.length > 1800 ? `${body.slice(0, 1800)}\n…(kısaltıldı; tamamı için Panoya kopyala)` : body;
  const link = $<HTMLAnchorElement>("report-mail");
  link.href = `mailto:${REPORT_ADDRESS}?subject=${encodeURIComponent(`Vefeyât ${__VERSION__} sorun bildirimi`)}&body=${encodeURIComponent(short)}`;
}
