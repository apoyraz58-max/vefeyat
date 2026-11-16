/**
 * Klavye kısayolları (keyboard-shortcut manifest only). Lists each Vefeyât
 * shortcut as Word currently has it, asks Word which key combinations are taken,
 * and lets the user move a command to a free one. Word keeps the choice per
 * user (Microsoft 365 sign-in); "Varsayılan" returns to shortcuts.json.
 */
import { $, setMessage } from "./core.js";

type Actions = {
  getShortcuts: () => Promise<Record<string, string | null>>;
  areShortcutsInUse: (keys: string[]) => Promise<{ shortcut: string; inUse: boolean }[]>;
  replaceShortcuts: (map: Record<string, string | null>) => Promise<void>;
};
type Definition = { actions: { id: string; name: string }[]; shortcuts: { action: string; key: { default: string } }[] };

const actions = (): Actions | null => {
  const a = (Office as unknown as { actions?: Partial<Actions> }).actions;
  return a?.getShortcuts && a.areShortcutsInUse && a.replaceShortcuts ? (a as Actions) : null;
};

/** Three-key combinations offered as alternatives (Ctrl+Alt is AltGr on a Turkish keyboard). */
export function candidateKeys(): string[] {
  const keys = [..."ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"];
  return [...keys.map((k) => `Alt+Shift+${k}`), ...keys.map((k) => `Ctrl+Shift+${k}`)];
}

const same = (a: string | null | undefined, b: string | null | undefined): boolean => (a ?? "").toUpperCase() === (b ?? "").toUpperCase();

let definition: Definition | null = null;

export async function renderShortcuts(): Promise<void> {
  const api = actions();
  const section = $("shortcuts");
  if (!api) return;
  try {
    if (!definition) {
      const response = await fetch("shortcuts.json", { cache: "no-cache" });
      definition = response.ok ? ((await response.json()) as Definition) : null;
    }
    if (!definition) return;
    const current = await api.getShortcuts();
    const ours = definition.actions.filter((a) => a.id in current);
    if (!ours.length) return; // regular manifest: Word knows no Vefeyât shortcuts
    const defaults = new Map(definition.shortcuts.map((s) => [s.action, s.key.default] as const));
    const taken = new Set(Object.values(current).filter(Boolean).map((k) => k!.toUpperCase()));
    const report = await api.areShortcutsInUse(candidateKeys());
    const busy = new Set(report.filter((r) => r.inUse).map((r) => r.shortcut.toUpperCase()));
    const list = $("shortcuts-list");
    list.replaceChildren();
    let conflicts = 0;
    for (const action of ours) {
      const key = current[action.id] ?? null;
      const li = document.createElement("li");
      const name = document.createElement("strong");
      name.textContent = action.name;
      const now = document.createElement("span");
      if (key === null) conflicts++;
      now.textContent = key === null ? " — kısayol Word'ün başka bir komutuna bırakılmış" : ` — ${key}`;
      const select = document.createElement("select");
      select.dataset.action = action.id;
      select.setAttribute("aria-label", `${action.name} kısayolu`);
      const add = (value: string, label: string): void => {
        const option = document.createElement("option");
        option.value = value;
        option.textContent = label;
        option.selected = same(value, key);
        select.append(option);
      };
      add(key ?? "", key ?? "seçin");
      add("default", `Varsayılan (${defaults.get(action.id) ?? "?"})`);
      for (const candidate of candidateKeys()) {
        const k = candidate.toUpperCase();
        if (same(candidate, key) || busy.has(k) || taken.has(k)) continue;
        add(candidate, `${candidate} (boşta)`);
      }
      li.append(name, now, document.createElement("br"), select);
      list.append(li);
    }
    $("shortcuts-summary").textContent =
      (conflicts ? `${conflicts} kısayol kullanılamıyor; yeni bir tuş seçin. ` : "") +
      "Bir kısayol başka şey yapıyorsa listeden boşta olan bir tuş seçip Kaydet'e basın. \"Boşta\" yazanlar Word'ün bildiğine göre başka komuta bağlı değil.";
    section.hidden = false;
  } catch {
    section.hidden = true;
  }
}

export async function saveShortcuts(): Promise<void> {
  const api = actions();
  if (!api) return;
  const current = await api.getShortcuts();
  const changes: Record<string, string | null> = {};
  for (const select of Array.from(document.querySelectorAll<HTMLSelectElement>("#shortcuts-list select"))) {
    const id = select.dataset.action!;
    if (select.value === "default") changes[id] = null;
    else if (select.value && !same(select.value, current[id])) changes[id] = select.value;
  }
  const chosen = Object.values(changes).filter(Boolean).map((k) => k!.toUpperCase());
  if (new Set(chosen).size !== chosen.length) {
    setMessage("Aynı tuşu iki komuta veremezsiniz.", "warning");
    return;
  }
  if (!Object.keys(changes).length) {
    setMessage("Değişiklik yok.", "info");
    return;
  }
  try {
    await api.replaceShortcuts(changes);
    setMessage("Kısayollar kaydedildi. Düğme açıklamalarında varsayılan tuşlar yazmaya devam eder.", "success");
  } catch {
    setMessage("Word kısayolu kaydetmedi. Microsoft 365 hesabıyla oturum açık olmalı.", "error");
  }
  await renderShortcuts();
}
