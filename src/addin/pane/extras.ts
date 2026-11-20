/**
 * Ek kataloglar: field catalogs built from Wikidata (sociology and philosophy,
 * Ottoman-Turkish history, literature, economics and politics), published next
 * to the DİA catalog and added to it on the user's choice. DİA is
 * authoritative: an entry with the same name and death year as a DİA entry is
 * left out. The choice is kept in this browser.
 */
import { Catalog, normalize } from "../../core/index.js";
import { $, setMessage, state } from "./core.js";

export interface ExtraCatalog {
  readonly id: string;
  readonly title: string;
  readonly file: string;
  readonly count: number;
}

const CHOICE_KEY = "vefeyat.extraCatalogs";
let available: ExtraCatalog[] | null = null;
const loaded = new Map<string, Catalog>();

/**
 * The chosen extra catalogs. Until the user changes the choice every catalog is
 * on: the field catalogs work from the first run, with nothing to tick.
 */
export async function chosenExtras(): Promise<string[]> {
  let stored: unknown = null;
  try {
    const raw = window.localStorage.getItem(CHOICE_KEY);
    stored = raw === null ? null : (JSON.parse(raw) as unknown);
  } catch {
    stored = null;
  }
  if (Array.isArray(stored)) return stored.filter((v): v is string => typeof v === "string");
  return (await availableExtras()).map((e) => e.id);
}

function saveChoice(ids: string[]): void {
  try {
    window.localStorage.setItem(CHOICE_KEY, JSON.stringify(ids));
  } catch {
    // Storage blocked: the choice lasts for this session only.
  }
}

export async function availableExtras(): Promise<ExtraCatalog[]> {
  if (available) return available;
  try {
    const response = await fetch("catalogs.json", { cache: "no-cache" });
    available = response.ok ? ((await response.json()) as ExtraCatalog[]) : [];
  } catch {
    available = [];
  }
  return available;
}

async function load(extra: ExtraCatalog): Promise<Catalog> {
  const known = loaded.get(extra.id);
  if (known) return known;
  const response = await fetch(extra.file, { cache: "no-cache" });
  if (!response.ok) throw new Error(`${extra.title} indirilemedi (${response.status}).`);
  const catalog = Catalog.fromCsv(await response.text());
  loaded.set(extra.id, catalog);
  // One catalog at a time: the window gets a moment between them.
  await new Promise((resolve) => setTimeout(resolve, 0));
  return catalog;
}

const lastYear = (date: string): string => (date.match(/\d{1,4}/g) ?? []).pop() ?? "";
const key = (name: string, date: string): string => `${normalize(name.includes(",") ? `${name.split(",")[1]} ${name.split(",")[0]}` : name).replace(/ /g, "")}|${lastYear(date)}`;

/** The chosen extra catalogs, without the people the base catalog (DİA) or an earlier extra already has. */
export async function extraCatalogs(base: Catalog): Promise<Catalog[]> {
  const ids = new Set(await chosenExtras());
  const list = (await availableExtras()).filter((e) => ids.has(e.id));
  if (list.length === 0) return [];
  const taken = new Set(base.people.map((p) => key(p.shortName, p.deathDate)));
  const out: Catalog[] = [];
  for (const extra of list) {
    const catalog = await load(extra);
    out.push(
      catalog.without((person) => {
        const k = key(person.shortName, person.deathDate);
        if (taken.has(k)) return true;
        taken.add(k);
        return false;
      }),
    );
  }
  return out;
}

/** The Ek kataloglar section of the Katalog window. */
export async function renderExtras(onChange: () => Promise<void>): Promise<void> {
  const list = await availableExtras();
  const section = $("extras");
  section.hidden = list.length === 0;
  const box = $("extras-list");
  box.replaceChildren();
  const chosen = new Set(await chosenExtras());
  for (const extra of list) {
    const label = document.createElement("label");
    label.className = "check";
    const input = document.createElement("input");
    input.type = "checkbox";
    input.checked = chosen.has(extra.id);
    input.addEventListener("change", async () => {
      const ids = new Set(await chosenExtras());
      if (input.checked) ids.add(extra.id);
      else ids.delete(extra.id);
      saveChoice([...ids]);
      void onChange().then(() =>
        setMessage(input.checked ? `${extra.title} kataloğu eklendi (${extra.count.toLocaleString("tr-TR")} kişi).` : `${extra.title} kataloğu çıkarıldı.`, "success"),
      );
    });
    label.append(input, ` ${extra.title} `);
    const count = document.createElement("span");
    count.className = "muted";
    count.textContent = `(${extra.count.toLocaleString("tr-TR")} kişi)`;
    label.append(count);
    box.append(label);
  }
  void state;
}
