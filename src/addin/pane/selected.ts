/** Vefeyât pane — the selected-name window: candidates, dates, spellings, earlier mentions, Şahıs değil. */
import { isLocal } from "../persons.js";
import { applyLocalPersons, fillLocalForm, renderLocals } from "./catalog.js";
import { type WikidataHit, searchWikidata } from "../wikidata-live.js";
import { addLocals, localFromHit, standardLabel } from "./wikidata.js";
import { type NameIndex, type Person, alreadyHasDate, findMatches, formatDate, isValidDate, nameKey, nameOptions, nameQuery, normalize, personKey, selectionAlias, textName, withEnding } from "../../core/index.js";
import { sourceLabel, addConfirmation, confirmedPerson, exclusionKey, mentionKey, candidateLabel, dateTag, earlierSpellings, leadingDate, lookup, matchHint, orderByNameForm, planDate, quickInsertBlocker, sameDate, scopeRule } from "../logic.js";
import { type InventoryRow, placeOf } from "../inventory.js";
import { rowKey } from "../bulk.js";
import { preferenceOf, preferredSpelling, withPreferred } from "../prefs.js";
import { readSpellings, writeSpellings, readExclusions, writeExclusions, firstMentionBefore, findEarlier, insertDate, readConfirmations, readSelection, replaceSelection, replaceNameInSelection, writeConfirmations, selectionInNote, updateDateAfterSelection, removeDateAfter, selectMention, replaceMentions } from "../word.js";

import { $, documentRows, loadSettings, logHistory, setMessage, state, versions, showView } from "./core.js";

// ---- Selected name ----

export let contextNote: string | null = null;

/**
 * The last Şahıs dökümü or denetim already read the whole document: use its
 * reading of this mention (identity from the written date or from the other
 * mentions; otherwise the order of the candidates).
 */
export function applyDocumentContext(mention: string): void {
  if (!state.found || state.rows.length === 0) return;
  const row = state.rows.find((r) => rowKey(r, state.segments) === mention);
  if (!row) return;
  const order = new Map(row.candidates.map((c, k) => [personKey(c), k] as const));
  const found = state.found;
  const candidates = [...found.candidates].sort((a, b) => (order.get(personKey(a)) ?? 999) - (order.get(personKey(b)) ?? 999));
  state.found = { ...found, candidates };
  if (row.person && (row.resolved === "date" || row.resolved === "context" || row.resolved === "topic")) {
    state.chosen = candidates.find((c) => personKey(c) === personKey(row.person as Person)) ?? null;
    if (state.chosen) {
      contextNote =
        row.resolved === "date" ? "yazılı tarihe göre" : row.resolved === "topic" ? `cümlenin konusuna göre (${row.topic}); doğrulayın` : "belgedeki diğer geçişlere göre; doğrulayın";
    }
  }
}

export function renderCandidates(): void {
  const list = $("candidates");
  list.replaceChildren();
  const found = state.found;
  if (!found) return;
  if (found.candidates.length === 0) {
    const empty = document.createElement("p");
    empty.className = "muted";
    empty.textContent = "Katalogda bu adla eşleşen şahıs yok.";
    list.append(empty);
    // Kendi şahıslarım: the selected name, ready to complete with a date.
    const add = document.createElement("button");
    add.type = "button";
    add.textContent = "Kendi şahıslarıma ekle";
    add.addEventListener("click", () => {
      showView("katalog");
      renderLocals();
      fillLocalForm({ shortName: nameQuery(state.alias ?? "") });
    });
    list.append(add);
    list.append(wikidataSearch(true));
    return;
  }
  // A long list is shortened: with a chosen person only that one, otherwise the first five;
  // the rest opens on request. Ölüm tarihi ekle then stays close below.
  if (expandedFor !== state.alias) {
    expandedFor = state.alias;
    filterText = "";
    expanded = false;
  }
  const chosenKey = state.chosen ? personKey(state.chosen) : null;
  const shownIndexes = found.candidates
    .map((person, index) => ({ person, index }))
    .filter(({ person, index }) => expanded || (chosenKey ? personKey(person) === chosenKey : index < CANDIDATE_PREVIEW));
  const hidden = found.candidates.length - shownIndexes.length;
  // A long open list gets a search box: a few letters of the name, place or field narrow it.
  const rows: HTMLElement[] = [];
  let filter: HTMLInputElement | null = null;
  if (expanded && shownIndexes.length > CANDIDATE_FILTER_FROM) {
    filter = document.createElement("input");
    filter.type = "search";
    filter.id = "candidate-filter";
    filter.className = "wide candidate-filter";
    filter.placeholder = "Adaylarda ara (ad, yüzyıl, alan…)";
    filter.setAttribute("aria-label", "Adaylarda ara");
    filter.value = filterText;
    const input = filter;
    input.addEventListener("input", () => {
      filterText = input.value;
      applyFilter(rows, filterText);
    });
    list.append(filter);
  }
  shownIndexes.forEach(({ person, index }) => {
    const id = `candidate-${index}`;
    const row = document.createElement("label");
    row.className = "candidate";
    row.htmlFor = id;
    const radio = document.createElement("input");
    radio.type = "radio";
    radio.name = "candidate";
    radio.id = id;
    radio.checked = state.chosen !== null && personKey(state.chosen) === personKey(person);
    row.classList.toggle("chosen", radio.checked);
    row.addEventListener("pointerdown", () => {
      pointerChoice = true;
    });
    radio.addEventListener("change", () => {
      state.chosen = person;
      hideEarlier();
      if (pointerChoice) {
        // A click settles the choice: the list shortens to the chosen person.
        pointerChoice = false;
        expanded = false;
        renderCandidates();
      } else {
        // Arrow keys walk the list: it stays as it is, only the mark moves.
        markChosen(list, row);
      }
      renderDetails();
    });
    radio.addEventListener("keydown", (event) => {
      pointerChoice = false;
      // Enter on a candidate adds the date (when it can be added straight away).
      if (event.key !== "Enter") return;
      const insert = document.getElementById("insert") as HTMLButtonElement | null;
      if (insert && !insert.disabled && !insert.closest("[hidden]")) {
        event.preventDefault();
        insert.click();
      }
    });
    const text = document.createElement("span");
    const title = document.createElement("strong");
    title.textContent = candidateLabel(person);
    text.append(title);
    if (radio.checked) {
      const tag = document.createElement("span");
      tag.className = "chosen-tag";
      tag.textContent = "seçili";
      text.append(" ", tag);
    }
    const chosenHere = state.chosen !== null && personKey(state.chosen) === personKey(person);
    const note =
      state.confirmed && chosenHere
        ? "bu geçiş için onayladınız"
        : contextNote && chosenHere
          ? contextNote
          : matchHint(found.scores.get(personKey(person)) ?? 0);
    const own = isLocal(person) ? "kendi kaydınız" : sourceLabel(person.source) === "Wikidata" ? "Wikidata" : "";
    if (note || own) {
      const hint = document.createElement("span");
      hint.className = "hint";
      hint.textContent = ` · ${[own, note].filter(Boolean).join(" · ")}`;
      text.append(hint);
    }
    if (person.description) {
      const description = document.createElement("small");
      description.textContent = person.description.length > 140 ? `${person.description.slice(0, 140)}…` : person.description;
      text.append(document.createElement("br"), description);
    }
    row.append(radio, text);
    row.dataset.search = normalize(`${candidateLabel(person)} ${person.fullName} ${person.deathDate} ${person.description ?? ""}`);
    rows.push(row);
    list.append(row);
  });
  if (filter) {
    const empty = document.createElement("p");
    empty.className = "muted filter-empty";
    empty.textContent = "Aramaya uyan aday yok.";
    empty.hidden = true;
    list.append(empty);
    rows.push(empty);
    applyFilter(rows, filterText);
  }
  if (hidden > 0 || (expanded && (chosenKey !== null || found.candidates.length > CANDIDATE_PREVIEW))) {
    const more = document.createElement("button");
    more.type = "button";
    more.className = "link more-candidates";
    more.textContent = expanded ? "Adayları kısalt" : chosenKey ? `Diğer adayları göster (${hidden})` : `Daha fazla göster (${hidden})`;
    more.addEventListener("click", () => {
      expanded = !expanded;
      filterText = "";
      renderCandidates();
      if (expanded) document.getElementById("candidate-filter")?.focus();
    });
    list.append(more);
  }  if (expanded || found.candidates.length <= CANDIDATE_PREVIEW) list.append(wikidataSearch(false));
}

/**
 * Wikidata'da ara: the selected name is looked up on Wikidata; the person picked
 * becomes a record of the user's own and is chosen at once.
 */
function wikidataSearch(prominent: boolean): HTMLElement {
  const box = document.createElement("div");
  box.className = "wikidata-search";
  const go = document.createElement("button");
  go.type = "button";
  go.id = "wikidata-search";
  go.className = prominent ? "" : "link";
  go.textContent = prominent ? "Wikidata'da ara" : "Aradığınız kişi yok mu? Wikidata'da ara";
  const results = document.createElement("div");
  results.id = "wikidata-results";
  go.addEventListener("click", () => {
    const alias = state.alias;
    if (!alias) return;
    go.disabled = true;
    results.replaceChildren(muted("Wikidata'da aranıyor…"));
    searchWikidata(nameQuery(alias))
      .then((hits) => showWikidataHits(results, alias, hits))
      .catch(() => results.replaceChildren(muted("Wikidata şu an yanıt vermiyor ya da internete ulaşılamıyor. Biraz sonra yeniden deneyin.")))
      .finally(() => {
        go.disabled = false;
      });
  });
  box.append(go, results);
  return box;
}

function muted(text: string): HTMLElement {
  const p = document.createElement("p");
  p.className = "muted";
  p.textContent = text;
  return p;
}

function showWikidataHits(results: HTMLElement, alias: string, hits: WikidataHit[]): void {
  results.replaceChildren();
  if (hits.length === 0) {
    results.append(muted("Wikidata'da bu adla ölüm tarihi bilinen bir kişi bulunamadı. Kendi şahıslarınıza elle ekleyebilirsiniz."));
    return;
  }
  results.append(muted("Wikidata'da bulunanlar (ölüm tarihi bilinenler). Doğru kişiyi ekleyin:"));
  for (const hit of hits) {
    const row = document.createElement("div");
    row.className = "wikidata-hit";
    const text = document.createElement("span");
    const title = document.createElement("strong");
    const standard = standardLabel(hit);
    title.textContent = `${standard} ${hit.deathDate}`;
    text.append(title);
    const others = [standard !== hit.label ? `Wikidata: ${hit.label}` : "", hit.english].filter(Boolean);
    if (others.length) text.append(` · ${others.join(" · ")}`);
    if (hit.description) {
      const description = document.createElement("small");
      description.textContent = hit.description.length > 140 ? `${hit.description.slice(0, 140)}…` : hit.description;
      text.append(document.createElement("br"), description);
    }
    const add = document.createElement("button");
    add.type = "button";
    add.className = "link";
    add.textContent = "Ekle ve seç";
    add.addEventListener("click", () => void addWikidataHit(alias, hit));
    const link = document.createElement("a");
    link.href = `https://www.wikidata.org/wiki/${hit.id}`;
    link.target = "_blank";
    link.rel = "noopener noreferrer";
    link.textContent = "Wikidata";
    row.append(text, " ", add, " · ", link);
    results.append(row);
  }
}

async function addWikidataHit(alias: string, hit: WikidataHit): Promise<void> {
  const record = localFromHit(alias, hit);
  if (typeof record === "string") {
    setMessage(record, "warning");
    return;
  }
  await addLocals([record]);
  applyLocalPersons();
  state.alias = null;
  await refresh(true);
  setMessage(`${record.shortName} ${hit.deathDate} kendi şahıslarınıza eklendi.`, "success");
}

/** Shows the candidate rows holding every word of the query; the last row is the "nothing found" note. */
function applyFilter(rows: HTMLElement[], query: string): void {
  const words = normalize(query).split(" ").filter(Boolean);
  let shown = 0;
  for (const row of rows) {
    if (row.classList.contains("filter-empty")) {
      row.hidden = shown > 0;
      continue;
    }
    const match = words.every((w) => (row.dataset.search ?? "").includes(w));
    row.hidden = !match;
    if (match) shown += 1;
  }
}

/** Moves the "seçili" mark to the given row without rebuilding the list (keeps keyboard focus). */
function markChosen(list: HTMLElement, row: HTMLElement): void {
  for (const other of list.querySelectorAll<HTMLElement>(".candidate.chosen")) other.classList.remove("chosen");
  list.querySelectorAll(".chosen-tag").forEach((tag) => tag.remove());
  row.classList.add("chosen");
  const tag = document.createElement("span");
  tag.className = "chosen-tag";
  tag.textContent = "seçili";
  row.querySelector("strong")?.after(" ", tag);
}

let pointerChoice = false;
const CANDIDATE_FILTER_FROM = 8;
let filterText = "";
const CANDIDATE_PREVIEW = 5;
let expanded = false;
let expandedFor: string | null = null;

export let spellingOwner = "";

/** The selected name without a title the lookup left out: "Hz. Hatice" → "Hatice". */
function bareAlias(): string {
  const alias = state.alias ?? "";
  const title = state.found?.title ?? "";
  return title && alias.startsWith(title) ? alias.slice(title.length) : alias;
}

export function renderDetails(): void {
  const person = state.chosen;
  $("details").hidden = person === null;
  // A choice among several bearers can be given to every other uncertain mention of this spelling at once.
  $("apply-all").hidden = person === null || !state.alias || !state.found || state.found.automatic !== null || state.found.candidates.length < 2;
  $<HTMLButtonElement>("insert").disabled = person === null || state.busy;
  const spelling = $<HTMLSelectElement>("spelling");
  const picked = spelling.value; // kept across re-renders (busy state, settings)
  spelling.replaceChildren();
  $<HTMLButtonElement>("respell").disabled = true;
  if (!person || !state.alias) return;
  $("detail-full").textContent = textName(person.fullName) || "—";
  $("detail-original").textContent = person.originalName || "—";
  $("detail-date").textContent = person.deathDate || "Katalogda tarih yok";
  const link = $<HTMLAnchorElement>("detail-link");
  // A record of one's own added from Wikidata keeps its item in the description: "… · Wikidata Q469299".
  const ownItem = isLocal(person) ? /Wikidata (Q\d+)/.exec(person.description)?.[1] : undefined;
  const source = ownItem ? `https://www.wikidata.org/wiki/${ownItem}` : person.source;
  const origin = ownItem ? "Wikidata" : sourceLabel(person.source);
  link.hidden = origin === "";
  link.href = link.hidden ? "#" : source;
  link.textContent = origin === "Wikidata" ? "Wikidata kaydını aç" : "DİA maddesini aç";
  $("details-heading").textContent = `Seçili şahıs: ${textName(person.shortName)}`;
  // The current spelling heads the list and is what is selected: nothing changes until another one is picked.
  const current = document.createElement("option");
  current.value = "";
  current.textContent = `${state.alias} (metindeki yazım; değişmez)`;
  current.selected = true;
  spelling.append(current);
  let alternatives = 0;
  const bare = bareAlias();
  for (const option of withPreferred(orderByNameForm(nameOptions(person, bare), state.settings.nameForm), preferredSpelling(readSpellings(), person))) {
    if (option.value === bare || option.value === nameQuery(bare)) continue;
    const element = document.createElement("option");
    element.value = option.value;
    element.textContent = `${option.value} (${option.label})`;
    spelling.append(element);
    alternatives++;
  }
  // Only for the same name and person: a new selection starts from its own spelling.
  const owner = `${state.alias}|${personKey(person)}`;
  if (picked && owner === spellingOwner && [...spelling.options].some((o) => o.value === picked)) spelling.value = picked;
  spellingOwner = owner;
  spelling.disabled = alternatives === 0;
  if (alternatives === 0) current.textContent = `${state.alias} (başka yazım önerisi yok)`;
  renderSpellingChoice();
}

/** "Değiştir" works only once another spelling is picked, and says what it will write. */
export function renderSpellingChoice(): void {
  const spelling = $<HTMLSelectElement>("spelling");
  const chosen = spelling.value;
  const preview = $("spelling-preview");
  $<HTMLButtonElement>("respell").disabled = !chosen || state.busy;
  $<HTMLButtonElement>("respell-all").disabled = !chosen || state.busy;
  preview.hidden = !chosen || !state.alias;
  if (chosen && state.alias) preview.textContent = `Değiştir'e basınca: ${state.alias} → ${state.found?.title ?? ""}${withEnding(bareAlias(), chosen)}`;
}

/** The mentions "Tümünü değiştir" would change: this spelling, this person, anywhere in the document. */
interface RespellTarget {
  readonly segment: number;
  readonly start: number;
  readonly literal: string;
  readonly text: string;
  readonly part: "main" | "note";
  /** In a citation or the bibliography, which the inventory leaves out. */
  readonly cited: boolean;
}

let respellPlan: RespellTarget[] = [];

/**
 * The mentions "Tümünü değiştir" changes: this spelling (with any ending) of
 * this person anywhere in the document, and the selected one. Citations and
 * the bibliography ("Gazali, el-Mustasfâ", "Gazali, Ebû Hâmid.") are not in
 * the inventory; they change too when no other person is written this way in
 * the document, as their bearer cannot be told there.
 */
async function respellTargets(spelling: string): Promise<RespellTarget[]> {
  const person = state.chosen;
  const alias = state.alias ? bareAlias() : null;
  if (!person || !alias) return [];
  const rows = await documentRows();
  const key = nameKey(nameQuery(alias));
  const target = personKey(person);
  const same = rows.filter((r) => nameKey(nameQuery(r.text)) === key);
  const out: RespellTarget[] = same
    .filter((r) => (r.person !== null && personKey(r.person) === target) || (state.mention !== null && rowKey(r, state.segments) === state.mention))
    .map((r) => ({ segment: r.segment, start: r.start, literal: r.text, text: withEnding(r.text, spelling), part: r.part, cited: false }));
  const others = same.some((r) => r.person !== null && personKey(r.person) !== target);
  if (!others && state.index) {
    const known = new Set(rows.map((r) => `${r.segment}:${r.start}`));
    state.segments.forEach((segment, i) => {
      for (const hit of findMatches(segment.text, state.index as NameIndex, false)) {
        if (known.has(`${i}:${hit.start}`) || nameKey(nameQuery(hit.text)) !== key) continue;
        out.push({ segment: i, start: hit.start, literal: hit.text, text: withEnding(hit.text, spelling), part: segment.part, cited: true });
      }
    });
  }
  return out.filter((t) => t.text !== t.literal);
}

/** "Tümünü değiştir": first says how many mentions would change, then waits for Onayla. */
export async function previewRespellAll(): Promise<void> {
  const spelling = $<HTMLSelectElement>("spelling").value;
  if (!spelling || !state.chosen || !state.alias) return;
  respellPlan = await respellTargets(spelling);
  if (respellPlan.length === 0) {
    setMessage("Bu yazımla bu kişiye bağlı başka geçiş bulunamadı; seçili geçiş için Değiştir'i kullanın.", "info");
    return;
  }
  const notes = respellPlan.filter((t) => t.part === "note").length;
  const cited = respellPlan.filter((t) => t.cited).length;
  const parts = [notes ? `${notes} tanesi dipnot/sonnotta` : "", cited ? `${cited} tanesi atıf ya da kaynakçada` : ""].filter(Boolean);
  $("respell-confirm-text").textContent =
    `${bareAlias()} → ${withEnding(bareAlias(), spelling)}: belgede ${respellPlan.length} geçiş değişecek` +
    (parts.length ? ` (${parts.join(", ")})` : "") +
    ". Bu yazımın başka bir kişiye bağlı ya da kimliği kesin olmayan geçişlerine dokunulmaz.";
  $("respell-confirm").hidden = false;
  setMessage("", "info");
}

export async function respellAll(): Promise<void> {
  $("respell-confirm").hidden = true;
  const plan = respellPlan;
  respellPlan = [];
  if (plan.length === 0) return;
  versions.identity++;
  const items = plan.flatMap((t) => {
    const segment = state.segments[t.segment];
    return segment ? [{ segment, start: t.start, literal: t.literal, text: t.text }] : [];
  });
  const done = await replaceMentions(items);
  logHistory("spelling", `${state.alias ?? ""}: ${done} geçiş değiştirildi`);
  setMessage(
    done === plan.length ? `${done} geçişin yazımı değiştirildi.` : `${done} / ${plan.length} geçiş değiştirildi; kalanlar metinde bulunamadı (belge değişmiş olabilir). Dökümü yenileyip tekrar deneyin.`,
    done === plan.length ? "success" : "warning",
  );
  state.alias = null;
  await refresh();
}

export function hideEarlier(): void {
  $("earlier").hidden = true;
  $("respell-confirm").hidden = true;
  $("existing").hidden = true;
  $("dated").hidden = true;
  existingDate = null;
  datedRow = null;
}

/** Another mention of the same person that already carries a date (İlk kullanım: no second date). */
export let datedRow: InventoryRow | null = null;
let datedLater = false;

/**
 * The person's dated mention in the same part (text or notes; in the same
 * chapter for "Her bölümde ilk kullanım"), other than the selected one. A
 * Vefeyât date and a hand-written one count alike.
 */
async function datedMention(person: Person): Promise<InventoryRow | null> {
  const rows = await documentRows();
  const target = personKey(person);
  const part = (await selectionInNote().catch(() => false)) ? "note" : "main";
  const current = state.mention ? rows.find((r) => r.part === part && rowKey(r, state.segments) === state.mention) : undefined;
  let same = rows.filter((r) => r !== current && r.part === part && r.date.trim() !== "" && r.person !== null && personKey(r.person) === target);
  if (state.settings.scope === "chapter" && current) same = same.filter((r) => r.chapter === current.chapter);
  const found = same[0] ?? null;
  datedLater = found !== null && current !== undefined && rows.indexOf(current) < rows.indexOf(found);
  return found;
}

function showDated(row: InventoryRow): void {
  datedRow = row;
  const where = placeOf(state.segments[row.segment]);
  $("dated-text").textContent =
    `Bu şahsın tarihi belgede zaten var: ${row.text} ${row.date.trim()} (${where}). Aynı kişiye ikinci tarih eklenmez.` +
    (datedLater ? " Seçtiğiniz geçiş ondan önce geliyor; tarihi buraya taşıyabilirsiniz." : "");
  $("move-date").hidden = !datedLater;
  $("dated").hidden = false;
  setMessage("", "info");
}

/** "Tarihe git": selects the mention that carries the date. */
export async function gotoDated(): Promise<void> {
  const row = datedRow;
  const segment = row ? state.segments[row.segment] : undefined;
  if (!row || !segment || !(await selectMention(segment, row.start, row.text))) {
    setMessage("Tarihli geçiş metinde bulunamadı; metin değişmiş olabilir.", "warning");
  }
  hideEarlier();
}

/** "Tarihi buraya taşı": the later date goes, the selected (earlier) mention gets it. */
export async function moveDateHere(): Promise<boolean> {
  const person = state.chosen;
  const row = datedRow;
  const segment = row ? state.segments[row.segment] : undefined;
  if (!person || !row || !segment) return false;
  await loadSettings();
  const current = await currentSelection();
  if (!current) return false;
  const plan = planDate(person, state.settings.prefix, current.following, state.settings.dateForm);
  if (!plan.ok) {
    setMessage(plan.reason, "warning");
    return false;
  }
  if (!(await removeDateAfter(segment, row.start, row.text, row.date.trim()))) {
    hideEarlier();
    setMessage("Öteki tarih metinde bulunamadı; hiçbir şey değişmedi. Belge değişmiş olabilir.", "warning");
    return false;
  }
  await rememberChoice();
  const control = await insertDate(plan.text, dateTag(person), `Vefeyât: ${textName(person.shortName)}`);
  logHistory("date", `${current.alias} ${plan.text} (${row.text} geçişinden taşındı)`, control);
  hideEarlier();
  setMessage(`Tarih buraya taşındı: ${plan.text}`, "success");
  state.alias = null;
  return true;
}

/** The date found after the name when Ölüm tarihi ekle was refused, for "Güncelle". */
let existingDate: string | null = null;

/**
 * Ölüm tarihi ekle on a name that already carries a date: when it differs from
 * the catalog (another year, another Kısaltma, an old catalog), offers to
 * replace it; when it is the same, says so.
 */
function offerUpdate(person: Person, following: string): boolean {
  const old = leadingDate(following);
  if (!old || !isValidDate(person.deathDate)) return false;
  const value = formatDate(person.deathDate, state.settings.prefix, state.settings.dateForm);
  if (sameDate(old, value)) {
    setMessage(`Adın arkasındaki tarih zaten katalogdakiyle aynı: ${value}`, "success");
    return true;
  }
  existingDate = old;
  $("existing-text").textContent = `Adın arkasında zaten bir tarih var: ${old}. Katalogdaki tarih: ${value}.`;
  $("existing").hidden = false;
  setMessage("", "info");
  return true;
}

/** "Güncelle": the date after the name is replaced with the catalog's, as a Vefeyât date. */
export async function updateDate(): Promise<boolean> {
  const person = state.chosen;
  const old = existingDate;
  if (!person || !old) return false;
  await loadSettings();
  const current = await currentSelection();
  if (!current) return false;
  if (leadingDate(current.following) !== old) {
    hideEarlier();
    setMessage("Adın arkasındaki tarih değişti. Yeniden Ölüm tarihi ekle'ye basın.", "warning");
    return false;
  }
  const value = formatDate(person.deathDate, state.settings.prefix, state.settings.dateForm);
  await rememberChoice();
  const control = await updateDateAfterSelection(old, value, dateTag(person), `Vefeyât: ${textName(person.shortName)}`);
  hideEarlier();
  if (control === null) {
    setMessage("Tarih adın hemen arkasında bulunamadı; değiştirilmedi.", "warning");
    return false;
  }
  logHistory("date", `${current.alias} ${old} → ${value}`, control);
  setMessage(`Tarih güncellendi: ${old} → ${value}`, "success");
  state.alias = null;
  return true;
}

export async function refresh(force = false): Promise<void> {
  if (state.busy && !force) return;
  hideEarlier();
  try {
    await loadSettings();
    const selection = await readSelection();
    const alias = selectionAlias(selection.text);
    if (alias === state.alias && state.found) {
      renderDetails(); // settings such as Ad biçimi may have changed
      return;
    }
    if (alias !== state.alias) setMessage("");
    state.alias = alias;
    state.found = null;
    state.chosen = null;
    state.confirmed = false;
    contextNote = null;
    state.mention = alias ? mentionKey(alias, selection.preceding, selection.following) : null;
    $("selected-name").textContent = alias ?? "Metinde bir şahıs adı seçin.";
    if (alias && state.catalog) {
      state.found = lookup(state.catalog, alias);
      state.chosen = state.found.automatic;
      if (!state.chosen && state.mention) {
        state.chosen = confirmedPerson(readConfirmations(), state.mention, state.found.candidates);
        state.confirmed = state.chosen !== null;
      }
      if (!state.chosen && state.mention) applyDocumentContext(state.mention);
      if (!state.chosen && state.found.candidates.length > 1) {
        // "Râzî (ö. 313/925)": the date written after the name picks the bearer.
        const dated = state.found.candidates.filter((c) => isValidDate(c.deathDate) && alreadyHasDate(selection.following, c.deathDate));
        if (dated.length === 1) {
          state.chosen = dated[0] as Person;
          contextNote = "metindeki ölüm tarihine göre";
          state.found = { ...state.found, candidates: [dated[0] as Person, ...state.found.candidates.filter((c) => c !== dated[0])] };
        }
      }
    }
    renderCandidates();
    renderDetails();
    renderExclusionToggle();
  } catch (error) {
    console.error(error);
  }
}

/** The selection must still be the name the pane shows. */
export async function currentSelection(): Promise<{ alias: string; following: string } | null> {
  const selection = await readSelection();
  const alias = selectionAlias(selection.text);
  if (!alias || alias !== state.alias) {
    setMessage("Seçim değişti. Adı yeniden seçin.", "warning");
    await refresh();
    return null;
  }
  return { alias, following: selection.following };
}

export async function addDate(force: boolean): Promise<boolean> {
  const person = state.chosen;
  if (!person) return false;
  await loadSettings();
  const rule = scopeRule(state.settings.scope);
  if (!rule.allowed) {
    setMessage(rule.reason ?? "", "warning");
    return false;
  }
  const current = await currentSelection();
  if (!current) return false;
  const plan = planDate(person, state.settings.prefix, current.following, state.settings.dateForm);
  if (!plan.ok) {
    if (offerUpdate(person, current.following)) return false;
    setMessage(plan.reason, "warning");
    return false;
  }
  // One date per person: a mention that already carries one blocks a second (also after "Yine de ekle").
  if (rule.checkEarlier) {
    const dated = await datedMention(person);
    if (dated) {
      $("earlier").hidden = true;
      showDated(dated);
      return false;
    }
  }
  if (!force && rule.checkEarlier && (await earlierMention(person))) {
    $("earlier").hidden = false;
    setMessage("", "info");
    return false;
  }
  await rememberChoice();
  const control = await insertDate(plan.text, dateTag(person), `Vefeyât: ${textName(person.shortName)}`);
  logHistory("date", `${current.alias} ${plan.text}`, control);
  hideEarlier();
  setMessage(`Tarih eklendi: ${plan.text}`, "success");
  state.alias = null; // the caret moved; read the next selection afresh
  return true;
}

export let earlierRow: InventoryRow | null = null;

/**
 * "Daha önce geçti": an earlier mention of the same person, by identity, not
 * by spelling. Fahreddin er-Râzî before does not count for Ebû Bekir er-Râzî;
 * "Gazali" settled as Gazzâlî does. Falls back to spellings when the selected
 * mention is not found in the reading of the document.
 */
export async function earlierMention(person: Person): Promise<boolean> {
  earlierRow = null;
  if (!state.catalog) return state.alias ? findEarlier(earlierSpellings(person, state.alias)) : false;
  const rows = await documentRows();
  const target = personKey(person);
  // The main text and the notes are counted apart: only mentions in the same part count.
  const inNote = await selectionInNote().catch(() => false);
  let same = rows.filter((r) => r.person !== null && personKey(r.person) === target && r.part === (inNote ? "note" : "main"));
  if (state.settings.scope === "chapter") {
    // "Her bölümde ilk kullanım": only this chapter's mentions count.
    const current = state.mention ? rows.find((r) => rowKey(r, state.segments) === state.mention) : undefined;
    if (current) same = same.filter((r) => r.chapter === current.chapter);
  }
  if (inNote) {
    // Notes are other stories: find this mention in the reading, earlier notes come before it.
    const here = state.mention ? rows.findIndex((r) => r.part === "note" && rowKey(r, state.segments) === state.mention) : -1;
    earlierRow = here < 0 ? null : (same.find((r) => r.id < here) ?? null);
    return earlierRow !== null;
  }
  // Word itself tells which of them stands before the selection.
  earlierRow = await firstMentionBefore(state.segments, same);
  return earlierRow !== null;
}

/** Remembers the user's choice for an ambiguous mention in the document. */
export async function rememberChoice(): Promise<void> {
  versions.identity++; // identities may change
  if (!state.chosen || !state.mention || !state.found || state.found.automatic) return;
  await writeConfirmations(addConfirmation(readConfirmations(), state.mention, personKey(state.chosen)));
}

export async function respell(): Promise<void> {
  versions.identity++; // identities may change
  const spelling = $<HTMLSelectElement>("spelling").value;
  if (!spelling) return;
  const selection = await readSelection();
  const name = state.alias ? bareAlias() : null;
  const title = state.found?.title ?? "";
  // The selection may also hold the space a double click takes, a comma or a quote: only the name changes
  // (a title before it, "Hz. ", stays).
  if (!name || selectionAlias(selection.text) !== state.alias) {
    setMessage("Seçim değişti. Adı yeniden seçip Değiştir'e basın.", "warning");
    await refresh();
    return;
  }
  const replacement = withEnding(name, spelling);
  if (state.chosen && state.found && !state.found.automatic && state.mention) {
    // The sentence keeps its words; only the name changes, so key the new spelling.
    const [before, , after] = state.mention.split("|");
    const key = `${before ?? ""}|${title}${replacement}|${after ?? ""}`;
    await writeConfirmations(addConfirmation(readConfirmations(), key, personKey(state.chosen)));
  }
  if (!(await replaceNameInSelection(name, replacement))) {
    setMessage("Ad seçimin içinde bulunamadı; yazım değiştirilmedi. Adı yeniden seçin.", "warning");
    return;
  }
  logHistory("spelling", `${name} → ${replacement}`);
  setMessage(`Yazım değiştirildi: ${title}${replacement}`, "success");
  state.alias = null;
  await refresh();
}

// ---- Belge istisnaları, kimlik eşleşmeleri, kapsam ----

export function renderExclusionToggle(): void {
  const key = state.alias ? exclusionKey(state.alias) : "";
  const excluded = key.length > 0 && readExclusions().includes(key);
  $("excluded-note").hidden = !excluded;
  const button = $<HTMLButtonElement>("not-person");
  button.hidden = !state.alias;
  button.textContent = excluded ? "İşareti kaldır (şahıs adı)" : "Şahıs değil olarak işaretle";
}

/** Şahıs değil / Şahıs say. Returns the key when the phrase was just marked as not a person. */
export async function toggleExclusion(): Promise<string | null> {
  versions.identity++; // identities may change
  if (!state.alias) return null;
  const key = exclusionKey(state.alias);
  const list = readExclusions();
  const excluded = list.includes(key);
  await writeExclusions(excluded ? list.filter((k) => k !== key) : [...list, key]);
  setMessage(excluded ? `"${state.alias}" yeniden şahıs adı olarak taranacak.` : `"${state.alias}" bu belgede şahıs adı sayılmayacak.`, "success");
  renderExclusionToggle();
  return excluded ? null : key;
}

// ---- Tercihler ----

export async function preferSelection(): Promise<void> {
  versions.identity++; // identities may change
  const person = state.chosen;
  if (!state.alias || !person) {
    setMessage(quickInsertBlocker(state.catalog !== null, state.alias, state.found) ?? "Önce şahsı seçin.", "info");
    return;
  }
  const spelling = preferenceOf(state.alias);
  await writeSpellings({ ...readSpellings(), [personKey(person)]: spelling });
  setMessage(`Bu belgede ${textName(person.shortName)} için tercih edilen yazım: ${spelling}. Döküm ve toplu düzeltme bu yazımı esas alır.`, "success");
  renderDetails();
}

/**
 * "Bu yazımın bütün belirsiz geçişlerine uygula": every mention written like
 * the selected one ("Râzî", "Râzî'nin") whose bearer is not certain yet — no
 * identity, or one only guessed from the document or the sentence's subject —
 * is confirmed as the chosen person. Mentions settled otherwise (catalog,
 * written date, an earlier confirmation) keep theirs.
 */
export async function applyToAll(): Promise<void> {
  const person = state.chosen;
  const alias = state.alias ? bareAlias() : null;
  if (!person || !alias) return;
  const rows = await documentRows();
  const key = nameKey(nameQuery(alias));
  const target = personKey(person);
  const open = rows.filter(
    (r) =>
      nameKey(nameQuery(r.text)) === key &&
      r.candidates.some((c) => personKey(c) === target) &&
      (r.person === null || r.resolved === "context" || r.resolved === "topic"),
  );
  let map = readConfirmations();
  for (const row of open) map = addConfirmation(map, rowKey(row, state.segments), target);
  if (state.mention) map = addConfirmation(map, state.mention, target);
  await writeConfirmations(map);
  versions.identity++;
  state.confirmed = true;
  logHistory("bulk", `${alias}: ${open.length} geçiş ${textName(person.shortName)} olarak onaylandı`);
  setMessage(
    open.length
      ? `"${alias}" yazımının ${open.length} belirsiz geçişi ${textName(person.shortName)} olarak onaylandı. Kimlik eşleşmelerinden geri alınabilir.`
      : `"${alias}" yazımının başka belirsiz geçişi yok; bu geçiş ${textName(person.shortName)} olarak onaylandı.`,
    "success",
  );
  renderCandidates();
}
