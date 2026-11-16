/** Vefeyât pane — the selected-name window: candidates, dates, spellings, earlier mentions, Şahıs değil. */
import { isLocal } from "../persons.js";
import { fillLocalForm, renderLocals } from "./catalog.js";
import { type Person, alreadyHasDate, isValidDate, nameOptions, nameQuery, personKey, selectionAlias, textName, withEnding } from "../../core/index.js";
import { addConfirmation, confirmedPerson, exclusionKey, mentionKey, candidateLabel, dateTag, earlierSpellings, lookup, matchHint, orderByNameForm, planDate, quickInsertBlocker, scopeRule } from "../logic.js";
import { type InventoryRow } from "../inventory.js";
import { rowKey } from "../bulk.js";
import { preferenceOf, preferredSpelling, withPreferred } from "../prefs.js";
import { readSpellings, writeSpellings, readExclusions, writeExclusions, firstMentionBefore, findEarlier, insertDate, readConfirmations, readSelection, replaceSelection, writeConfirmations } from "../word.js";

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
    return;
  }
  found.candidates.forEach((person, index) => {
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
    radio.addEventListener("change", () => {
      state.chosen = person;
      hideEarlier();
      renderCandidates();
      renderDetails();
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
    const own = isLocal(person) ? "kendi kaydınız" : "";
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
    list.append(row);
  });
}

export let spellingOwner = "";

export function renderDetails(): void {
  const person = state.chosen;
  $("details").hidden = person === null;
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
  link.hidden = !person.source.startsWith("https://islamansiklopedisi.org.tr/");
  link.href = link.hidden ? "#" : person.source;
  $("details-heading").textContent = `Seçili şahıs: ${textName(person.shortName)}`;
  // The current spelling heads the list and is what is selected: nothing changes until another one is picked.
  const current = document.createElement("option");
  current.value = "";
  current.textContent = `${state.alias} (metindeki yazım; değişmez)`;
  current.selected = true;
  spelling.append(current);
  let alternatives = 0;
  for (const option of withPreferred(orderByNameForm(nameOptions(person, state.alias), state.settings.nameForm), preferredSpelling(readSpellings(), person))) {
    if (option.value === state.alias || option.value === nameQuery(state.alias)) continue;
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
  preview.hidden = !chosen || !state.alias;
  if (chosen && state.alias) preview.textContent = `Değiştir'e basınca: ${state.alias} → ${withEnding(state.alias, chosen)}`;
}

export function hideEarlier(): void {
  $("earlier").hidden = true;
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

export async function addDate(force: boolean): Promise<void> {
  const person = state.chosen;
  if (!person) return;
  await loadSettings();
  const rule = scopeRule(state.settings.scope);
  if (!rule.allowed) {
    setMessage(rule.reason ?? "", "warning");
    return;
  }
  const current = await currentSelection();
  if (!current) return;
  const plan = planDate(person, state.settings.prefix, current.following);
  if (!plan.ok) {
    setMessage(plan.reason, "warning");
    return;
  }
  if (!force && rule.checkEarlier && (await earlierMention(person))) {
    $("earlier").hidden = false;
    setMessage("", "info");
    return;
  }
  await rememberChoice();
  const control = await insertDate(plan.text, dateTag(person), `Vefeyât: ${textName(person.shortName)}`);
  logHistory("date", `${current.alias} ${plan.text}`, control);
  hideEarlier();
  setMessage(`Tarih eklendi: ${plan.text}`, "success");
  state.alias = null; // the caret moved; read the next selection afresh
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
  const same = rows.filter((r) => r.person !== null && personKey(r.person) === target);
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
  if (selection.text !== state.alias) {
    setMessage("Yalnız adı seçin (boşluk veya noktalama olmadan); yazım yalnız seçili metinde değişir.", "warning");
    return;
  }
  const replacement = withEnding(selection.text, spelling);
  if (state.chosen && state.found && !state.found.automatic && state.mention) {
    // The sentence keeps its words; only the name changes, so key the new spelling.
    const [before, , after] = state.mention.split("|");
    const key = `${before ?? ""}|${replacement}|${after ?? ""}`;
    await writeConfirmations(addConfirmation(readConfirmations(), key, personKey(state.chosen)));
  }
  await replaceSelection(replacement);
  logHistory("spelling", `${selection.text} → ${replacement}`);
  setMessage(`Yazım değiştirildi: ${replacement}`, "success");
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

export async function toggleExclusion(): Promise<void> {
  versions.identity++; // identities may change
  if (!state.alias) return;
  const key = exclusionKey(state.alias);
  const list = readExclusions();
  const excluded = list.includes(key);
  await writeExclusions(excluded ? list.filter((k) => k !== key) : [...list, key]);
  setMessage(excluded ? `"${state.alias}" yeniden şahıs adı olarak taranacak.` : `"${state.alias}" bu belgede şahıs adı sayılmayacak.`, "success");
  renderExclusionToggle();
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
