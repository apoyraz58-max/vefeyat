/**
 * Wikidata'da ara, shared by the selected-name window and Şahıs dökümü: a hit
 * in DİA's spelling, and a hit made a record of the user's own.
 */
import { diaStandard, nameQuery, standardName, wantsLongVowels, wantsStandard } from "../../core/index.js";
import { type LocalPerson, upsertLocal, validateLocal } from "../persons.js";
import { type WikidataHit } from "../wikidata-live.js";
import { readLocalPersons, writeLocalPersons } from "../word.js";
import { state } from "./core.js";

/** The hit's name in DİA's spelling when it is of the Muslim world or an Ottoman subject. */
export function standardLabel(hit: WikidataHit): string {
  const base = state.baseCatalog;
  if (!base || !wantsStandard(hit.countries, hit.deathYear)) return hit.label;
  return standardName(hit.label, diaStandard(base.people), 0.5, hit.arabic, wantsLongVowels(hit.countries));
}

/**
 * The record for a hit; the name as written in the text is kept as a spelling
 * of its own, so the mention finds the record. A reason when it cannot be made.
 */
export function localFromHit(written: string, hit: WikidataHit): LocalPerson | string {
  const label = standardLabel(hit);
  const plain = nameQuery(written);
  const spellings = [hit.label, hit.english, plain].filter((v, i, all) => v && v !== label && all.indexOf(v) === i);
  const checked = validateLocal({
    shortName: label,
    fullName: spellings.join(" | "),
    deathDate: hit.deathDate,
    description: [hit.description, `Wikidata ${hit.id}`].filter(Boolean).join(" · "),
  });
  return checked.ok ? checked.person : checked.reason;
}

/** Adds the records to Kendi şahıslarım (a record with the same short name is replaced). */
export async function addLocals(records: readonly LocalPerson[]): Promise<void> {
  let list = readLocalPersons();
  for (const record of records) list = upsertLocal(list, record);
  await writeLocalPersons(list);
}
