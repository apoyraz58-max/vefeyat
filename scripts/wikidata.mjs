// Builds field catalogs from Wikidata (CC0) next to the DİA catalog:
//   node scripts/wikidata.mjs            → data/wikidata/<alan>.csv for every field
//   node scripts/wikidata.mjs sosyoloji  → one field
// Each field is a list of queries (an occupation, a citizenship …) with a
// notability threshold (number of Wikipedia articles). People DİA has are left
// out: Wikidata's TDV İslâm Ansiklopedisi ID (P7314), or the same name and
// death year as a DİA entry. DİA is authoritative where the two meet.
import { mkdir, readFile, writeFile } from "node:fs/promises";

const ENDPOINT = "https://query.wikidata.org/sparql";
const AGENT = "VefeyatCatalogBuilder/0.1 (Word add-in catalog generator; CC0 data)";

// Occupations (P106) and citizenships (P27) by Wikidata item.
const Q = {
  sociologist: "Q2306091", philosopher: "Q4964182", anthropologist: "Q4773904", psychologist: "Q212980",
  historian: "Q201788", politician: "Q82955", militaryOfficer: "Q189290", diplomat: "Q193391", universityTeacher: "Q1622272",
  monarch: "Q116", writer: "Q36180", poet: "Q49757", novelist: "Q6625963", journalist: "Q1930187",
  economist: "Q188094", politicalScientist: "Q1238570", jurist: "Q185351", statesman: "Q372436",
  playwright: "Q214917", essayist: "Q11774202", literaryCritic: "Q4263842",
  ottoman: "Q12560", turkey: "Q43",
};

/**
 * Countries and states whose politicians a social-science text names most:
 * Germany, France, United Kingdom (and Great Britain), USA, Russia (Empire, USSR),
 * Italy, Spain, Austria (Austria-Hungary), Prussia, Netherlands, Greece, Poland,
 * Hungary, Iran, Egypt, India, China, Japan, Israel, Iraq, Syria, Saudi Arabia,
 * Pakistan, Afghanistan, Bulgaria, Serbia, Romania, Roman and Byzantine Empires.
 */
const POLITICAL_COUNTRIES = [
  "Q183", "Q142", "Q145", "Q174193", "Q30", "Q159", "Q34266", "Q15180", "Q38", "Q29", "Q40", "Q28513", "Q27306", "Q55", "Q41", "Q36", "Q28",
  "Q794", "Q79", "Q668", "Q148", "Q17", "Q801", "Q796", "Q858", "Q851", "Q843", "Q889", "Q219", "Q403", "Q218", "Q2277", "Q12544",
];

/** One query: people with the occupation (or citizenship) and at least `min` sitelinks. */
const occupation = (item, min) => ({ where: `?p wdt:P106 wd:${item} .`, min });
const citizen = (country, min, occupations = []) => ({
  where: `?p wdt:P27 wd:${country} .` + (occupations.length ? ` ?p wdt:P106 ?occ . VALUES ?occ { ${occupations.map((o) => `wd:${o}`).join(" ")} }` : ""),
  min,
});

export const FIELDS = {
  "sosyoloji-felsefe": {
    title: "Sosyoloji ve felsefe",
    queries: [occupation(Q.sociologist, 4), occupation(Q.philosopher, 12), occupation(Q.anthropologist, 8), occupation(Q.psychologist, 15),
      citizen(Q.turkey, 2, [Q.sociologist, Q.philosopher, Q.psychologist, Q.anthropologist])],
  },
  "osmanli-turk-tarihi": {
    title: "Osmanlı-Türk tarihi",
    queries: [citizen(Q.ottoman, 3), citizen(Q.turkey, 3, [Q.politician, Q.historian, Q.militaryOfficer, Q.diplomat, Q.universityTeacher]), occupation(Q.historian, 15)],
  },
  edebiyat: {
    title: "Edebiyat",
    // "writer" holds hundreds of thousands of people, past Wikidata's time limit: its narrower occupations instead.
    queries: [occupation(Q.novelist, 30), occupation(Q.poet, 30), occupation(Q.playwright, 30), occupation(Q.essayist, 25), occupation(Q.literaryCritic, 20),
      citizen(Q.turkey, 2, [Q.writer, Q.poet, Q.novelist, Q.journalist]), citizen(Q.ottoman, 2, [Q.writer, Q.poet])],
  },
  "iktisat-siyaset": {
    title: "İktisat ve siyaset",
    // "politician" is too large for one query: country by country (Turkish ones are in Osmanlı-Türk tarihi).
    queries: [occupation(Q.economist, 8), occupation(Q.politicalScientist, 8), occupation(Q.statesman, 20), occupation(Q.monarch, 40),
      citizen(Q.turkey, 2, [Q.economist, Q.politicalScientist, Q.jurist]),
      ...POLITICAL_COUNTRIES.map((country) => citizen(country, 30, [Q.politician]))],
  },
};

function sparql({ where, min }) {
  return `SELECT ?p ?links ?death ?dprec ?birth ?bprec ?dia ?trLabel ?enLabel ?trDesc ?enDesc ?alt WHERE {
  ${where}
  ?p wdt:P31 wd:Q5 ; wikibase:sitelinks ?links .
  FILTER(?links >= ${min})
  ?p p:P570/psv:P570 [ wikibase:timeValue ?death ; wikibase:timePrecision ?dprec ] .
  OPTIONAL { ?p p:P569/psv:P569 [ wikibase:timeValue ?birth ; wikibase:timePrecision ?bprec ] }
  OPTIONAL { ?p wdt:P7314 ?dia }
  OPTIONAL { ?p rdfs:label ?trLabel FILTER(LANG(?trLabel) = "tr") }
  OPTIONAL { ?p rdfs:label ?enLabel FILTER(LANG(?enLabel) = "en") }
  OPTIONAL { ?p schema:description ?trDesc FILTER(LANG(?trDesc) = "tr") }
  OPTIONAL { ?p schema:description ?enDesc FILTER(LANG(?enDesc) = "en") }
  OPTIONAL { ?p skos:altLabel ?alt FILTER(LANG(?alt) = "tr") }
}`;
}

async function run(query, attempt = 1) {
  const response = await fetch(`${ENDPOINT}?${new URLSearchParams({ query: query.raw ?? sparql(query), format: "json" })}`, {
    headers: { Accept: "application/sparql-results+json", "User-Agent": AGENT },
  });
  if (response.status === 429 || response.status >= 500) {
    if (attempt > 3) throw new Error(`Wikidata ${response.status}`);
    await new Promise((r) => setTimeout(r, 15000 * attempt));
    return run(query, attempt + 1);
  }
  if (!response.ok) throw new Error(`Wikidata ${response.status}: ${(await response.text()).slice(0, 200)}`);
  const text = await response.text();
  try {
    return JSON.parse(text).results.bindings;
  } catch {
    // A query that runs past Wikidata's time limit comes back cut off: try again, then give up on it.
    if (attempt > 2) throw new Error(`Wikidata yanıtı yarım kaldı (sorgu çok büyük): ${query.where}`);
    await new Promise((r) => setTimeout(r, 5000));
    return run(query, attempt + 1);
  }
}

/** "1917-11-15T00:00:00Z" → 1917; "-0322-…" → -322. */
const year = (value) => {
  const m = /^(-?)(\d+)-/.exec(value ?? "");
  return m ? Number(m[2]) * (m[1] ? -1 : 1) : null;
};

/** "(1858-1917)", "(ö. 1917)", "(m.ö. 384-322)", "(ö. 6. yy)"; "" when the date says too little. */
function lifespan(birth, death, precision) {
  if (death === null || precision < 7) return "";
  if (precision < 9) {
    // A decade or a century only: the century, as an approximate date.
    const century = Math.ceil(Math.abs(death) / 100) || 1;
    return death < 0 ? `(m.ö. ${century}. yy)` : `(ö. ${century}. yy)`;
  }
  if (death < 0) {
    if (birth !== null && birth < 0 && -birth > -death && -birth - -death < 120) return `(m.ö. ${-birth}-${-death})`;
    return `(m.ö. ö. ${-death})`;
  }
  if (birth !== null && birth < 0) return death - birth < 120 ? `(m.ö. ${-birth} - m.s. ${death})` : `(ö. ${death})`;
  return birth !== null && birth < death && death - birth < 120 ? `(${birth}-${death})` : `(ö. ${death})`;
}

const fold = (s) =>
  s.toLocaleLowerCase("tr").normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/ı/g, "i").replace(/[^a-z0-9]+/g, "");

/** DİA short names and death years, to leave out the people DİA already has. */
async function diaKeys() {
  const csv = await readFile(new URL("../data/catalog.csv", import.meta.url), "utf8");
  const keys = new Set();
  for (const line of csv.split(/\r?\n/).slice(1)) {
    const [short, , , date] = line.split(";");
    if (!short) continue;
    const years = (date ?? "").match(/\d{3,4}/g) ?? [];
    const last = years[years.length - 1];
    const natural = short.includes(",") ? `${short.split(",")[1]} ${short.split(",")[0]}` : short;
    for (const name of [short, natural]) if (last) keys.add(`${fold(name)}|${last}`);
  }
  return keys;
}

// No ; or line breaks inside a field; a field with a quote is quoted (CSV).
const cell = (s) => {
  const value = (s ?? "").replace(/[;\r\n]+/g, ",").trim();
  return value.includes('"') ? `"${value.replace(/"/g, '""')}"` : value;
};

/**
 * Turkish occupation and country labels of the given people, in batches:
 * the material of a Turkish description where Wikidata has none
 * ("filozof, sosyolog (Almanya)").
 */
async function turkishFacts(ids) {
  const facts = new Map();
  for (let k = 0; k < ids.length; k += 250) {
    const batch = ids.slice(k, k + 250);
    const query = {
      raw: `SELECT ?p ?occ ?country WHERE {
  VALUES ?p { ${batch.map((id) => `wd:${id}`).join(" ")} }
  OPTIONAL { ?p wdt:P106 ?o . ?o rdfs:label ?occ FILTER(LANG(?occ) = "tr") }
  OPTIONAL { ?p wdt:P27 ?c . ?c rdfs:label ?country FILTER(LANG(?country) = "tr") }
}`,
    };
    let rows = [];
    try {
      rows = await run(query);
    } catch (error) {
      console.log(`  Türkçe tanıtım bilgisi alınamadı (${k}-${k + batch.length}): ${error.message}`);
      continue;
    }
    for (const r of rows) {
      const id = r.p.value.split("/").pop();
      const fact = facts.get(id) ?? { occupations: [], countries: [] };
      if (r.occ?.value && !fact.occupations.includes(r.occ.value)) fact.occupations.push(r.occ.value);
      if (r.country?.value && !fact.countries.includes(r.country.value)) fact.countries.push(r.country.value);
      facts.set(id, fact);
    }
    await new Promise((r) => setTimeout(r, 1000));
  }
  return facts;
}

/** "filozof, sosyolog (Almanya)", or "" when Wikidata has no Turkish labels for either. */
function turkishDescription(fact) {
  if (!fact || fact.occupations.length === 0) return "";
  const what = fact.occupations.slice(0, 3).join(", ");
  return fact.countries.length ? `${what} (${fact.countries[0]})` : what;
}

async function buildField(name, field, dia) {
  const people = new Map();
  const skipped = [];
  for (const query of field.queries) {
    const started = Date.now();
    let rows;
    try {
      rows = await run(query);
    } catch (error) {
      // One query past Wikidata's limits does not sink the field: it is reported and left out.
      console.log(`  ATLANDI: ${query.where.slice(0, 70)} (en az ${query.min} madde): ${error.message}`);
      skipped.push(query.where.slice(0, 70));
      continue;
    }
    console.log(`  ${query.where.slice(0, 70)} (en az ${query.min} madde): ${rows.length} satır, ${Math.round((Date.now() - started) / 1000)} sn`);
    for (const r of rows) {
      const id = r.p.value.split("/").pop();
      const person = people.get(id) ?? { id, links: 0, alts: new Set(), dprec: -1 };
      person.links = Math.max(person.links, Number(r.links?.value ?? 0));
      person.dia ||= Boolean(r.dia);
      person.tr ??= r.trLabel?.value;
      person.en ??= r.enLabel?.value;
      person.trDesc ??= r.trDesc?.value;
      person.enDesc ??= r.enDesc?.value;
      // Precision 9 = year, 8 = decade, 7 = century: only a known year is written as one.
      const d = year(r.death?.value);
      const b = year(r.birth?.value);
      const dp = Number(r.dprec?.value ?? 0);
      const bp = Number(r.bprec?.value ?? 0);
      if (d !== null && (person.death === undefined || dp > person.dprec || (dp === person.dprec && d < person.death))) {
        person.death = d;
        person.dprec = dp;
      }
      if (b !== null && bp >= 9 && (person.birth === undefined || b < person.birth)) person.birth = b;
      if (r.alt?.value) person.alts.add(r.alt.value);
      people.set(id, person);
    }
    await new Promise((r) => setTimeout(r, 1500));
  }
  const out = ["Kısa Ad;Tam Ad;Orijinal Ad;Vefat Tarihi;Tanıtım;Kaynak Link"];
  let skippedDia = 0;
  const missing = [...people.values()].filter((p) => !p.trDesc).map((p) => p.id);
  const facts = missing.length ? await turkishFacts(missing) : new Map();
  const list = [...people.values()].filter((p) => (p.tr || p.en) && p.death !== undefined && lifespan(p.birth ?? null, p.death, p.dprec) && !/^Q\d+$/.test(p.tr ?? p.en));
  list.sort((a, b) => b.links - a.links);
  for (const p of list) {
    const short = p.tr ?? p.en;
    if (p.dia || dia.has(`${fold(short)}|${p.death}`) || (p.en && dia.has(`${fold(p.en)}|${p.death}`))) {
      skippedDia++;
      continue;
    }
    const original = p.en && p.en !== short ? p.en : "---";
    // Turkish aliases ("Aristoteles" for "Aristo") go into the full-name column, which the add-in also searches.
    const alts = [...p.alts].filter((a) => a !== short && a !== original && a.length <= 60 && !a.includes(",") && !/\.$/.test(a) && (a.match(/\p{L}/gu) ?? []).length >= 3).slice(0, 3);
    // Other Turkish spellings, separated by " | ": the add-in takes each as a name of its own.
    const full = alts.length ? alts.join(" | ") : "---";
    const description = p.trDesc ?? (turkishDescription(facts.get(p.id)) || p.enDesc || "");
    out.push([short, full, original, lifespan(p.birth ?? null, p.death, p.dprec ?? 0), description, `https://www.wikidata.org/wiki/${p.id}`].map(cell).join(";"));
  }
  await mkdir(new URL("../data/wikidata/", import.meta.url), { recursive: true });
  await writeFile(new URL(`../data/wikidata/${name}.csv`, import.meta.url), "﻿" + out.join("\r\n") + "\r\n");
  console.log(`${name}: ${out.length - 1} kişi (${skippedDia} kişi DİA'da olduğu için alınmadı)${skipped.length ? `; atlanan sorgu: ${skipped.length}` : ""}`);
  // data/wikidata/index.json: the fields the add-in offers (title, file, count).
  const indexUrl = new URL("../data/wikidata/index.json", import.meta.url);
  let index = [];
  try {
    index = JSON.parse(await readFile(indexUrl, "utf8"));
  } catch {
    index = [];
  }
  index = index.filter((e) => e.id !== name);
  index.push({ id: name, title: field.title, file: `catalogs/${name}.csv`, count: out.length - 1, source: "Wikidata", built: new Date().toISOString().slice(0, 10) });
  index.sort((a, b) => Object.keys(FIELDS).indexOf(a.id) - Object.keys(FIELDS).indexOf(b.id));
  await writeFile(indexUrl, JSON.stringify(index, null, 2) + "\n");
}

// node scripts/wikidata.mjs --index: rewrite data/wikidata/index.json from the CSV files present.
if (process.argv[2] === "--index") {
  const index = [];
  for (const [name, field] of Object.entries(FIELDS)) {
    let text;
    try {
      text = await readFile(new URL(`../data/wikidata/${name}.csv`, import.meta.url), "utf8");
    } catch {
      continue;
    }
    const count = text.split(/\r?\n/).filter((l) => l.trim()).length - 1;
    index.push({ id: name, title: field.title, file: `catalogs/${name}.csv`, count, source: "Wikidata" });
  }
  await writeFile(new URL("../data/wikidata/index.json", import.meta.url), JSON.stringify(index, null, 2) + "\n");
  console.log(index.map((e) => `${e.id}: ${e.count}`).join(", "));
} else {
  const wanted = process.argv[2];
  const dia = await diaKeys();
  for (const [name, field] of Object.entries(FIELDS)) {
    if (wanted && wanted !== name) continue;
    await buildField(name, field, dia);
  }
}
