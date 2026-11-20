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
  theologian: "Q1234713", islamicJurist: "Q1999841", imam: "Q125482", mufti: "Q185992", ayatollah: "Q184402", marja: "Q19641",
  cleric: "Q2259532", religiousLeader: "Q15995642", lawyer: "Q40348", activist: "Q15253558", politicalActivist: "Q11499147",
  revolutionary: "Q3242115", physician: "Q39631", physicist: "Q169470", chemist: "Q593644", mathematician: "Q170790",
  biologist: "Q864503", astronomer: "Q11063", computerScientist: "Q82594", engineer: "Q81096", geographer: "Q901402",
  linguist: "Q14467526", orientalist: "Q1731155", educator: "Q974144", scientist: "Q901",
  ottoman: "Q12560", turkey: "Q43",
};

/**
 * The Muslim world outside Turkey (which has fields of its own): the member
 * states of the Organisation of Islamic Cooperation and a few states before
 * them (Mandatory Palestine, British Raj), plus India for its Muslim scholars.
 */
const MUSLIM_COUNTRIES = [
  "Q889", "Q222", "Q262", "Q227", "Q398", "Q902", "Q225", "Q921", "Q965", "Q657", "Q970", "Q977", "Q79", "Q1005", "Q1006",
  "Q252", "Q794", "Q796", "Q810", "Q232", "Q1246", "Q817", "Q813", "Q822", "Q1016", "Q833", "Q826", "Q912", "Q1025", "Q1028",
  "Q1032", "Q1033", "Q842", "Q843", "Q219060", "Q846", "Q851", "Q1041", "Q1044", "Q1045", "Q1049", "Q858", "Q863", "Q948",
  "Q874", "Q878", "Q265", "Q805", "Q193714", "Q23681", "Q129286", "Q668",
];

/** Thinkers, scholars and men of religion of a country of the Muslim world. */
const MUSLIM_LEARNED = [Q.theologian, Q.islamicJurist, Q.imam, Q.mufti, Q.ayatollah, Q.marja, Q.cleric, Q.religiousLeader,
  Q.writer, Q.poet, Q.journalist, Q.historian, Q.philosopher, Q.sociologist, Q.economist, Q.jurist, Q.universityTeacher,
  Q.linguist, Q.educator, Q.physician, Q.scientist];
/** Its men of state and of public life. */
const MUSLIM_PUBLIC = [Q.politician, Q.militaryOfficer, Q.diplomat, Q.revolutionary, Q.activist, Q.politicalActivist, Q.lawyer, Q.monarch];

/**
 * Countries and states whose politicians a social-science text names most:
 * Germany, France, United Kingdom (and Great Britain), Russia (Empire, USSR),
 * Italy, Spain, Austria (Austria-Hungary), Prussia, Netherlands, Greece, Poland,
 * Hungary, Iran, Egypt, India, China, Japan, Israel, Iraq, Syria, Saudi Arabia,
 * Pakistan, Afghanistan, Bulgaria, Serbia, Romania, Roman and Byzantine Empires.
 */
const POLITICAL_COUNTRIES = [
  "Q183", "Q142", "Q145", "Q174193", "Q159", "Q34266", "Q15180", "Q38", "Q29", "Q40", "Q28513", "Q27306", "Q55", "Q41", "Q36", "Q28",
  "Q794", "Q79", "Q668", "Q148", "Q17", "Q801", "Q796", "Q858", "Q851", "Q843", "Q889", "Q219", "Q403", "Q218", "Q2277", "Q12544",
];

/** People who held one of the offices (P39). */
const office = (items, min) => ({ where: `?p wdt:P39 ?office . VALUES ?office { ${items.map((o) => `wd:${o}`).join(" ")} }`, min });

/** One query: people with the occupation (or citizenship) and at least `min` sitelinks. */
const occupation = (item, min) => ({ where: `?p wdt:P106 wd:${item} .`, min });
const citizen = (country, min, occupations = []) => ({
  country,
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
      ...POLITICAL_COUNTRIES.map((country) => citizen(country, 30, [Q.politician])),
      // Holders of the great offices, whatever their occupation is recorded as: US presidents, vice presidents and
      // secretaries of state, German chancellors, French presidents, UN secretaries-general, popes.
      office(["Q11696", "Q11699", "Q14213", "Q4970706", "Q191954", "Q81066", "Q19546"], 0),
      // Every other sovereign state, with a higher bar: a politician known in many languages is not left out.
      { worldPoliticians: 40 }],
  },
  "islam-dunyasi": {
    title: "Çağdaş İslâm dünyası",
    // Country by country: scholars and thinkers from five Wikipedia articles, public figures from eight.
    queries: [
      ...MUSLIM_COUNTRIES.map((country) => citizen(country, 5, MUSLIM_LEARNED)),
      ...MUSLIM_COUNTRIES.filter((c) => c !== "Q668" && c !== "Q129286").map((country) => citizen(country, 8, MUSLIM_PUBLIC)),
      // Scholars whose citizenship Wikidata does not record.
      ...[occupation(Q.islamicJurist, 3), occupation(Q.mufti, 3), occupation(Q.ayatollah, 2), occupation(Q.marja, 2)].map((q) => ({ ...q, standard: true }))],
  },
  bilim: {
    title: "Bilim ve tıp",
    queries: [occupation(Q.physicist, 25), occupation(Q.chemist, 25), occupation(Q.mathematician, 25), occupation(Q.biologist, 25),
      occupation(Q.astronomer, 20), occupation(Q.physician, 35), occupation(Q.computerScientist, 20), occupation(Q.engineer, 40),
      occupation(Q.geographer, 15), occupation(Q.linguist, 15), occupation(Q.orientalist, 6),
      citizen(Q.turkey, 2, [Q.physician, Q.physicist, Q.chemist, Q.mathematician, Q.biologist, Q.engineer, Q.scientist, Q.linguist])],
  },
};

/** Sovereign states not yet asked about one by one (the political countries, Turkey and the USA are). */
async function worldCountries() {
  const rows = await run({ raw: "SELECT DISTINCT ?c WHERE { ?c wdt:P31 wd:Q3624078 . }" });
  const asked = new Set([...POLITICAL_COUNTRIES, Q.turkey, "Q30"]);
  return rows.map((r) => r.c.value.split("/").pop()).filter((c) => !asked.has(c));
}

/** The queries of a field, with placeholders (every other country's politicians) spelled out. */
async function fieldQueries(field) {
  const out = [];
  for (const query of field.queries) {
    if (query.worldPoliticians) out.push(...(await worldCountries()).map((c) => citizen(c, query.worldPoliticians, [Q.politician])));
    else out.push(query);
  }
  return out;
}

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
/** Labels in the main Latin-alphabet languages and citizenships of the people, by id. */
async function westernLabels(ids) {
  const out = new Map();
  for (let k = 0; k < ids.length; k += 200) {
    const batch = ids.slice(k, k + 200);
    let rows = [];
    try {
      rows = await cached({
        raw: `SELECT ?p ?l ?c WHERE { VALUES ?p { ${batch.map((id) => `wd:${id}`).join(" ")} } OPTIONAL { ?p rdfs:label ?l FILTER(LANG(?l) IN ("de", "en", "fr", "it", "es", "nl", "sv", "pt", "pl", "da", "nb", "cs", "hu")) } OPTIONAL { ?p wdt:P27 ?c } }`,
      });
    } catch (error) {
      console.log(`  Batı dillerindeki adlar alınamadı (${batch.length} kişi): ${error.message}`);
    }
    // Each label once per language: the rows repeat it for every citizenship.
    const seen = new Map();
    for (const r of rows) {
      const id = r.p.value.split("/").pop();
      const facts = out.get(id) ?? { labels: [], countries: [] };
      if (r.l?.value) {
        const lang = r.l["xml:lang"] ?? "";
        const key = `${id}|${lang}`;
        if (!seen.has(key)) {
          seen.set(key, true);
          facts.labels.push(r.l.value);
        }
      }
      const country = r.c?.value?.split("/").pop();
      if (country && !facts.countries.includes(country)) facts.countries.push(country);
      out.set(id, facts);
    }
  }
  return out;
}

/** Arabic labels (Persian when there is none) of the people, by id. */
async function arabicLabels(ids) {
  const out = new Map();
  for (let k = 0; k < ids.length; k += 300) {
    const batch = ids.slice(k, k + 300);
    let rows = [];
    try {
      rows = await cached({ raw: `SELECT ?p ?ar ?fa WHERE { VALUES ?p { ${batch.map((id) => `wd:${id}`).join(" ")} } OPTIONAL { ?p rdfs:label ?ar FILTER(LANG(?ar) = "ar") } OPTIONAL { ?p rdfs:label ?fa FILTER(LANG(?fa) = "fa") } }` });
    } catch (error) {
      console.log(`  Arapça adlar alınamadı (${batch.length} kişi): ${error.message}`);
    }
    for (const r of rows) {
      const name = r.ar?.value ?? r.fa?.value;
      if (name) out.set(r.p.value.split("/").pop(), name);
    }
  }
  return out;
}

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
      rows = await cached(query);
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

/**
 * A query's answer, kept for a day in data/wikidata/.cache/ (not committed): a run
 * that is cut off picks up where it stopped instead of asking Wikidata again.
 */
async function cached(query) {
  const { createHash } = await import("node:crypto");
  const dir = new URL("../data/wikidata/.cache/", import.meta.url);
  const file = new URL(`${createHash("sha1").update(query.raw ?? sparql(query)).digest("hex")}.json`, dir);
  try {
    const { stat } = await import("node:fs/promises");
    if (Date.now() - (await stat(file)).mtimeMs < 86400000) return JSON.parse(await readFile(file, "utf8"));
  } catch {
    // not cached yet
  }
  const rows = await run(query);
  await mkdir(dir, { recursive: true });
  await writeFile(file, JSON.stringify(rows));
  return rows;
}

/**
 * DİA's spelling (src/core/standard.ts, the add-in's own code, bundled for Node):
 * `wants` says whose name takes it, `name` writes a name that way.
 */
let standardTools = null;
async function spellingStandard() {
  if (standardTools) return standardTools;
  const { build } = await import("esbuild");
  const outfile = new URL("../node_modules/.cache/vefeyat-standard.mjs", import.meta.url);
  await build({ entryPoints: [new URL("../src/core/index.ts", import.meta.url).pathname], bundle: true, format: "esm", platform: "node", outfile: outfile.pathname, logLevel: "error" });
  const core = await import(outfile.href);
  const dia = core.parseCatalogCsv(await readFile(new URL("../data/catalog.csv", import.meta.url), "utf8"));
  const table = core.buildSpellingStandard(dia);
  standardTools = {
    wants: (countries, death) => core.wantsStandard(countries, death ?? null),
    name: (n, arabic, countries = []) => core.standardName(n, table, 0.5, arabic ?? "", core.wantsLongVowels(countries)),
    western: (n, labels, countries) => core.westernSpelling(n, labels, countries),
    near: (a, b) => core.nearWord(core.spellingKey(a), core.spellingKey(b), 2),
    badCase: (n) => core.badCase(n),
  };
  return standardTools;
}

async function buildField(name, field, dia) {
  const people = new Map();
  const skipped = [];
  for (const query of await fieldQueries(field)) {
    const started = Date.now();
    let rows;
    try {
      rows = await cached(query);
    } catch (error) {
      // One query past Wikidata's limits does not sink the field: it is reported and left out.
      console.log(`  ATLANDI: ${query.where.slice(0, 70)} (en az ${query.min} madde): ${error.message}`);
      skipped.push(query.where.slice(0, 70));
      continue;
    }
    console.log(`  ${query.where.slice(0, 70)} (en az ${query.min} madde): ${rows.length} satır, ${Math.round((Date.now() - started) / 1000)} sn`);
    for (const r of rows) {
      const id = r.p.value.split("/").pop();
      const person = people.get(id) ?? { id, links: 0, alts: new Set(), dprec: -1, countries: new Set(), standard: false };
      if (query.country) person.countries.add(query.country);
      person.standard ||= Boolean(query.standard);
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
  const standard = await spellingStandard();
  // The Arabic (or Persian) names of those whose names take DİA's spelling: they tell apart Hâmid and Hamîd.
  const arabic = await arabicLabels(list.filter((p) => p.standard || standard.wants([...p.countries], p.death)).map((p) => p.id));
  // Western names whose Turkish label is a letter or two from the English one: a slip, or a Turkish rendering.
  const western = await westernLabels(
    list.filter((p) => !(p.standard || standard.wants([...p.countries], p.death)) && p.tr && p.en && p.tr !== p.en && standard.near(p.tr, p.en)).map((p) => p.id),
  );
  let mended = 0;
  let standardized = 0;
  for (const p of list) {
    const labels = western.get(p.id);
    const fixed = labels && p.tr ? standard.western(p.tr, labels.labels, labels.countries) : null;
    if (fixed) {
      mended++;
      p.alts.add(p.tr);
    }
    const written = fixed ?? p.tr ?? p.en;
    // Names of the Muslim world and of Ottoman subjects in DİA's spelling; Wikidata's is kept among the names.
    const short = p.standard || standard.wants([...p.countries], p.death) ? standard.name(written, arabic.get(p.id), [...p.countries]) : written;
    if (p.dia || [written, short, p.en].some((n) => n && dia.has(`${fold(n)}|${p.death}`))) {
      skippedDia++;
      continue;
    }
    if (short !== written) {
      standardized++;
      p.alts = new Set([written, ...p.alts]);
    }
    const english = p.en && p.en !== short ? p.en : "---";
    const original = arabic.has(p.id) ? arabic.get(p.id) : english;
    if (original !== english && english !== "---") p.alts = new Set([...p.alts].slice(0, 1).concat(english, [...p.alts].slice(1)));
    // Turkish aliases ("Aristoteles" for "Aristo") go into the full-name column, which the add-in also searches.
    // An alias with wrong capitals ("Friedrich daniel ernst schleiermacher") is a careless one: left out.
    const alts = [...p.alts].filter((a) => !standard.badCase(a) && a !== short && a !== original && a.length <= 60 && !a.includes(",") && !/\.$/.test(a) && (a.match(/\p{L}/gu) ?? []).length >= 3).slice(0, 3);
    // Other Turkish spellings, separated by " | ": the add-in takes each as a name of its own.
    const full = alts.length ? alts.join(" | ") : "---";
    const description = p.trDesc ?? (turkishDescription(facts.get(p.id)) || p.enDesc || "");
    out.push([short, full, original, lifespan(p.birth ?? null, p.death, p.dprec ?? 0), description, `https://www.wikidata.org/wiki/${p.id}#${p.links}`].map(cell).join(";"));
  }
  await mkdir(new URL("../data/wikidata/", import.meta.url), { recursive: true });
  await writeFile(new URL(`../data/wikidata/${name}.csv`, import.meta.url), "﻿" + out.join("\r\n") + "\r\n");
  console.log(`${name}: ${out.length - 1} kişi (${skippedDia} kişi DİA'da olduğu için alınmadı; ${standardized} ad DİA yazımına çevrildi; ${mended} Batılı adın yazım hatası düzeltildi)${skipped.length ? `; atlanan sorgu: ${skipped.length}` : ""}`);
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

// node scripts/wikidata.mjs --links: add the notability (#sitelinks) to the links of existing files.
if (process.argv[2] === "--links") {
  for (const name of Object.keys(FIELDS)) {
    if (process.argv[3] && process.argv[3] !== name) continue;
    const url = new URL(`../data/wikidata/${name}.csv`, import.meta.url);
    let text;
    try {
      text = await readFile(url, "utf8");
    } catch {
      continue;
    }
    const lines = text.replace(/^\ufeff/, "").split(/\r?\n/).filter((l) => l.trim());
    const ids = lines.slice(1).map((l) => /wiki\/(Q\d+)(?:#\d+)?\s*$/.exec(l)?.[1]).filter(Boolean);
    const links = new Map();
    for (let k = 0; k < ids.length; k += 400) {
      const batch = ids.slice(k, k + 400);
      const rows = await run({ raw: `SELECT ?p ?n WHERE { VALUES ?p { ${batch.map((id) => `wd:${id}`).join(" ")} } ?p wikibase:sitelinks ?n . }` });
      for (const r of rows) links.set(r.p.value.split("/").pop(), r.n.value);
      await new Promise((r) => setTimeout(r, 500));
    }
    const out = [lines[0], ...lines.slice(1).map((l) => l.replace(/wiki\/(Q\d+)(?:#\d+)?\s*$/, (m, id) => `wiki/${id}${links.has(id) ? `#${links.get(id)}` : ""}`))];
    await writeFile(url, "\ufeff" + out.join("\r\n") + "\r\n");
    console.log(`${name}: ${links.size} / ${ids.length} kişiye tanınırlık eklendi`);
  }
} else if (process.argv[2] === "--compare") {
  // node scripts/wikidata.mjs --compare old-index.json: fails when a field lost more than a tenth of its people.
  const before = JSON.parse(await readFile(process.argv[3], "utf8"));
  const after = JSON.parse(await readFile(new URL("../data/wikidata/index.json", import.meta.url), "utf8"));
  let shrunk = false;
  for (const old of before) {
    const now = after.find((e) => e.id === old.id);
    const count = now?.count ?? 0;
    const ok = count >= old.count * 0.9;
    console.log(`${old.id}: ${old.count} → ${count}${ok ? "" : "  (çok azaldı)"}`);
    if (!ok) shrunk = true;
  }
  if (shrunk) {
    console.log("Bir alan onda birden fazla küçüldü: Wikidata sorgularından biri yarım kalmış olabilir. Kataloglar değiştirilmedi.");
    process.exit(1);
  }
} else if (process.argv[2] === "--index") {
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
