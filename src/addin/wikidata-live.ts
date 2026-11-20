/**
 * Wikidata'da ara: a person no catalog has is looked up on Wikidata (CC0)
 * while the user waits, and the one they pick becomes a record of their own
 * (Kendi şahıslarım). Only the selected name leaves the computer.
 */
import { normalize } from "../core/index.js";

export interface WikidataHit {
  readonly id: string;
  readonly label: string;
  readonly english: string;
  readonly deathDate: string;
  readonly description: string;
  readonly links: number;
}

type Fetch = (url: string, init?: RequestInit) => Promise<Pick<Response, "ok" | "json">>;

const API = "https://www.wikidata.org/w/api.php";
const SPARQL = "https://query.wikidata.org/sparql";

/**
 * The forms to search for: as written, without circumflexes and Turkish
 * letters, with the Arabic article as Wikidata's English labels write it
 * ("et-Türâbî" → "al-Turabi"), and without the article.
 */
export function searchForms(name: string): string[] {
  const plain = name
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/ı/g, "i")
    .replace(/İ/g, "I")
    .replace(/[‘’ʿʾ`']/g, "")
    .replace(/\s+/g, " ")
    .trim();
  const article = plain.replace(/(^|\s)(e[lştsdrznl]|a[lştsdrznl]|ü[lştsdrznl])-/gi, "$1al-");
  const bare = plain.replace(/(^|\s)(e[lştsdrznl]|a[lştsdrznl]|ü[lştsdrznl])-/gi, "$1");
  return [...new Set([name.trim(), plain, article, bare].filter((f) => f.length >= 2))];
}

const year = (value: string | undefined): number | null => {
  const m = /^(-?)(\d+)-/.exec(value ?? "");
  return m ? Number(m[2]) * (m[1] ? -1 : 1) : null;
};

/** "(1932-2016)", "(ö. 2016)", "(m.ö. 384-322)", "(ö. 6. yy)"; "" when the date says too little. */
export function lifespan(birth: number | null, death: number | null, precision: number): string {
  if (death === null || precision < 7) return "";
  if (precision < 9) {
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

interface Binding {
  readonly [key: string]: { readonly value: string } | undefined;
}

async function sparql(query: string, fetcher: Fetch): Promise<Binding[]> {
  const response = await fetcher(`${SPARQL}?${new URLSearchParams({ query, format: "json" })}`, { headers: { Accept: "application/sparql-results+json" } }).catch(() => null);
  if (!response) throw new Error("Wikidata'ya ulaşılamadı.");
  if (!response.ok) throw new Error("Wikidata şu an yanıt vermiyor.");
  return ((await response.json()) as { results?: { bindings?: Binding[] } }).results?.bindings ?? [];
}

/** Search position of each item found (form, then rank); null when every request was refused. */
async function apiSearch(forms: string[], fetcher: Fetch): Promise<Map<string, number> | null> {
  const answers = await Promise.all(
    forms.map(async (form) => {
      const url = `${API}?${new URLSearchParams({ action: "wbsearchentities", search: form, language: "tr", uselang: "tr", type: "item", limit: "12", format: "json", origin: "*" })}`;
      try {
        const response = await fetcher(url);
        return response.ok ? ((await response.json()) as { search?: Array<{ id?: string }> }).search ?? [] : null;
      } catch {
        return null;
      }
    }),
  );
  if (answers.every((a) => a === null)) return null;
  const order = new Map<string, number>();
  answers.forEach((items, k) =>
    (items ?? []).forEach((item, rank) => {
      if (item.id && !order.has(item.id)) order.set(item.id, k * 100 + rank);
    }),
  );
  return order;
}

async function sparqlSearch(forms: string[], fetcher: Fetch): Promise<Map<string, number>> {
  const searches = forms
    .map(
      (form, k) => `{ SERVICE wikibase:mwapi { bd:serviceParam wikibase:endpoint "www.wikidata.org"; wikibase:api "EntitySearch"; mwapi:search ${JSON.stringify(form)}; mwapi:language "tr" .
    ?p wikibase:apiOutputItem mwapi:item . ?rank wikibase:apiOrdinal true . } BIND(${k} AS ?form) }`,
    )
    .join(" UNION ");
  const order = new Map<string, number>();
  for (const r of await sparql(`SELECT ?p ?rank ?form WHERE { ${searches} }`, fetcher)) {
    const id = (r.p?.value ?? "").split("/").pop() ?? "";
    const position = Number(r.form?.value ?? 0) * 100 + Number(r.rank?.value ?? 0);
    if (/^Q\d+$/.test(id) && (!order.has(id) || position < (order.get(id) ?? 0))) order.set(id, position);
  }
  return order;
}

/** People (with a known death) Wikidata finds under the name, best known first. */
export async function searchWikidata(name: string, fetcher: Fetch = fetch): Promise<WikidataHit[]> {
  // Wikidata's search for each form (its API, at once; SPARQL's slower search when the API refuses),
  // then the facts of what it found.
  const forms = searchForms(name);
  const order = (await apiSearch(forms, fetcher)) ?? (await sparqlSearch(forms, fetcher));
  const ids = [...order.keys()].sort((x, y) => (order.get(x) ?? 0) - (order.get(y) ?? 0)).slice(0, 40);
  if (ids.length === 0) return [];
  const rows = await sparql(
    `SELECT ?p ?links ?death ?dprec ?birth ?bprec ?tr ?en ?trDesc ?enDesc WHERE {
  VALUES ?p { ${ids.map((id) => `wd:${id}`).join(" ")} }
  ?p wdt:P31 wd:Q5 ; wikibase:sitelinks ?links .
  ?p p:P570/psv:P570 [ wikibase:timeValue ?death ; wikibase:timePrecision ?dprec ] .
  OPTIONAL { ?p p:P569/psv:P569 [ wikibase:timeValue ?birth ; wikibase:timePrecision ?bprec ] }
  OPTIONAL { ?p rdfs:label ?tr FILTER(LANG(?tr) = "tr") }
  OPTIONAL { ?p rdfs:label ?en FILTER(LANG(?en) = "en") }
  OPTIONAL { ?p schema:description ?trDesc FILTER(LANG(?trDesc) = "tr") }
  OPTIONAL { ?p schema:description ?enDesc FILTER(LANG(?enDesc) = "en") }
}`,
    fetcher,
  );
  const hits = new Map<string, WikidataHit>();
  for (const r of rows) {
    const id = (r.p?.value ?? "").split("/").pop() ?? "";
    if (!id || hits.has(id)) continue;
    const label = r.tr?.value ?? r.en?.value ?? "";
    const birth = Number(r.bprec?.value ?? 0) >= 9 ? year(r.birth?.value) : null;
    const deathDate = lifespan(birth, year(r.death?.value), Number(r.dprec?.value ?? 0));
    if (!label || !deathDate || /^Q\d+$/.test(label)) continue;
    hits.set(id, {
      id,
      label,
      english: r.en?.value && r.en.value !== label ? r.en.value : "",
      deathDate,
      description: r.trDesc?.value ?? r.enDesc?.value ?? "",
      links: Number(r.links?.value ?? 0),
    });
  }
  // The order of the search, then the best known: the one the user means is usually among the first.
  const wanted = normalize(name).split(" ").filter((w) => w.length > 2);
  const score = (h: WikidataHit): number => {
    const words = normalize(`${h.label} ${h.english}`);
    const shared = wanted.filter((w) => words.includes(w)).length;
    return shared * 1000 + Math.min(h.links, 99) * 2 - Math.min(order.get(h.id) ?? 0, 150);
  };
  return [...hits.values()].sort((a, b) => score(b) - score(a)).slice(0, 8);
}
