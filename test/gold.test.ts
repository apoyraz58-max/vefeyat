/**
 * Academic challenge set of the VBA add-in (test/fixtures/academic-gold.json:
 * 48 hand-labelled texts, 32 person mentions, 30 identity queries) run on the
 * TypeScript scanner with the catalog it was labelled against
 * (data/catalog.csv, SHA-256 d70e907b…). Metrics follow the VBA harness
 * scripts/tests/check_word_academic.py: exact UTF-16 spans, strict helpers,
 * footnotes through citation islands. The VBA 0.3.41 baseline is
 * strict precision 0.96 (24 tp, 1 fp), recall 0.75 (8 fn); identity: target
 * first for 21 queries, 5 automatic acceptances, all correct.
 */
import { createHash } from "node:crypto";
import { existsSync, readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { buildInventory, notePeopleOrMatches } from "../src/addin/inventory.js";
import { Catalog, NameIndex, confirmationCandidates, safeCandidate } from "../src/core/index.js";

const catalogPath = new URL("../data/catalog.csv", import.meta.url);
const gold = JSON.parse(readFileSync(new URL("./fixtures/academic-gold.json", import.meta.url), "utf8")) as {
  catalog_sha256: string;
  documents: Array<{ id: string; category: string; text: string; story: string; entities: Array<{ start_utf16: number; end_utf16: number; catalog: boolean }> }>;
  queries: Array<{ id: string; query: string; target_article: string | null; must_confirm: boolean }>;
};

const haveCatalog = existsSync(catalogPath);

describe.runIf(haveCatalog)("VBA academic challenge set", () => {
  const csv = haveCatalog ? readFileSync(catalogPath, "utf8") : "";
  const catalog = haveCatalog ? Catalog.fromCsv(csv) : new Catalog([]);
  const index = new NameIndex(catalog.people);

  it("uses the catalog the set was labelled against", () => {
    expect(createHash("sha256").update(readFileSync(catalogPath)).digest("hex")).toBe(gold.catalog_sha256);
  });

  let tp = 0;
  let fp = 0;
  let fn = 0;
  const errors: string[] = [];
  for (const doc of gold.documents) {
    const hits = notePeopleOrMatches({ part: doc.story === "footnote" ? "note" : "main", text: doc.text }, index);
    const wanted = new Set(doc.entities.map((e) => `${e.start_utf16}:${e.end_utf16}`));
    const actual = new Set(hits.map((h) => `${h.start}:${h.start + h.length}`));
    const hit = [...wanted].filter((x) => actual.has(x)).length;
    tp += hit;
    fp += actual.size - hit;
    fn += wanted.size - hit;
    if (hit !== wanted.size || hit !== actual.size) {
      errors.push(`${doc.id} ${doc.category}: beklenen ${[...wanted].join(",") || "-"}, bulunan ${[...actual].join(",") || "-"}`);
    }
  }
  const precision = tp / (tp + fp);
  const recall = tp / (tp + fn);

  it("strict mentions at least as good as VBA 0.3.41 (P 0.96, R 0.75)", () => {
    console.log(`strict: tp ${tp}, fp ${fp}, fn ${fn}, precision ${precision.toFixed(3)}, recall ${recall.toFixed(3)}\n  ${errors.join("\n  ")}`);
    expect(precision).toBeGreaterThanOrEqual(0.96);
    expect(recall).toBeGreaterThanOrEqual(0.75);
  });

  it("Şahıs dökümü discovery at least as good as VBA 0.3.41 (P 0.964, R 0.844)", () => {
    let dtp = 0;
    let dfp = 0;
    let dfn = 0;
    const misses: string[] = [];
    for (const doc of gold.documents) {
      const rows = buildInventory([{ part: doc.story === "footnote" ? "note" : "main", anchor: 0, order: 0, text: doc.text }], catalog, index);
      const wanted = new Set(doc.entities.map((e) => `${e.start_utf16}:${e.end_utf16}`));
      const actual = new Set(rows.map((r) => `${r.start}:${r.start + r.length}`));
      const hit = [...wanted].filter((x) => actual.has(x)).length;
      dtp += hit;
      dfp += actual.size - hit;
      dfn += wanted.size - hit;
      if (hit !== wanted.size || hit !== actual.size) misses.push(`${doc.id} ${doc.category}`);
    }
    const p = dtp / (dtp + dfp);
    const r = dtp / (dtp + dfn);
    console.log(`inventory: tp ${dtp}, fp ${dfp}, fn ${dfn}, precision ${p.toFixed(3)}, recall ${r.toFixed(3)}; ${misses.join(", ")}`);
    expect(p).toBeGreaterThanOrEqual(0.964);
    expect(r).toBeGreaterThanOrEqual(0.84375);
  });

  it("identity queries at least as good as VBA 0.3.41", () => {
    let top1 = 0;
    let auto = 0;
    let wrongAuto = 0;
    let retrieved = 0;
    let targets = 0;
    const misses: string[] = [];
    for (const q of gold.queries) {
      const candidates = confirmationCandidates(catalog.find(q.query, false), q.query);
      const articles = candidates.map((p) => p.source.replace(/\/+$/, "").split("/").pop());
      const pick = safeCandidate(candidates, q.query);
      const automatic = pick !== null;
      if (q.target_article) {
        targets++;
        if (articles[0] === q.target_article) top1++;
        else misses.push(`${q.id} ${q.query}: ${articles.slice(0, 3).join(", ")}`);
        if (articles.includes(q.target_article)) retrieved++;
      }
      if (automatic) {
        auto++;
        const picked = pick?.source.replace(/\/+$/, "").split("/").pop();
        if (!(q.target_article && picked === q.target_article && !q.must_confirm)) wrongAuto++;
      }
    }
    console.log(`identity: top1 ${top1}/${targets}, retrieved ${retrieved}/${targets}, automatic ${auto} (wrong ${wrongAuto})\n  ${misses.join("\n  ")}`);
    expect(top1).toBeGreaterThanOrEqual(21);
    expect(retrieved).toBe(targets);
    expect(wrongAuto).toBe(0);
    expect(auto).toBeGreaterThanOrEqual(5);
  });
});
