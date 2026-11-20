import { describe, expect, it } from "vitest";
import { lifespan, searchForms, searchWikidata } from "../src/addin/wikidata-live.js";
import { Catalog } from "../src/core/index.js";
import { withLocal } from "../src/addin/persons.js";

describe("Wikidata'da ara", () => {
  it("searches the written form, a plain one and the English article", () => {
    const forms = searchForms("Hasan et-Türâbî");
    expect(forms).toContain("Hasan et-Türâbî");
    expect(forms).toContain("Hasan et-Turabi");
    expect(forms).toContain("Hasan al-Turabi");
    expect(forms).toContain("Hasan Turabi");
  });

  it("writes dates as the catalogs do", () => {
    expect(lifespan(1932, 2016, 11)).toBe("(1932-2016)");
    expect(lifespan(null, 2016, 9)).toBe("(ö. 2016)");
    expect(lifespan(-384, -322, 9)).toBe("(m.ö. 384-322)");
    expect(lifespan(null, 550, 7)).toBe("(ö. 6. yy)");
    expect(lifespan(null, 2016, 6)).toBe("");
  });

  it("returns people with a known death, the name's bearer first", async () => {
    const calls: string[] = [];
    const fake = async (url: string) => {
      calls.push(url);
      if (url.includes("wbsearchentities")) {
        const search = new URL(url).searchParams.get("search");
        return { ok: true, json: async () => ({ search: search === "Hasan al-Turabi" ? [{ id: "Q469299" }, { id: "Q1" }] : [{ id: "Q2" }] }) };
      }
      return {
        ok: true,
        json: async () => ({
          results: {
            bindings: [
              { p: { value: "http://www.wikidata.org/entity/Q2" }, links: { value: "90" }, death: { value: "1950-01-01T00:00:00Z" }, dprec: { value: "9" }, en: { value: "Someone Else" } },
              {
                p: { value: "http://www.wikidata.org/entity/Q469299" },
                links: { value: "34" },
                death: { value: "2016-03-05T00:00:00Z" },
                dprec: { value: "11" },
                birth: { value: "1932-02-01T00:00:00Z" },
                bprec: { value: "11" },
                tr: { value: "Hasan el-Turabi" },
                en: { value: "Hasan al-Turabi" },
                enDesc: { value: "Sudanese politician" },
              },
            ],
          },
        }),
      };
    };
    const hits = await searchWikidata("Hasan et-Türâbî", fake);
    expect(hits[0]).toMatchObject({ id: "Q469299", label: "Hasan el-Turabi", english: "Hasan al-Turabi", deathDate: "(1932-2016)", description: "Sudanese politician" });
    expect(hits.map((h) => h.id)).not.toContain("Q1"); // no death known: left out
    expect(calls.some((c) => c.includes("query.wikidata.org"))).toBe(true);
  });

  it("falls back to SPARQL's search when the API refuses", async () => {
    const fake = async (url: string) => {
      if (url.includes("wbsearchentities")) return { ok: false, json: async () => ({}) };
      if (decodeURIComponent(url).includes("mwapi")) return { ok: true, json: async () => ({ results: { bindings: [{ p: { value: "http://www.wikidata.org/entity/Q5" }, rank: { value: "0" }, form: { value: "0" } }] } }) };
      return { ok: true, json: async () => ({ results: { bindings: [{ p: { value: "http://www.wikidata.org/entity/Q5" }, links: { value: "3" }, death: { value: "1977-01-01T00:00:00Z" }, dprec: { value: "9" }, tr: { value: "Ali Şeriati" } }] } }) };
    };
    expect((await searchWikidata("Ali Şeriati", fake)).map((h) => `${h.label} ${h.deathDate}`)).toEqual(["Ali Şeriati (ö. 1977)"]);
  });

  it("says so when Wikidata cannot be reached", async () => {
    await expect(searchWikidata("Ali Şeriati", async () => ({ ok: false, json: async () => ({}) }))).rejects.toThrow();
  });

  it("the added record finds the name as written", () => {
    const catalog = withLocal(Catalog.fromCsv("Kısa Ad;Tam Ad;Orijinal Ad;Vefat Tarihi;Tanıtım;Kaynak Link\r\nGazzâlî;---;---;(ö. 505/1111);;x"), [
      { shortName: "Hasan el-Turabi", fullName: "Hasan al-Turabi | Hasan et-Türâbî", deathDate: "(1932-2016)", description: "Wikidata Q469299" },
    ]);
    expect(catalog.find("Hasan et-Türâbî", false).map((p) => p.shortName)).toEqual(["Hasan el-Turabi"]);
    expect(catalog.find("Türâbî", false).map((p) => p.shortName)).toEqual(["Hasan el-Turabi"]);
  });
});
