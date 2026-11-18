import { describe, expect, it } from "vitest";
import { Catalog, NameIndex, fuseLinks, honorificAliases, type Person } from "../src/core/index.js";
import { lookup } from "../src/addin/logic.js";
import { buildInventory } from "../src/addin/inventory.js";

const p = (shortName: string, fullName: string, deathDate: string): Person => ({ shortName, fullName, originalName: "", deathDate, description: "", source: "x" + shortName });
const catalog = new Catalog([
  p("Pezdevî, Ebü’l-Usr", "Ebü’l-Hasen Ebü’l-Usr Fahrü’l-İslâm Alî b. Muhammed b. el-Hüseyn b. Abdilkerîm el-Pezdevî", "(ö. 482/1089)"),
  p("Pezdevî, Ebü’l-Yüsr", "Sadrü’l-İslâm Ebü’l-Yüsr Muhammed b. Muhammed b. el-Hüseyn b. Abdilkerîm el-Pezdevî", "(ö. 493/1100)"),
  p("Gazzâlî", "Hüccetü’l-İslâm Ebû Hâmid Muhammed b. Muhammed b. Muhammed b. Ahmed el-Gazzâlî et-Tûsî", "(ö. 505/1111)"),
]);

describe("lakap + nisba, joined or apart", () => {
  it("helpers", () => {
    expect(fuseLinks("Fahrü’l-İslâm el-Pezdevî")).toBe("Fahrülİslâm el-Pezdevî");
    expect(honorificAliases(catalog.people[0]!.fullName)).toContain("Fahrü’l-İslâm el-Pezdevî");
    expect(honorificAliases(catalog.people[0]!.fullName)).toContain("Fahrü’l-İslâm Ebü’l-Usr el-Pezdevî");
  });
  it("selected name: the right Pezdevî is found and chosen", () => {
    for (const alias of ["Fahrülislâm el-Pezdevî", "Fahrülislam Pezdevî", "Fahrü'l-İslâm Pezdevî", "Fahrülislâm Pezdevî"]) {
      const found = lookup(catalog, alias);
      expect(found.automatic?.shortName, alias).toBe("Pezdevî, Ebü’l-Usr");
    }
    expect(lookup(catalog, "Hüccetülislâm Gazzâlî").automatic?.shortName).toBe("Gazzâlî");
    expect(lookup(catalog, "Sadrülislâm Ebü'l-Yüsr el-Pezdevî").candidates[0]?.shortName).toBe("Pezdevî, Ebü’l-Yüsr");
  });
  it("running text", () => {
    const text = "Fahrülislâm Pezdevî bu konuda ayrışır. Sadrülislâm Ebü'l-Yüsr el-Pezdevî de katılır. Fahrülislâm Ebü'l-Usr el-Pezdevî ise böyle der. Hüccetülislâm Gazzâlî de.";
    const rows = buildInventory([{ part: "main", anchor: 0, order: 0, text }], catalog, new NameIndex(catalog.people));
    expect(rows.map((r) => [r.text, r.person?.shortName])).toEqual([
      ["Fahrülislâm Pezdevî", "Pezdevî, Ebü’l-Usr"],
      ["Sadrülislâm Ebü'l-Yüsr el-Pezdevî", "Pezdevî, Ebü’l-Yüsr"],
      ["Fahrülislâm Ebü'l-Usr el-Pezdevî", "Pezdevî, Ebü’l-Usr"],
      ["Hüccetülislâm Gazzâlî", "Gazzâlî"],
    ]);
  });
});
