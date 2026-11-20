import { describe, expect, it } from "vitest";
import { leadingDate, sameDate } from "../src/addin/logic.js";

describe("leadingDate / sameDate", () => {
  it("reads the date right after a name", () => {
    expect(leadingDate(" (ö. 505/1112) der.")).toBe("(ö. 505/1112)");
    expect(leadingDate(" (1913-1998) bu konuyu")).toBe("(1913-1998)");
    expect(leadingDate(" der (ö. 505/1111).")).toBeNull();
    expect(leadingDate(" (bk. s. 12) der.")).toBeNull();
    expect(leadingDate(" (ö. 505/1111\r) der.")).toBeNull();
  });
  it("same date apart from spacing and case", () => {
    expect(sameDate("(ö. 505/1111)", "(Ö.505/1111)")).toBe(true);
    expect(sameDate("(ö. 505/1111)", "(v. 505/1111)")).toBe(false);
  });
});

describe("d. (died) for English texts", () => {
  it("formats, is recognised as a date and as a prefix variant", async () => {
    const { formatDate, startsWithDate, requirePrefix } = await import("../src/core/index.js");
    const { checkDate } = await import("../src/addin/health.js");
    expect(requirePrefix("d")).toBe("d");
    expect(formatDate("(ö. 505/1111)", "d")).toBe("(d. 505/1111)");
    expect(formatDate("(1844-1921)", "d")).toBe("(1844-1921)");
    expect(startsWithDate(" (d. 505/1111) argues")).toBe(true);
    const person = { shortName: "Gazzâlî", fullName: "", originalName: "", deathDate: "(ö. 505/1111)", description: "", source: "" };
    expect(checkDate("(ö. 505/1111)", person, "d")).toEqual({ health: "prefix", expected: "(d. 505/1111)" });
    expect(checkDate(" (d. 505/1111)", person, "d").health).toBe("ok");
  });
});

describe("buildInventoryAsync", () => {
  it("gives the same rows as buildInventory", async () => {
    const { buildInventory, buildInventoryAsync } = await import("../src/addin/inventory.js");
    const { Catalog } = await import("../src/core/index.js");
    const p = (shortName: string, deathDate: string) => ({ shortName, fullName: "", originalName: "", deathDate, description: "", source: "x" + shortName });
    const catalog = new Catalog([p("Gazzâlî", "(ö. 505/1111)"), p("Cüveynî", "(ö. 478/1085)")]);
    const segments = Array.from({ length: 300 }, (_, k) => ({ part: "main" as const, anchor: k, order: 0, text: k % 2 ? "Gazzâlî der." : "Cüveynî (ö. 478/1085) der." }));
    let reported = 0;
    const rows = await buildInventoryAsync(segments, catalog, undefined, {}, [], {}, () => reported++);
    expect(rows).toEqual(buildInventory(segments, catalog));
  });
});

describe("titles before a selected name", () => {
  it("Hz., Hazret-i, İmam, Şeyh are left out; a title that is part of the name stays", async () => {
    const { lookup, splitTitle } = await import("../src/addin/logic.js");
    const { Catalog } = await import("../src/core/index.js");
    const p = (shortName: string, deathDate: string) => ({ shortName, fullName: "", originalName: "", deathDate, description: "", source: "x" + shortName });
    const cat = new Catalog([p("Hatice", "(ö. 3/619)"), p("Molla Hüsrev", "(ö. 885/1480)"), p("Şâfiî", "(ö. 204/820)")]);
    expect(splitTitle("Hz. Hatice")).toEqual(["Hz. ", "Hatice"]);
    expect(splitTitle("Hz.Hatice")).toEqual(["Hz.", "Hatice"]);
    expect(splitTitle("İmam Şâfiî")).toEqual(["İmam ", "Şâfiî"]);
    expect(splitTitle("Hatice")).toEqual(["", "Hatice"]);
    for (const alias of ["Hz. Hatice", "Hz. Hatice'nin", "Hazret-i Hatice", "Hz.Hatice"]) expect(lookup(cat, alias).automatic?.shortName, alias).toBe("Hatice");
    expect(lookup(cat, "İmam Şâfiî").automatic?.shortName).toBe("Şâfiî");
    expect(lookup(cat, "Molla Hüsrev").automatic?.shortName).toBe("Molla Hüsrev");
    expect(lookup(cat, "Molla Hüsrev").title).toBeUndefined();
  });
});

describe("Tarih biçimi", () => {
  it("source, Gregorian only, death year only; with any Kısaltma", async () => {
    const { formatDate, alreadyHasDate } = await import("../src/core/index.js");
    expect(formatDate("(ö. 505/1111)", "o", "gregorian")).toBe("(ö. 1111)");
    expect(formatDate("(ö. 505/1111)", "v", "death")).toBe("(v. 1111)");
    expect(formatDate("(1858-1917)", "o", "gregorian")).toBe("(1858-1917)");
    expect(formatDate("(1858-1917)", "o", "death")).toBe("(ö. 1917)");
    expect(formatDate("(1858-1917)", "d", "death")).toBe("(d. 1917)");
    expect(formatDate("(m.ö. 384-322)", "o", "death")).toBe("(m.ö. 384-322)");
    expect(formatDate("(ö. 505/1111)", "o")).toBe("(ö. 505/1111)");
    expect(alreadyHasDate(" (ö. 1111) der", "(ö. 505/1111)")).toBe(true);
  });
});
