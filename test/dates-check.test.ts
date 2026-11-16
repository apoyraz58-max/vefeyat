import { describe, expect, it } from "vitest";
import { checkWrittenDate, deathYears } from "../src/core/dates.js";

describe("written dates", () => {
  it("parses the forms in the catalog", () => {
    expect(deathYears("(ö. 505/1111)")).toEqual({ hijri: [505], gregorian: [1111], bc: false });
    expect(deathYears("(ö. 305/917-18)")).toEqual({ hijri: [305], gregorian: [917, 918], bc: false });
    expect(deathYears("(1844-1921)")).toEqual({ hijri: [], gregorian: [1921], bc: false });
    expect(deathYears("(m.ö. 384-322)")).toEqual({ hijri: [], gregorian: [322], bc: true });
    expect(deathYears("(ö. 600 [?])")).toEqual({ hijri: [], gregorian: [], single: 600, bc: false });
  });
  it.each([
    ["(ö. 505/1111)", "(ö. 505/1111)", "ok"],
    ["(v. 505/1111)", "(ö. 505/1111)", "ok"],
    ["(ö. 1111)", "(ö. 505/1111)", "ok"],
    ["(ö. 505)", "(ö. 505/1111)", "ok"],
    ["(ö. 506/1112)", "(ö. 505/1111)", "different"],
    ["(ö. 505/1211)", "(ö. 505/1111)", "inconsistent"],
    ["(ö. 305/918)", "(ö. 305/917-18)", "ok"],
    ["(ö. 1112)", "(ö. 505/1111)", "different"],
    ["(1844-1921)", "(1844-1921)", "ok"],
    ["(ö. 1921)", "(1844-1921)", "ok"],
    ["(ö. 1920)", "(1844-1921)", "different"],
  ] as const)("%s vs catalog %s → %s", (written, catalog, expected) => {
    expect(checkWrittenDate(written, catalog)).toBe(expected);
  });
});
