import { describe, expect, it } from "vitest";
import { TagIndex, checkControls, checkDate, repairText, repairable } from "../src/addin/health.js";
import { dateTag } from "../src/addin/logic.js";
import { Catalog, type Person } from "../src/core/index.js";

const gazzali: Person = { shortName: "Gazzâlî", fullName: "", originalName: "", deathDate: "(ö. 505/1111)", description: "", source: "https://islamansiklopedisi.org.tr/gazzali" };
const catalog = new Catalog([gazzali]);

describe("checkDate", () => {
  it.each([
    [" (ö. 505/1111)", "o", "ok"],
    [" (v. 505/1111)", "o", "prefix"],
    [" (ö. 505/1112)", "o", "changed"],
    ["", "o", "empty"],
    [" sözde", "o", "broken"],
    [" (v. 505/1111)", "v", "ok"],
  ] as const)("%j with %s → %s", (text, prefix, health) => {
    expect(checkDate(text, gazzali, prefix).health).toBe(health);
  });
  it("unknown person", () => {
    expect(checkDate(" (ö. 1)", null, "o")).toEqual({ health: "unknown", expected: "" });
  });
});

describe("checkControls", () => {
  it("finds the person by tag, repairs keep the leading space", () => {
    const rows = checkControls(
      [
        { id: 1, part: "main", tag: dateTag(gazzali), text: " (v. 505/1111)" },
        { id: 2, part: "note", tag: "vefeyat:date:00000000", text: " (ö. 1)" },
      ],
      new TagIndex(catalog),
      "o",
    );
    expect(rows.map((r) => [r.person?.shortName ?? null, r.health, repairable(r)])).toEqual([
      ["Gazzâlî", "prefix", true],
      [null, "unknown", false],
    ]);
    expect(repairText(rows[0]!, " (v. 505/1111)")).toBe(" (ö. 505/1111)");
  });
});

describe("VBA dates", async () => {
  const { isLegacyDateField, legacyPerson } = await import("../src/addin/health.js");
  const { NameIndex } = await import("../src/core/index.js");
  const ebuBekir: Person = { shortName: "Ebû Bekir er-Râzî", fullName: "", originalName: "", deathDate: "(ö. 313/925)", description: "", source: "x/razi-ebu-bekir" };
  const fahreddin: Person = { shortName: "Fahreddin er-Râzî", fullName: "", originalName: "", deathDate: "(ö. 606/1210)", description: "", source: "x/fahreddin" };
  const index = new NameIndex([gazzali, ebuBekir, fahreddin]);
  it("field codes", () => {
    expect(isLegacyDateField(" ADDIN Vefeyat.Date.D0123456789ABCDEF.P1 ")).toBe(true);
    expect(isLegacyDateField(" ADDIN ZOTERO_ITEM CSL_CITATION ")).toBe(false);
    expect(isLegacyDateField(" ADDIN Vefeyat.Date.D01.P1 extra ")).toBe(false);
  });
  it("the name right before the date, homonyms by the date", () => {
    expect(legacyPerson("Bu konuda Gazzâlî", " (ö. 505/1111)", index)).toMatchObject({ person: gazzali, literal: "Gazzâlî", start: 10 });
    expect(legacyPerson("Sonra Râzî", " (ö. 313/925)", index)?.person).toBe(ebuBekir);
    expect(legacyPerson("Sonra Râzî", " (ö. 1/1)", index)?.person).toBeNull();
    expect(legacyPerson("Gazzâlî bunu söyler ve", " (ö. 505/1111)", index)).toBeNull();
  });
});
