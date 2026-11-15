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
