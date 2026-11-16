import { describe, expect, it } from "vitest";
import { buildInventory } from "../src/addin/inventory.js";
import { isLocal, normalizeDate, parseLocalPersons, upsertLocal, validateLocal, withLocal } from "../src/addin/persons.js";
import { Catalog } from "../src/core/index.js";

describe("own people", () => {
  it("dates as Vefeyât writes them", () => {
    expect(normalizeDate("1950")).toBe("(ö. 1950)");
    expect(normalizeDate("ö. 505/1111")).toBe("(ö. 505/1111)");
    expect(normalizeDate("1890 - 1950")).toBe("(1890-1950)");
    expect(normalizeDate("(v. 1950)")).toBe("(v. 1950)");
    expect(normalizeDate("")).toBe("");
    expect(normalizeDate("bilinmiyor")).toBeNull();
  });
  it("validation and storage", () => {
    expect(validateLocal({ shortName: "A", fullName: "", deathDate: "", description: "" }).ok).toBe(false);
    expect(validateLocal({ shortName: "Asım Köksal", fullName: "", deathDate: "x", description: "" }).ok).toBe(false);
    const ok = validateLocal({ shortName: "  Asım  Köksal ", fullName: "Asım Köksal", deathDate: "1913-1998", description: "Tarihçi" });
    expect(ok.ok && ok.person).toEqual({ shortName: "Asım Köksal", fullName: "Asım Köksal", deathDate: "(1913-1998)", description: "Tarihçi" });
    const list = upsertLocal([], (ok as { person: never }).person);
    expect(parseLocalPersons(JSON.stringify(list))).toEqual(list);
    expect(upsertLocal(list, { ...list[0]!, deathDate: "(ö. 1998)" })).toHaveLength(1);
  });
  it("found in the text like catalog people", () => {
    const catalog = withLocal(new Catalog([]), [{ shortName: "Asım Köksal", fullName: "", deathDate: "(1913-1998)", description: "" }]);
    const rows = buildInventory([{ part: "main", anchor: 0, order: 0, text: "Asım Köksal bu konuyu inceler." }], catalog);
    expect(rows.map((r) => [r.text, r.person && isLocal(r.person)])).toEqual([["Asım Köksal", true]]);
  });
});
