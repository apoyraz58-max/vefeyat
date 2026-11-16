import { describe, expect, it } from "vitest";
import { fieldsOf, sharedFields } from "../src/core/topics.js";

describe("fields", () => {
  it("of sentences and descriptions", () => {
    expect([...fieldsOf("Râzî tefsirinde bu meseleyi tartışır.")]).toEqual(["tefsir"]);
    expect([...fieldsOf("Râzî tıpta meşhurdur.")]).toEqual(["tıp"]);
    expect([...fieldsOf("Ünlü hekim ve filozof.")].sort()).toEqual(["felsefe", "tıp"]);
    expect(sharedFields(fieldsOf("Râzî tefsirinde"), "Kelâm, felsefe, tefsir ve usûl-i fıkıh alanlarına dair çalışmalarıyla tanınan Eş‘arî âlimi.")).toEqual(["tefsir"]);
  });
  it("short stems only as whole words", () => {
    expect(fieldsOf("Tipik bir örnek.").has("tıp")).toBe(false);
    expect(fieldsOf("Hattı takip etti.").has("hat")).toBe(false);
  });
});
