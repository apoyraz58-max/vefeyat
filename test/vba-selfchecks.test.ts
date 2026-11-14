/**
 * The VBA add-in ships two built-in self-checks that run before every
 * installation: `VfpRunPolicyChecks` (date formatting, identity) and
 * `VfaRunLookupChecks` (name lookup, selected names, reading order).
 * Each case below is one assertion of those procedures, with the VBA label,
 * so the TypeScript core is held to exactly the same behaviour.
 * Word-only assertions (field codes, bindings) belong to the add-in layer.
 */
import { describe, expect, it } from "vitest";
import {
  InvalidSettingError,
  MAX_NAME_LENGTH,
  confirmationCandidates,
  formatDate,
  nameCandidates,
  nameScore,
  orderBefore,
  personKey,
  selectionAlias,
  type Person,
} from "../src/core/index.js";

const person = (...f: [string, string, string, string, string, string]): Person => ({
  shortName: f[0],
  fullName: f[1],
  originalName: f[2],
  deathDate: f[3],
  description: f[4],
  source: f[5],
});

describe("VfpRunPolicyChecks: date marker formatting", () => {
  const raw = "  (ö.  680/1282) ";
  const cases: Array<[label: string, input: string, prefix: string, expected: string]> = [
    ["Original bytes", raw, "original", raw],
    ["Spaces", raw, "ol", "  (öl.  680/1282) "],
    ["V prefix", raw, "v", "  (v.  680/1282) "],
    ["Long marker", "(öl. 680/1282)", "o", "(ö. 680/1282)"],
    ["V marker", "(v. 680/1282)", "o", "(ö. 680/1282)"],
    ["Uppercase marker", "(ÖL. 680/1282)", "v", "(v. 680/1282)"],
    ["Adjacent year", "ö.680/1282", "v", "v.680/1282"],
    ["Uncertainty", "(v. ?)", "ol", "(öl. ?)"],
    ["Lifespan", "(1844-1921)", "ol", "(1844-1921)"],
    ["BC expression", "(m.ö. 384-322)", "v", "(m.ö. 384-322)"],
    ["Uppercase BC", "(M.Ö. 384-322)", "o", "(M.Ö. 384-322)"],
    ["Other event", "(gaybeti 260/874)", "ol", "(gaybeti 260/874)"],
    ["Unknown expression", "(tarih bilinmiyor)", "o", "(tarih bilinmiyor)"],
    ["Name substring", "Öztürk", "v", "Öztürk"],
    ["Marker boundary", "(ö.ztürk)", "v", "(ö.ztürk)"],
    ["Nonleading marker", "(not: ö. 680/1282)", "v", "(not: ö. 680/1282)"],
    ["Only leading marker", "(ö. 680/1282; kaynak: ö. 681)", "v", "(v. 680/1282; kaynak: ö. 681)"],
    ["Unicode spaces", " ( ö. 680/1282)", "v", " ( v. 680/1282)"],
    ["Empty expression", "", "o", ""],
    ["Bare marker", "v.", "ol", "v."],
    ["Normalized setting", "(V. 680/1282)", " OL ", "(öl. 680/1282)"],
  ];
  it.each(cases)("%s", (_label, input, prefix, expected) => {
    expect(formatDate(input, prefix)).toBe(expected);
  });

  it("Blank prefix rejected", () => {
    expect(() => formatDate("(v. 680)", "")).toThrow(InvalidSettingError);
  });
  it("Unknown prefix rejected", () => {
    expect(() => formatDate("(v. 680)", "invalid")).toThrow(InvalidSettingError);
  });
});

describe("VfpRunPolicyChecks: person identity", () => {
  const rowA = person("A:B", "C", "", "old date", "description", "url");
  it("Identity delimiters", () => {
    expect(personKey(rowA)).not.toBe(personKey(person("A", "B:C", "", "old date", "description", "url")));
  });
  it("Stable metadata identity", () => {
    expect(personKey(rowA)).toBe(personKey(person("A:B", "C", "", "new date", "new description", "url")));
  });
  it("Source identity", () => {
    expect(personKey(rowA)).not.toBe(personKey(person("A:B", "C", "", "old date", "description", "other url")));
  });
});

describe("VfaRunLookupChecks: candidates and ranks", () => {
  it("unique exact name, description-only candidate excluded", () => {
    const candidates = [
      person("Ali", "Ali b. Hasan", "", "(o. 10)", "", "one"),
      person("Alim", "Alim b. Hasan", "", "(o. 20)", "", "two"),
      person("Zeyd", "Zeyd b. Hasan", "", "(o. 30)", "Ali", "three"),
    ];
    const { exact, partial } = nameCandidates(candidates, "Ali");
    expect(exact).toHaveLength(1); // Unique exact name
    expect(exact[0]?.shortName).toBe("Ali"); // Exact identity
    expect(partial).toHaveLength(1); // Description-only candidate excluded
    // Confirmation retains exact and partial alternatives
    expect(confirmationCandidates(candidates, "Ali")).toHaveLength(2);
  });

  it("Exact name rank", () => {
    expect(nameScore("Abaka", "Abaka")).toBe(4);
  });
  it("Joined alias rank", () => {
    expect(nameScore("Kıvâmü’s-sünne", "Kıvâmüssünne")).toBe(3);
  });
  it("Joined alias within full name", () => {
    expect(nameScore("Ebu'l-Kasim Kıvâmü’s-sünne Ismail", "Kıvâmüssünne")).toBeGreaterThan(0);
  });

  it("Distinct homonyms remain ambiguous; duplicate identity removed", () => {
    const candidates = [
      person("Ali", "Ali b. Hasan", "", "(o. 10)", "", "one"),
      person("Ali", "Ali b. Huseyin", "", "(o. 20)", "", "two"),
      person("Ali", "Ali b. Hasan", "", "(o. 11)", "Updated date", "one"),
    ];
    const { exact, partial } = nameCandidates(candidates, "Ali");
    expect(exact).toHaveLength(2);
    expect(partial).toHaveLength(0);
  });

  it("Accent-folded and original Arabic names are exact", () => {
    const candidates = [person("Gazzâlî", "Abu Hamid", "الغزالي", "(o. 505)", "", "arabic")];
    expect(nameCandidates(candidates, "Gazzali").exact).toHaveLength(1);
    expect(nameCandidates(candidates, "الغزالي").exact).toHaveLength(1);
  });

  it("Substring requires a chooser", () => {
    const { exact, partial } = nameCandidates([person("Ali b. Hasan", "Ali b. Hasan", "", "(o. 10)", "", "one")], "Ali");
    expect(exact).toHaveLength(0);
    expect(partial).toHaveLength(1);
  });
});

describe("VfaRunLookupChecks: reading order of main text and footnotes", () => {
  const at = (anchor: number, part: 0 | 1, local: number) => ({ anchor, part, local });
  it.each([
    ["Main before later footnote", at(10, 0, 0), at(20, 1, 5), true],
    ["Footnote before following main", at(20, 1, 5), at(21, 0, 0), true],
    ["Main reference precedes footnote", at(20, 0, 0), at(20, 1, 1), true],
    ["Within-footnote order", at(20, 1, 3), at(20, 1, 9), true],
    ["Reverse note order rejected", at(20, 1, 9), at(20, 1, 3), false],
    ["Equal order is stable", at(20, 1, 3), at(20, 1, 3), false],
  ] as const)("%s", (_label, a, b, expected) => {
    expect(orderBefore(a, b)).toBe(expected);
  });
});

describe("VfaRunLookupChecks: selected name", () => {
  it.each([
    ["Trailing paragraph mark accepted", "Abaka\r", "Abaka"],
    ["Unicode edge spaces with paragraph accepted", " Ali   Hasan \r　", "Ali   Hasan"],
    ["Unicode edge spaces without paragraph accepted", " Abaka ", "Abaka"],
    ["Literal maximum accepted despite selection extras", `${"A".repeat(MAX_NAME_LENGTH)} \r`, "A".repeat(MAX_NAME_LENGTH)],
  ])("%s", (_label, selected, expected) => {
    expect(selectionAlias(selected)).toBe(expected);
  });

  it.each([
    ["Interior paragraph rejected", "Ali\rHasan"],
    ["Leading paragraph rejected", "\rAli"],
    ["Multiple paragraphs rejected", "Ali\r\r"],
    ["Spaced paragraphs rejected", "Ali\r \r"],
    ["Trailing tab rejected", "Ali\t"],
    ["Line feed rejected", "Ali\n"],
    ["CRLF rejected", "Ali\r\n"],
    ["Word cell marker rejected", "Ali\r\u0007"],
    ["Manual line break rejected", "Ali\u000b"],
    ["Other C0 control rejected", "Ali\u001b"],
    ["C1 control rejected", "Ali\u0080"],
    ["Unicode next-line control rejected", "Ali\u0085"],
    ["Unicode line separator rejected", "Ali "],
    ["Empty paragraph rejected", " \r "],
    ["Overlong literal rejected", `${"A".repeat(MAX_NAME_LENGTH + 1)}\r`],
  ])("%s", (_label, selected) => {
    expect(selectionAlias(selected)).toBeNull();
  });
});
