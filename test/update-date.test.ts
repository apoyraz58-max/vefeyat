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
