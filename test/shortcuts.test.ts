import { describe, expect, it } from "vitest";

describe("keyboard-shortcut manifest functions", () => {
  it("every go_ function in the trial manifest is registered and parses", async () => {
    const { readFileSync } = await import("node:fs");
    const { GO_FUNCTIONS, goQuery } = await import("../src/addin/actions.js");
    const { parseLaunch } = await import("../src/addin/logic.js");
    const xml = readFileSync(new URL("../manifest/manifest-kisayol.template.xml", import.meta.url), "utf8");
    const used = [...new Set([...xml.matchAll(/<FunctionName>(go_\w+)<\/FunctionName>/g)].map((m) => m[1]!))];
    expect(used.length).toBeGreaterThan(20);
    for (const name of used) expect(GO_FUNCTIONS).toContain(name);
    expect(goQuery("go_action_insert")).toBe("?action=insert");
    expect(parseLaunch(goQuery("go_tool_dizinsil")!)).toEqual({ view: "teslim", action: null, tool: "dizinsil" });
    const shortcuts = JSON.parse(readFileSync(new URL("../manifest/shortcuts.json", import.meta.url), "utf8"));
    for (const s of shortcuts.shortcuts) expect(GO_FUNCTIONS).toContain(s.action);
  });
});
