// End-to-end smoke test of the built task pane in Chromium with a fake Word.
// It checks the UI wiring, not Word itself: the fake implements only the
// WordApi calls word.ts uses. Run after `npm run build`:  npm run smoke
import { createRequire } from "node:module";
import { createServer } from "node:http";
import { readFile } from "node:fs/promises";
import { extname } from "node:path";

const require = createRequire(import.meta.url);
const { chromium } = require("playwright");
const dist = new URL("../dist/addin/", import.meta.url);
const shotPath = process.env.SMOKE_SCREENSHOT;

const types = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".png": "image/png", ".map": "application/json" };
const server = createServer(async (req, res) => {
  try {
    const path = new URL(`.${new URL(req.url, "http://x").pathname}`, dist);
    res.writeHead(200, { "content-type": types[extname(path.pathname)] ?? "application/octet-stream" });
    res.end(await readFile(path));
  } catch {
    res.writeHead(404).end();
  }
});
await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
const base = `http://127.0.0.1:${server.address().port}/`;

const FAKE_OFFICE = String.raw`
(() => {
  const doc = (window.__doc = { text: "", sel: [0, 0], controls: [], settings: {} });
  const shift = (at, by) => { doc.controls.forEach((c) => { if (c.range.s >= at) { c.range.s += by; c.range.e += by; } }); };
  class R {
    constructor(s, e) { this.s = s; this.e = e; }
    load() { return this; }
    get text() { return doc.text.slice(this.s, this.e); }
    get paragraphs() {
      const end = doc.text.indexOf("\r", this.e);
      return { getLast: () => new R(this.s, end < 0 ? doc.text.length : end) };
    }
    getRange(loc) {
      if (loc === "End" || loc === "After") return new R(this.e, this.e);
      if (loc === "Start" || loc === "Before") return new R(this.s, this.s);
      return new R(this.s, this.e);
    }
    expandTo(r) { return new R(Math.min(this.s, r.s), Math.max(this.e, r.e)); }
    insertText(t, loc) {
      if (loc === "After") { doc.text = doc.text.slice(0, this.e) + t + doc.text.slice(this.e); shift(this.e, t.length); return new R(this.e, this.e + t.length); }
      if (loc === "Replace") { doc.text = doc.text.slice(0, this.s) + t + doc.text.slice(this.e); shift(this.e, t.length - (this.e - this.s)); return new R(this.s, this.s + t.length); }
      throw new Error("unsupported insert " + loc);
    }
    insertContentControl() {
      const range = this;
      const control = { tag: "", title: "", appearance: "", range, getRange: (loc) => (loc === "After" ? new R(range.e, range.e) : range) };
      doc.controls.push(control);
      return control;
    }
    select(mode) { doc.sel = mode === "Start" ? [this.s, this.s] : [this.s, this.e]; }
    compareLocationWith(o) {
      const value = this.e < o.s ? "Before" : this.e === o.s ? "AdjacentBefore" : this.s >= o.e ? "After" : "Overlaps";
      return { value };
    }
  }
  const search = (needle) => {
    const items = [];
    const hay = doc.text.toLocaleLowerCase("tr");
    const n = needle.replace(/\^\^/g, "^").toLocaleLowerCase("tr");
    for (let i = hay.indexOf(n); i >= 0; i = hay.indexOf(n, i + 1)) items.push(new R(i, i + n.length));
    return { items, load() { return this; } };
  };
  window.Word = {
    run: async (fn) => fn({ document: { getSelection: () => new R(doc.sel[0], doc.sel[1]), body: { search } }, sync: async () => {} }),
  };
  window.Office = {
    HostType: { Word: "Word" },
    EventType: { DocumentSelectionChanged: "selection" },
    AsyncResultStatus: { Succeeded: "succeeded" },
    onReady: (cb) => setTimeout(() => cb({ host: "Word" }), 0),
    context: {
      document: {
        settings: {
          get: (k) => doc.settings[k] ?? null,
          set: (k, v) => { doc.settings[k] = v; },
          saveAsync: (cb) => cb({ status: "succeeded" }),
        },
        addHandlerAsync: (_type, handler) => { window.__select = (s, e) => { doc.sel = [s, e]; handler(); }; },
      },
    },
  };
})();
`;

const CSV = [
  "Kısa Ad;Tam Ad;Orijinal Ad;Vefat Tarihi;Tanıtım;Kaynak Link",
  "Gazzâlî;Ebû Hâmid Muhammed b. Muhammed el-Gazzâlî;الغزالي;(ö. 505/1111);Eş'arî kelâmcısı, Şâfiî fakihi, mutasavvıf.;https://islamansiklopedisi.org.tr/gazzali",
  "Fahreddin er-Râzî;Ebû Abdillâh Fahreddîn Muhammed b. Ömer er-Râzî;;(ö. 606/1210);Eş'arî kelâmcısı, müfessir.;https://islamansiklopedisi.org.tr/fahreddin-er-razi",
  "Ebû Bekir er-Râzî;Ebû Bekir Muhammed b. Zekeriyyâ er-Râzî;;(ö. 313/925);Tabip ve filozof.;https://islamansiklopedisi.org.tr/razi-ebu-bekir",
].join("\r\n");

const checks = [];
const check = (label, condition, detail = "") => {
  checks.push({ label, ok: Boolean(condition), detail });
};

const browser = await chromium.launch();
try {
  const page = await browser.newPage({ viewport: { width: 360, height: 900 }, deviceScaleFactor: shotPath ? 2 : 1 });
  const errors = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await page.route("https://appsforoffice.microsoft.com/**", (route) => route.fulfill({ contentType: "text/javascript", body: FAKE_OFFICE }));
  await page.goto(`${base}taskpane.html`);
  await page.waitForFunction(() => document.getElementById("version")?.textContent);
  const text = "Gazzâlî bu görüşü savunur.\rSonra Gazzâlî der ki Gazali'nin eseri.\rRâzî de katılır.";
  await page.evaluate((t) => { window.__doc.text = t; }, text);
  const doc = () => page.evaluate(() => window.__doc.text);
  const select = async (needle, nth = 0) => {
    const current = await doc();
    let at = -1;
    for (let i = 0; i <= nth; i++) at = current.indexOf(needle, at + 1);
    await page.evaluate(([s, e]) => window.__select(s, e), [at, at + needle.length]);
    await page.waitForTimeout(400);
  };

  await page.setInputFiles("#catalog-file", { name: "people.csv", mimeType: "text/csv", buffer: Buffer.from(CSV, "utf8") });
  await page.waitForFunction(() => document.getElementById("catalog-status").textContent.includes("şahıs"));
  check("catalog loaded", (await page.textContent("#catalog-status")).includes("3 şahıs"), await page.textContent("#catalog-status"));

  await select("Gazzâlî");
  check("selected name shown", (await page.textContent("#selected-name")) === "Gazzâlî");
  check("single strong candidate chosen", await page.isEnabled("#insert"));
  await page.click("#insert");
  await page.waitForTimeout(200);
  check("date inserted at first mention", (await doc()).startsWith("Gazzâlî (ö. 505/1111) bu görüşü"), await doc());
  check("hidden tagged control", await page.evaluate(() => window.__doc.controls[0]?.appearance === "Hidden" && /^vefeyat:date:[0-9a-f]{8}$/.test(window.__doc.controls[0].tag)));

  await select("Gazzâlî", 1);
  await page.click("#insert");
  await page.waitForTimeout(200);
  check("earlier mention warning", await page.isVisible("#earlier"));
  check("nothing written before confirmation", !(await doc()).includes("Gazzâlî (ö. 505/1111) der"));
  await page.click("#insert-anyway");
  await page.waitForTimeout(200);
  check("insert anyway writes the date", (await doc()).includes("Sonra Gazzâlî (ö. 505/1111) der"), await doc());

  await select("Gazzâlî", 0);
  await page.click("#insert");
  await page.waitForTimeout(200);
  check("existing date is not duplicated", (await page.textContent("#message")).includes("zaten bir tarih var"));

  await select("Gazali'nin");
  check("weak spelling is not chosen automatically", !(await page.isVisible("#details")));
  await page.check("#candidate-0");
  const options = await page.$$eval("#spelling option", (o) => o.map((x) => x.value));
  check("spelling suggestion offered", options[0] === "Gazzâlî", options.join(" | "));
  await page.click("#respell");
  await page.waitForTimeout(300);
  check("spelling replaced with ending kept", (await doc()).includes("Gazzâlî'nin eseri"), await doc());

  await page.selectOption("#prefix", "v");
  await select("Râzî");
  const candidates = await page.$$eval("#candidates .candidate strong", (n) => n.map((x) => x.textContent));
  check("homonyms listed, well-known first", candidates[0]?.startsWith("Fahreddin er-Râzî") && candidates.length === 2, candidates.join(" | "));
  check("part-of-name hint", (await page.textContent("#candidates")).includes("adın bir parçası"));
  check("stale message cleared on new selection", await page.isHidden("#message"));
  check("choice required for homonyms", !(await page.isEnabled("#insert")));
  await page.check("#candidate-0");
  check("details shown", (await page.textContent("#detail-date")) === "(ö. 606/1210)");
  if (shotPath) await page.screenshot({ path: shotPath, fullPage: true });
  await page.click("#insert");
  await page.waitForTimeout(200);
  check("prefix setting applied", (await doc()).includes("Râzî (v. 606/1210) de"), await doc());
  check("prefix stored in document settings", (await page.evaluate(() => window.__doc.settings["vefeyat.prefix"])) === "v");
  check("no page errors", errors.length === 0, errors.join("; "));
} finally {
  await browser.close();
  server.close();
}

for (const c of checks) console.log(`${c.ok ? "PASS" : "FAIL"}  ${c.label}${c.ok || !c.detail ? "" : `  -> ${c.detail}`}`);
const failed = checks.filter((c) => !c.ok).length;
console.log(`${checks.length - failed}/${checks.length} smoke checks passed`);
process.exit(failed ? 1 : 0);
