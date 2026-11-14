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

const CSV = [
  "Kısa Ad;Tam Ad;Orijinal Ad;Vefat Tarihi;Tanıtım;Kaynak Link",
  "Gazzâlî;Ebû Hâmid Muhammed b. Muhammed el-Gazzâlî;الغزالي;(ö. 505/1111);Eş'arî kelâmcısı, Şâfiî fakihi, mutasavvıf.;https://islamansiklopedisi.org.tr/gazzali",
  "Fahreddin er-Râzî;Ebû Abdillâh Fahreddîn Muhammed b. Ömer er-Râzî;;(ö. 606/1210);Eş'arî kelâmcısı, müfessir.;https://islamansiklopedisi.org.tr/fahreddin-er-razi",
  "Ebû Bekir er-Râzî;Ebû Bekir Muhammed b. Zekeriyyâ er-Râzî;;(ö. 313/925);Tabip ve filozof.;https://islamansiklopedisi.org.tr/razi-ebu-bekir",
].join("\r\n");

const types = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".png": "image/png", ".map": "application/json" };
let serveCatalog = true;
const server = createServer(async (req, res) => {
  try {
    if (req.url.startsWith("/catalog.csv")) {
      if (!serveCatalog) return res.writeHead(404).end();
      res.writeHead(200, { "content-type": "text/csv; charset=utf-8" });
      return res.end(CSV);
    }
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
  const doc = (window.__doc = Object.assign({ text: "", sel: [0, 0], controls: [], settings: {} }, window.__docInit || {}));
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
          refreshAsync: (cb) => cb({ status: "succeeded" }),
        },
        addHandlerAsync: (_type, handler) => { window.__select = (s, e) => { doc.sel = [s, e]; handler(); }; },
      },
    },
  };
})();
`;



const checks = [];
const check = (label, condition, detail = "") => {
  checks.push({ label, ok: Boolean(condition), detail });
};

const browser = await chromium.launch();
try {
  const context = await browser.newContext({ viewport: { width: 360, height: 900 }, deviceScaleFactor: shotPath ? 2 : 1 });
  const errors = [];
  const open = async (query, text = "", needle = "", options = {}) => {
    const at = options.last ? text.lastIndexOf(needle) : text.indexOf(needle);
    const tab = await context.newPage();
    tab.on("pageerror", (error) => errors.push(error.message));
    await tab.route("https://appsforoffice.microsoft.com/**", (route) => route.fulfill({ contentType: "text/javascript", body: FAKE_OFFICE }));
    await tab.addInitScript((init) => { window.__docInit = init; }, {
      text,
      sel: needle && at >= 0 ? [at, at + needle.length] : [0, 0],
      settings: options.settings ?? {},
    });
    await tab.goto(`${base}${options.page ?? "taskpane.html"}${query}`);
    await tab.waitForTimeout(600);
    return tab;
  };
  const docText = (tab) => tab.evaluate(() => window.__doc.text);
  const select = async (tab, needle, nth = 0) => {
    const current = await docText(tab);
    let at = -1;
    for (let i = 0; i <= nth; i++) at = current.indexOf(needle, at + 1);
    await tab.evaluate(([s, e]) => window.__select(s, e), [at, at + needle.length]);
    await tab.waitForTimeout(400);
  };

  // Main window: catalog published with the add-in loads by itself.
  const text = "Gazzâlî bu görüşü savunur.\rSonra Gazzâlî der ki Gazali'nin eseri.\rRâzî de katılır.";
  const page = await open("", text);
  check("embedded catalog loads automatically", (await page.textContent("#view-secili")) !== null && (await page.isHidden("#no-catalog")));
  check("no settings tab in the window", (await page.$$(".tabs")).length === 0);
  check("settings summary at the top", (await page.textContent("#settings-summary")) === "Kısaltma: ö. · Yalnız ilk kullanım · Ad biçimi: Her seferinde sor");

  await select(page, "Gazzâlî");
  check("single strong candidate chosen", await page.isEnabled("#insert"));
  await page.click("#insert");
  await page.waitForTimeout(200);
  check("date inserted at first mention", (await docText(page)).startsWith("Gazzâlî (ö. 505/1111) bu görüşü"), await docText(page));
  check("hidden tagged control", await page.evaluate(() => window.__doc.controls[0]?.appearance === "Hidden" && /^vefeyat:date:[0-9a-f]{8}$/.test(window.__doc.controls[0].tag)));

  await select(page, "Gazzâlî", 1);
  await page.click("#insert");
  await page.waitForTimeout(200);
  check("earlier mention warning", await page.isVisible("#earlier"));
  check("nothing written before confirmation", !(await docText(page)).includes("Gazzâlî (ö. 505/1111) der"));
  await page.click("#insert-anyway");
  await page.waitForTimeout(200);
  check("insert anyway writes the date", (await docText(page)).includes("Sonra Gazzâlî (ö. 505/1111) der"), await docText(page));

  await select(page, "Gazzâlî", 0);
  await page.click("#insert");
  await page.waitForTimeout(200);
  check("existing date is not duplicated", (await page.textContent("#message")).includes("zaten bir tarih var"));

  await select(page, "Gazali'nin");
  check("weak spelling is not chosen automatically", !(await page.isVisible("#details")));
  await page.check("#candidate-0");
  const options = await page.$$eval("#spelling option", (o) => o.map((x) => x.value));
  check("spelling suggestion offered", options[0] === "Gazzâlî", options.join(" | "));
  await page.click("#respell");
  await page.waitForTimeout(300);
  check("spelling replaced with ending kept", (await docText(page)).includes("Gazzâlî'nin eseri"), await docText(page));

  // Ribbon Tercihler menus write document settings through the function file.
  const commands = await open("", "", "", { page: "commands.html" });
  await commands.waitForFunction(() => typeof window.set_prefix_v === "function");
  const done = await commands.evaluate(() => new Promise((resolve) => window.set_prefix_v({ completed: () => resolve(window.__doc.settings["vefeyat.prefix"]) })));
  check("Kısaltma menu stores the setting", done === "v");
  check("all 14 setting commands registered", (await commands.evaluate(() => Object.keys(window).filter((k) => k.startsWith("set_")).length)) === 14);
  await commands.close();

  await page.evaluate(() => { window.__doc.settings["vefeyat.prefix"] = "v"; window.__doc.settings["vefeyat.nameForm"] = "full"; });
  await select(page, "Râzî");
  check("window shows changed settings", (await page.textContent("#settings-summary")).includes("Kısaltma: v.") && (await page.textContent("#settings-summary")).includes("Tam / uzun ad"));
  const candidates = await page.$$eval("#candidates .candidate strong", (n) => n.map((x) => x.textContent));
  check("homonyms listed, well-known first", candidates[0]?.startsWith("Fahreddin er-Râzî") && candidates.length === 2, candidates.join(" | "));
  check("choice required for homonyms", !(await page.isEnabled("#insert")));
  await page.check("#candidate-0");
  check("Ad biçimi puts the full name first", (await page.$eval("#spelling option", (o) => o.textContent)).includes("Tam / uzun ad"));
  if (shotPath) await page.screenshot({ path: shotPath, fullPage: true });
  await page.click("#insert");
  await page.waitForTimeout(200);
  check("Kısaltma applied to the new date", (await docText(page)).includes("Râzî (v. 606/1210) de"), await docText(page));
  check("no page errors", errors.length === 0, errors.join("; "));
  await page.close();

  // Tarih ekleme setting.
  let tab = await open("", "Gazzâlî önce. Sonra Gazzâlî.", "Gazzâlî", { last: true, settings: { "vefeyat.scope": "none" } });
  await tab.click("#insert");
  await tab.waitForTimeout(200);
  check("Yalnız denetle never writes", (await tab.textContent("#message")).includes("Yalnız denetle") && (await docText(tab)) === "Gazzâlî önce. Sonra Gazzâlî.");
  await tab.close();
  tab = await open("", "Gazzâlî önce. Sonra Gazzâlî.", "Gazzâlî", { last: true, settings: { "vefeyat.scope": "all" } });
  await tab.click("#insert");
  await tab.waitForTimeout(200);
  check("Bütün kullanımlar skips the earlier-mention warning", (await docText(tab)) === "Gazzâlî önce. Sonra Gazzâlî (ö. 505/1111).", await docText(tab));
  await tab.close();

  // Ribbon buttons.
  tab = await open("?action=insert", "İlk kez Gazzâlî burada.", "Gazzâlî");
  check("ribbon insert writes at once when unambiguous", (await docText(tab)) === "İlk kez Gazzâlî (ö. 505/1111) burada.", await docText(tab));
  await tab.close();
  tab = await open("?action=insert", "Râzî der ki.", "Râzî");
  check("ribbon insert asks for a choice with homonyms", (await tab.textContent("#message")).includes("Birden fazla aday") && (await docText(tab)) === "Râzî der ki.");
  await tab.close();
  tab = await open("?action=insert", "Gazzâlî önce. Sonra Gazzâlî.", "Gazzâlî", { last: true });
  check("ribbon insert still warns about an earlier mention", (await tab.isVisible("#earlier")) && (await docText(tab)) === "Gazzâlî önce. Sonra Gazzâlî.");
  await tab.close();
  tab = await open("?tool=dizin");
  check("Teslim menu opens the tools window", (await tab.isVisible("#view-teslim")) && (await tab.getAttribute("#tool-dizin", "class")) === "focus" && (await tab.textContent("#view-title")) === "Teslim araçları");
  await tab.click("#back");
  check("back returns to the selected name", await tab.isVisible("#view-secili"));
  await tab.close();
  tab = await open("?view=yardim");
  check("Kılavuz opens the guide", await tab.isVisible("#view-yardim"));
  await tab.close();

  // Katalog: user's own CSV replaces the published one until reset.
  tab = await open("?view=katalog");
  check("catalog window shows the published catalog", (await tab.textContent("#catalog-status")).includes("eklentiyle gelen katalog"));
  await tab.setInputFiles("#catalog-file", { name: "benim.csv", mimeType: "text/csv", buffer: Buffer.from(CSV.split("\r\n").slice(0, 2).join("\r\n"), "utf8") });
  await tab.waitForTimeout(300);
  await tab.close();
  tab = await open("?view=katalog");
  check("user CSV persists", (await tab.textContent("#catalog-status")).includes("1 şahıs · sizin seçtiğiniz dosya: benim.csv"), await tab.textContent("#catalog-status"));
  await tab.click("#reset-catalog");
  await tab.waitForTimeout(300);
  check("reset returns to the published catalog", (await tab.textContent("#catalog-status")).includes("3 şahıs · eklentiyle gelen katalog"), await tab.textContent("#catalog-status"));
  await tab.close();

  // Without a published catalog the window asks for a CSV.
  serveCatalog = false;
  const fresh = await browser.newContext({ viewport: { width: 360, height: 900 } });
  const bare = await fresh.newPage();
  await bare.route("https://appsforoffice.microsoft.com/**", (route) => route.fulfill({ contentType: "text/javascript", body: FAKE_OFFICE }));
  await bare.goto(`${base}taskpane.html`);
  await bare.waitForTimeout(600);
  check("without a catalog the window asks for a CSV", await bare.isVisible("#no-catalog"));
  await fresh.close();
  check("no page errors in other windows", errors.length === 0, errors.join("; "));
} finally {
  await browser.close();
  server.close();
}

for (const c of checks) console.log(`${c.ok ? "PASS" : "FAIL"}  ${c.label}${c.ok || !c.detail ? "" : `  -> ${c.detail}`}`);
const failed = checks.filter((c) => !c.ok).length;
console.log(`${checks.length - failed}/${checks.length} smoke checks passed`);
process.exit(failed ? 1 : 0);
