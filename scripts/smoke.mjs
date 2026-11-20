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
  "Pezdevî, Ebü’l-Usr;Ebü’l-Hasen Ebü’l-Usr Fahrü’l-İslâm Alî b. Muhammed b. el-Hüseyn b. Abdilkerîm el-Pezdevî;;(ö. 482/1089);Hanefî fakihi.;https://islamansiklopedisi.org.tr/pezdevi-ebul-usr",
  "Pezdevî, Ebü’l-Yüsr;Sadrü’l-İslâm Ebü’l-Yüsr Muhammed b. Muhammed b. el-Hüseyn b. Abdilkerîm el-Pezdevî;;(ö. 493/1100);Mâtürîdî kelâmcısı.;https://islamansiklopedisi.org.tr/pezdevi-ebul-yusr",
].join("\r\n");

const types = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".png": "image/png", ".svg": "image/svg+xml", ".json": "application/json", ".map": "application/json" };
let serveCatalog = true;
const server = createServer(async (req, res) => {
  try {
    // Ek kataloglar: one small field catalog (a DİA person in it must give way to DİA).
    if (req.url.startsWith("/catalogs.json")) {
      res.writeHead(200, { "content-type": "application/json" });
      return res.end(JSON.stringify([{ id: "deneme", title: "Deneme alanı", file: "catalogs/deneme.csv", count: 2 }]));
    }
    if (req.url.startsWith("/catalogs/deneme.csv")) {
      res.writeHead(200, { "content-type": "text/csv; charset=utf-8" });
      return res.end(["Kısa Ad;Tam Ad;Orijinal Ad;Vefat Tarihi;Tanıtım;Kaynak Link", "Émile Durkheim;---;---;(1858-1917);Fransız sosyolog;https://www.wikidata.org/wiki/Q15948", "Gazzâlî;---;---;(ö. 505/1111);Wikidata kopyası;https://www.wikidata.org/wiki/Q9546", "Ahmed Kemal;---;---;(1840-1900);Türk şair;https://www.wikidata.org/wiki/Q900", "Ahmed Kemal;---;---;(1847-1905);Türk tarihçi;https://www.wikidata.org/wiki/Q901", "Ahmed Kemal;---;---;(1854-1910);Türk iktisatçı;https://www.wikidata.org/wiki/Q902", "Ahmed Kemal;---;---;(1861-1915);Türk hukukçu;https://www.wikidata.org/wiki/Q903", "Ahmed Kemal;---;---;(1868-1920);Türk ressam;https://www.wikidata.org/wiki/Q904", "Ahmed Kemal;---;---;(1875-1925);Türk besteci;https://www.wikidata.org/wiki/Q905", "Ahmed Kemal;---;---;(1882-1930);Türk gazeteci;https://www.wikidata.org/wiki/Q906", "Ahmed Kemal;---;---;(1889-1935);Türk subay;https://www.wikidata.org/wiki/Q907", "Ahmed Kemal;---;---;(1896-1940);Türk hekim;https://www.wikidata.org/wiki/Q908"].join("\r\n"));
    }
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
  const doc = (window.__doc = Object.assign({ text: "", sel: [0, 0], controls: [], settings: {}, fields: [], api: [], notes: [] }, window.__docInit || {}));
  try { Object.assign(doc.settings, JSON.parse(sessionStorage.getItem("__fakeSettings") || "{}")); } catch {}
  // Stories: the main text is doc itself; every footnote or endnote is its own story.
  for (const note of doc.notes) { note.controls = note.controls || []; note.fields = note.fields || []; }
  doc.selStory = doc;
  const shift = (st, at, by) => {
    st.controls.forEach((c) => { if (c.range.s >= at) { c.range.s += by; c.range.e += by; } });
    st.fields.forEach((f) => { if (f.s >= at) { f.s += by; f.e += by; } });
  };
  const paraOf = (offset) => doc.text.slice(0, offset).split("\r").length - 1;
  const noteItem = (note) => {
    const at = doc.text.split("\r").slice(0, note.para + 1).join("\r").length;
    return { body: bodyOf(note), get reference() { return new R(at, at, doc); }, load() { return this; } };
  };
  const notesIn = (kind, from, to) => ({ items: doc.notes.filter((n) => n.kind === kind && n.para >= from && n.para <= to).map(noteItem), load() { return this; } });
  const bodyOf = (st) => ({
    get text() { return st.text; },
    search: (needle, options) => new R(0, st.text.length, st).search(needle, options),
    getRange: (loc) => new R(0, st.text.length, st).getRange(loc),
    get contentControls() { return { items: st.controls.filter((c) => c.range.s >= 0), load() { return this; } }; },
    get fields() { return { items: [], load() { return this; } }; },
    type: st === doc ? "MainDoc" : st.kind === "end" ? "Endnote" : "Footnote",
    load() { return this; },
  });
  window.__R = (a, b) => new R(a, b);
  class R {
    constructor(s, e, st = doc) { this.s = s; this.e = e; this.st = st; }
    load() { return this; }
    get text() { return this.st.text.slice(this.s, this.e); }
    get paragraphs() {
      const t = this.st.text;
      const end = t.indexOf("\r", this.e);
      const start = t.lastIndexOf("\r", this.s - 1) + 1;
      const whole = () => new R(start, end < 0 ? t.length : end, this.st);
      return { getLast: whole, getFirst: whole };
    }
    get footnotes() { return this.st === doc ? notesIn("foot", paraOf(this.s), paraOf(this.e)) : { items: [], load() { return this; } }; }
    get endnotes() { return this.st === doc ? notesIn("end", paraOf(this.s), paraOf(this.e)) : { items: [], load() { return this; } }; }
    get parentBody() { return bodyOf(this.st); }
    getRange(loc) {
      if (loc === "End" || loc === "After") return new R(this.e, this.e, this.st);
      if (loc === "Start" || loc === "Before") return new R(this.s, this.s, this.st);
      return new R(this.s, this.e, this.st);
    }
    expandTo(r) { return new R(Math.min(this.s, r.s), Math.max(this.e, r.e), this.st); }
    insertText(t, loc) {
      const st = this.st;
      if (loc === "After") { st.text = st.text.slice(0, this.e) + t + st.text.slice(this.e); shift(st, this.e, t.length); return new R(this.e, this.e + t.length, st); }
      if (loc === "Replace") { st.text = st.text.slice(0, this.s) + t + st.text.slice(this.e); shift(st, this.e, t.length - (this.e - this.s)); return new R(this.s, this.s + t.length, st); }
      throw new Error("unsupported insert " + loc);
    }
    insertContentControl() {
      const range = this;
      const st = this.st;
      const control = { load() { return this; }, id: doc.controls.length + doc.notes.reduce((n, x) => n + x.controls.length, 0) + 1, tag: "", title: "", appearance: "", range,
        getRange(loc) { return loc === "After" ? new R(this.range.e, this.range.e, st) : loc === "Before" ? new R(this.range.s, this.range.s, st) : this.range; },
        get paragraphs() { return new R(this.range.s, this.range.e, st).paragraphs; },
        get text() { return st.text.slice(this.range.s, this.range.e); },
        insertText(t) { this.range = this.range.insertText(t, "Replace"); },
        delete(keep) { if (!keep) new R(this.range.s, this.range.e, st).delete(); this.range = new R(-1, -1, st); },
        select() { doc.sel = [this.range.s, this.range.e]; doc.selStory = st; } };
      st.controls.push(control);
      return control;
    }
    get contentControls() { return { items: this.st.controls.filter((c) => c.range.s >= this.s && c.range.e <= this.e && c.range.s >= 0), load() { return this; } }; }
    get parentContentControlOrNullObject() {
      const c = this.st.controls.find((c) => c.range.s <= this.s && this.e <= c.range.e && c.range.e > c.range.s);
      return c ? Object.assign(Object.create(c), { isNullObject: false, load() { return this; } }) : { isNullObject: true, load() { return this; } };
    }
    delete() { const st = this.st; st.text = st.text.slice(0, this.s) + st.text.slice(this.e); shift(st, this.e, -(this.e - this.s)); st.controls.forEach((c) => { if (c.range.s >= this.s && c.range.e <= this.e) { c.range.s = c.range.e = -1; } }); }
    select(mode) { doc.sel = mode === "Start" ? [this.s, this.s] : [this.s, this.e]; doc.selStory = this.st; }
    search(needle) {
      const items = [];
      const n = needle.replace(/\^\^/g, "^");
      const hay = this.st.text.slice(this.s, this.e);
      for (let i = hay.indexOf(n); i >= 0; i = hay.indexOf(n, i + 1)) items.push(new R(this.s + i, this.s + i + n.length, this.st));
      return { items, load() { return this; } };
    }
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
    run: async (fn) => {
      doc.runs = (doc.runs || 0) + 1;
      const items = [];
      let at = 0;
      for (const part of doc.text.split("\r")) {
        const range = new R(at, at + part.length);
        items.push({ text: part, uniqueLocalId: "p" + items.length, getRange: () => range, load() { return this; } });
        at += part.length + 1;
      }
      const paragraphs = { items, load() { return this; } };
      const sections = { items: [{ getHeader: () => ({ text: "", load() { return this; } }), getFooter: () => ({ text: "", load() { return this; } }) }], load() { return this; } };
      const fields = { get items() {
        return doc.fields.filter((f) => !f.gone).map((f) => ({ code: f.code, get result() { return new R(f.s, f.e); },
          delete() { doc.text = doc.text.slice(0, f.s) + doc.text.slice(f.e); const by = f.e - f.s, at = f.s; f.gone = true; shift(doc, at + 1, -by); } }));
      }, load() { return this; } };
      doc.handlers = doc.handlers || { changed: [], added: [], deleted: [] };
      const events = (kind) => ({ add: (h) => { doc.handlers[kind].push(h); } });
      return fn({ document: { sections, onParagraphChanged: events("changed"), onParagraphAdded: events("added"), onParagraphDeleted: events("deleted"),
        getParagraphByUniqueLocalId: (id) => items[Number(id.slice(1))], getSelection: () => new R(doc.sel[0], doc.sel[1], doc.selStory || doc), body: { search, paragraphs, fields, get footnotes() { return notesIn("foot", 0, 1e9); }, get endnotes() { return notesIn("end", 0, 1e9); }, get contentControls() { return { items: doc.controls.filter((c) => c.range.s >= 0), load() { return this; } }; } } }, sync: async () => {} });
    },
  };
  window.Office = {
    HostType: { Word: "Word" },
    addin: { showAsTaskpane: async () => { window.__shown = (window.__shown || 0) + 1; } },
    actions: doc.api.includes("KeyboardShortcuts 1.1") ? {
      associate: () => {},
      getShortcuts: async () => ({ go_action_insert: null, go_action_spelling: "Alt+Shift+Y", ...(window.__keys || {}) }),
      areShortcutsInUse: async (keys) => keys.map((k) => ({ shortcut: k, inUse: k === "Alt+Shift+A" })),
      replaceShortcuts: async (m) => { window.__keys = { ...(window.__keys || {}), ...m }; },
    } : undefined,
    EventType: { DocumentSelectionChanged: "selection" },
    AsyncResultStatus: { Succeeded: "succeeded" },
    onReady: (cb) => setTimeout(() => cb({ host: "Word" }), 0),
    context: {
      requirements: { isSetSupported: (name, version) => doc.api.includes(name + " " + version) },
      document: {
        settings: {
          get: (k) => doc.settings[k] ?? null,
          set: (k, v) => { doc.settings[k] = v; },
          remove: (k) => { delete doc.settings[k]; },
          saveAsync: (cb) => { try { sessionStorage.setItem("__fakeSettings", JSON.stringify(doc.settings)); } catch {} cb({ status: "succeeded" }); },
          refreshAsync: (cb) => cb({ status: "succeeded" }),
        },
        addHandlerAsync: (_type, handler) => { window.__select = (s, e) => { doc.sel = [s, e]; handler(); }; window.__fireSelection = () => handler(); },
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
      fields: options.fields ?? [],
      notes: options.notes ?? [],
      api: options.api ?? [],
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
  // Lakap + nisba, written joined: "Fahrülislâm el-Pezdevî" is in the catalog as "Fahrü’l-İslâm … el-Pezdevî".
  const lakapTab = await open("", "Fahrülislâm el-Pezdevî bu meseleyi ele alır.", "Fahrülislâm el-Pezdevî");
  const lakap = (await lakapTab.textContent("#view-secili")).replace(/\s+/g, " ");
  check("joined lakap + nisba is found and the right bearer chosen", lakap.includes("Ebü’l-Usr") && lakap.includes("482/1089") && !(await lakapTab.isDisabled("#insert")), lakap.slice(0, 300));
  await lakapTab.close();
  // Many bearers of one name: five shown, the rest behind a link, with a search box.
  const manyTab = await open("", "Ahmed Kemal bu yazıyı yazdı.", "Ahmed Kemal");
  const firstFive = (await manyTab.$$("#candidates .candidate")).length;
  check("long candidate list shortened to five", firstFive === 5 && (await manyTab.textContent("#candidates .more-candidates")).includes("Daha fazla göster (4)"), String(firstFive));
  await manyTab.click("#candidates .more-candidates");
  await manyTab.fill("#candidate-filter", "hekim");
  const visible = await manyTab.$$eval("#candidates .candidate", (n) => n.filter((x) => !x.hidden).map((x) => x.textContent));
  check("candidate search narrows the open list", visible.length === 1 && visible[0].includes("hekim"), visible.join(" | "));
  await manyTab.fill("#candidate-filter", "yokböyle");
  check("candidate search says when nothing fits", await manyTab.isVisible("#candidates .filter-empty"));
  await manyTab.close();
  // Wikidata'da ara: a name no catalog has is looked up on Wikidata and added as a record of one's own.
  const wdTab = await open("", "Hasan et-Türâbî bunu savundu.", "Hasan et-Türâbî");
  await wdTab.route("https://www.wikidata.org/**", (route) => route.fulfill({ contentType: "application/json", headers: { "access-control-allow-origin": "*" }, body: JSON.stringify({ search: [{ id: "Q469299" }] }) }));
  await wdTab.route("https://query.wikidata.org/**", (route) =>
    route.fulfill({
      contentType: "application/sparql-results+json",
      headers: { "access-control-allow-origin": "*" },
      body: JSON.stringify({ results: { bindings: [{ p: { value: "http://www.wikidata.org/entity/Q469299" }, links: { value: "34" }, death: { value: "2016-03-05T00:00:00Z" }, dprec: { value: "11" }, birth: { value: "1932-02-01T00:00:00Z" }, bprec: { value: "11" }, tr: { value: "Hasan el-Turabi" }, en: { value: "Hasan al-Turabi" }, enDesc: { value: "Sudanese politician" }, c: { value: "http://www.wikidata.org/entity/Q1049" } }] } }),
    }),
  );
  check("a name no catalog has offers Wikidata'da ara", (await wdTab.textContent("#candidates")).includes("Katalogda bu adla eşleşen şahıs yok") && (await wdTab.isVisible("#wikidata-search")));
  await wdTab.click("#wikidata-search");
  await wdTab.waitForSelector("#wikidata-results .wikidata-hit");
  check("Wikidata hits listed in DİA's spelling, with their dates", (await wdTab.textContent("#wikidata-results")).includes("Hasan et-Turabi (1932-2016)") && (await wdTab.textContent("#wikidata-results")).includes("Wikidata: Hasan el-Turabi"), await wdTab.textContent("#wikidata-results"));
  await wdTab.click("#wikidata-results .wikidata-hit button");
  await wdTab.waitForTimeout(300);
  check("the added Wikidata person is chosen for the selection", (await wdTab.textContent("#candidates")).includes("Hasan et-Turabi") && (await wdTab.textContent("#candidates")).includes("kendi kaydınız") && !(await wdTab.isDisabled("#insert")), await wdTab.textContent("#candidates"));
  await wdTab.click("#insert");
  await wdTab.waitForTimeout(200);
  check("its date goes into the text", (await docText(wdTab)).includes("Hasan et-Türâbî (1932-2016)"), await docText(wdTab));
  await wdTab.close();
  const keyTab = await open("", "Ahmed Kemal bu yazıyı yazdı.", "Ahmed Kemal");
  await keyTab.focus("#candidate-0");
  await keyTab.keyboard.press("ArrowDown");
  await keyTab.keyboard.press("ArrowDown");
  const walked = await keyTab.evaluate(() => ({ rows: document.querySelectorAll("#candidates .candidate").length, focus: document.activeElement?.id, chosen: document.querySelector("#candidates .chosen")?.querySelector("input")?.id }));
  check("arrow keys walk the candidates without shortening the list", walked.rows === 5 && walked.focus === "candidate-2" && walked.chosen === "candidate-2", JSON.stringify(walked));
  await keyTab.keyboard.press("Enter");
  await keyTab.waitForTimeout(200);
  check("Enter on a candidate adds the date", (await docText(keyTab)).includes("Ahmed Kemal (ö. 1854-1910)") || (await docText(keyTab)).includes("Ahmed Kemal (1854-1910)"), await docText(keyTab));
  await keyTab.close();
  // İsim yazımı with a double-click selection (the name and the space after it).
  const spellTab = await open("", "Bu konuda Gazali'nin eseri önemlidir.", "Gazali'nin ");
  await spellTab.check("#candidate-0");
  await spellTab.selectOption("#spelling", "Gazzâlî");
  await spellTab.click("#respell");
  await spellTab.waitForTimeout(300);
  check("İsim yazımı works when the space after the name is selected too", (await docText(spellTab)) === "Bu konuda Gazzâlî'nin eseri önemlidir.", await docText(spellTab));
  await spellTab.close();
  check("settings summary at the top", (await page.textContent("#settings-summary")) === "Kısaltma: ö. · Yalnız ilk kullanım · Ad biçimi: Her seferinde sor");

  await select(page, "Gazzâlî");
  check("single strong candidate chosen", await page.isEnabled("#insert"));
  await page.click("#insert");
  await page.waitForTimeout(200);
  check("date inserted at first mention", (await docText(page)).startsWith("Gazzâlî (ö. 505/1111) bu görüşü"), await docText(page));
  check("hidden tagged control", await page.evaluate(() => window.__doc.controls[0]?.appearance === "Hidden" && /^vefeyat:date:[0-9a-f]{8}$/.test(window.__doc.controls[0].tag)));

  await select(page, "Gazzâlî", 1);
  await page.click("#insert");
  await page.waitForTimeout(300);
  check("a person who already has a date gets no second one", (await page.isVisible("#dated")) && (await page.isHidden("#earlier")) && (await page.isHidden("#move-date")) && (await page.textContent("#dated-text")).includes("(ö. 505/1111)"), await page.textContent("#dated-text"));
  check("nothing written for the second mention", !(await docText(page)).includes("Gazzâlî (ö. 505/1111) der"), await docText(page));
  await page.click("#goto-dated");
  await page.waitForTimeout(200);
  check("Tarihe git selects the dated mention", (await page.evaluate(() => window.__doc.sel[0])) === 0 && (await page.isHidden("#dated")));

  await select(page, "Gazzâlî", 0);
  await page.click("#insert");
  await page.waitForTimeout(200);
  check("existing date is not duplicated", (await page.textContent("#message")).includes("zaten katalogdakiyle aynı") && (await docText(page)).split("(ö. 505/1111)").length === 2, await docText(page));

  await select(page, "Gazali'nin");
  const weak = (await page.textContent("#view-secili")).replace(/\s+/g, " ");
  check("weak spelling is not chosen automatically (at most suggested from the document, to verify)", !(await page.isVisible("#details")) || weak.includes("doğrulayın"), weak.slice(0, 300));
  await page.check("#candidate-0");
  const options = await page.$$eval("#spelling option", (o) => o.map((x) => x.value));
  check("spelling suggestion offered", options[0] === "" && options[1] === "Gazzâlî", options.join(" | "));
  check("current spelling is what is selected; nothing to change yet", (await page.isDisabled("#respell")) && (await page.$eval("#spelling", (e) => e.selectedOptions[0].textContent)).startsWith("Gazali'nin"));
  check("chosen candidate is marked", (await page.textContent(".candidate.chosen")).includes("seçili"));
  await page.selectOption("#spelling", "Gazzâlî");
  check("picking a spelling shows what will be written", (await page.textContent("#spelling-preview")).includes("Gazali'nin → Gazzâlî'nin") && !(await page.isDisabled("#respell")));
  await page.click("#respell");
  await page.waitForTimeout(300);
  check("spelling replaced with ending kept", (await docText(page)).includes("Gazzâlî'nin eseri"), await docText(page));

  // Ribbon Tercihler menus write document settings through the function file.
  const commands = await open("", "", "", { page: "commands.html" });
  await commands.waitForFunction(() => typeof window.set_prefix_v === "function");
  const done = await commands.evaluate(() => new Promise((resolve) => window.set_prefix_v({ completed: () => resolve(window.__doc.settings["vefeyat.prefix"]) })));
  check("Kısaltma menu stores the setting", done === "v");
  check("all 19 setting commands registered", (await commands.evaluate(() => Object.keys(window).filter((k) => k.startsWith("set_")).length)) === 19);
  await commands.close();

  await page.evaluate(() => { window.__doc.settings["vefeyat.prefix"] = "v"; window.__doc.settings["vefeyat.nameForm"] = "full"; });
  await select(page, "Râzî");
  check("window shows changed settings", (await page.textContent("#settings-summary")).includes("Kısaltma: v.") && (await page.textContent("#settings-summary")).includes("Tam / uzun ad"));
  const candidates = await page.$$eval("#candidates .candidate strong", (n) => n.map((x) => x.textContent));
  check("homonyms listed, well-known first", candidates[0]?.startsWith("Fahreddin er-Râzî") && candidates.length === 2, candidates.join(" | "));
  check("choice required for homonyms", !(await page.isEnabled("#insert")));
  await page.check("#candidate-0");
  check("chosen candidate shown alone, others behind a link", (await page.$$("#candidates .candidate")).length === 1 && (await page.textContent("#candidates .more-candidates")).includes("Diğer adayları göster (1)"));
  await page.click("#candidates .more-candidates");
  check("other candidates open on request", (await page.$$("#candidates .candidate")).length === 2);
  await page.click("#candidates .more-candidates");
  check("Ad biçimi puts the full name first", (await page.$$eval("#spelling option", (o) => o[1]?.textContent ?? "")).includes("Tam / uzun ad"));
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

  // A pane opened from Word's cache reloads itself once when a newer build is published.
  {
    const page = await context.newPage();
    page.on("pageerror", (error) => errors.push(error.message));
    await page.route("https://appsforoffice.microsoft.com/**", (route) => route.fulfill({ contentType: "text/javascript", body: FAKE_OFFICE }));
    await page.route("**/version.json*", (route) => route.fulfill({ contentType: "application/json", body: JSON.stringify({ version: "9.9.9.9" }) }));
    await page.goto(`${base}taskpane.html?view=yardim`);
    await page.waitForURL(/_v=9\.9\.9\.9/, { timeout: 5000 });
    await page.waitForTimeout(800);
    const url = page.url();
    check("stale pane reloads once with the new version", url.includes("_v=9.9.9.9") && url.includes("view=yardim") && (await page.textContent("#message")).includes("eski kopyayı"), url);
    await page.close();
  }
  // Kendi şahıslarım: a name the catalog lacks is added from the selected-name window and then dated.
  tab = await open("", "Asım Köksal bu konuyu inceler.", "Asım Köksal");
  await tab.waitForSelector("#candidates button");
  await tab.click("#candidates button");
  await tab.waitForSelector("#view-katalog:not([hidden])");
  const prefilled = await tab.inputValue("#local-short");
  await tab.fill("#local-date", "1913-1998");
  await tab.fill("#local-description", "Tarihçi");
  await tab.click("#local-add");
  await tab.waitForFunction(() => (window.__doc.settings["vefeyat.persons"] || "").includes("Asım Köksal"));
  check("own person added from the selected-name window", prefilled === "Asım Köksal" && (await tab.textContent("#locals-list")).includes("(1913-1998)"));
  await tab.evaluate(() => history.replaceState(null, "", "?action=insert"));
  await tab.reload();
  await tab.waitForFunction(() => window.__doc.text.includes("1998"), null, { timeout: 5000 }).catch(() => {});
  check("own person is dated like a catalog entry", (await docText(tab)) === "Asım Köksal (1913-1998) bu konuyu inceler.", await docText(tab));
  await tab.close();
  // Ribbon buttons.
  tab = await open("?action=insert", "İlk kez Gazzâlî burada.", "Gazzâlî");
  check("ribbon insert writes at once when unambiguous", (await docText(tab)) === "İlk kez Gazzâlî (ö. 505/1111) burada.", await docText(tab));
  await tab.close();
  // Keyboard-shortcut manifest: buttons and the shortcut are functions in the open pane.
  tab = await open("", "Giriş. İlk kez Gazzâlî burada.");
  await select(tab, "Gazzâlî");
  const completed = await tab.evaluate(() => new Promise((resolve) => window.go_action_insert({ completed: () => resolve(true) })));
  await tab.waitForTimeout(500);
  check("shortcut function shows the pane and inserts the date", completed && (await tab.evaluate(() => window.__shown)) === 1 && (await docText(tab)) === "Giriş. İlk kez Gazzâlî (ö. 505/1111) burada.", await docText(tab));
  await tab.evaluate(() => window.go_view_dokum(undefined));
  await tab.waitForTimeout(500);
  check("ribbon function switches the open pane to Şahıs dökümü", await tab.isVisible("#view-dokum"));
  await tab.evaluate(() => new Promise((resolve) => window.set_prefix_v({ completed: resolve })));
  check("Tercihler functions also live in the pane", (await tab.evaluate(() => window.__doc.settings["vefeyat.prefix"])) === "v");
  await tab.waitForSelector("#prefix-offer:not([hidden])", { timeout: 3000 }).catch(() => {});
  check("Kısaltma change offers to convert the document's dates", (await tab.textContent("#prefix-offer-text")).includes("1 Vefeyât tarihi"), await tab.textContent("#prefix-offer-text"));
  await tab.click("#prefix-convert");
  await tab.waitForTimeout(400);
  check("dates converted to the new Kısaltma", (await docText(tab)) === "Giriş. İlk kez Gazzâlî (v. 505/1111) burada.", await docText(tab));
  await tab.close();
  // A date already after the name: Güncelle.
  tab = await open("", "Gazzâlî (ö. 505/1112) der.", "Gazzâlî");
  await tab.click("#insert");
  await tab.waitForTimeout(300);
  check("existing different date offers Güncelle", await tab.isVisible("#existing") && (await tab.textContent("#existing-text")).includes("(ö. 505/1111)"));
  await tab.click("#update-date");
  await tab.waitForTimeout(300);
  check("Güncelle replaces the written date with a Vefeyât date", (await docText(tab)) === "Gazzâlî (ö. 505/1111) der." && (await tab.evaluate(() => window.__doc.controls.filter((c) => c.range.e > c.range.s && c.tag.startsWith("vefeyat:date:")).length)) === 1, await docText(tab));
  await tab.evaluate(() => { window.__doc.settings["vefeyat.prefix"] = "v"; });
  await select(tab, "Gazzâlî");
  await tab.click("#insert");
  await tab.waitForTimeout(300);
  await tab.click("#update-date");
  await tab.waitForTimeout(300);
  check("Güncelle on a Vefeyât date replaces it with its control", (await docText(tab)) === "Gazzâlî (v. 505/1111) der." && (await tab.evaluate(() => window.__doc.controls.filter((c) => c.range.e > c.range.s).length)) === 1, await docText(tab));
  await select(tab, "Gazzâlî");
  await tab.click("#insert");
  await tab.waitForTimeout(300);
  await tab.close();
  // A long thesis: Şahıs dökümü must not freeze the window (work is handed back to the browser).
  {
    const names = ["Gazzâlî", "Fahreddin er-Râzî", "İbn Sînâ", "Cüveynî", "Şâfiî", "Zemahşerî"];
    const filler = "Fıkıh usûlünün mahiyetine ilişkin görüşler, usûlcülerin bu ilim için yaptıkları tariflerden de anlaşılabilir.";
    const big = Array.from({ length: 3000 }, (_, k) => `${filler} ${names[k % names.length]} bu konuda şöyle der${k % 3 ? "" : " (1993: 45)"}. ${filler}`).join("\r");
    tab = await open("", big);
    await tab.waitForTimeout(1500);
    await tab.evaluate(() => {
      window.__longest = 0;
      let last = performance.now();
      window.__beat = setInterval(() => { const now = performance.now(); window.__longest = Math.max(window.__longest, now - last); last = now; }, 10);
    });
    const t0 = Date.now();
    await tab.evaluate(() => window.go_view_dokum(undefined));
    await tab.waitForFunction(() => /\d+ geçiş/.test(document.getElementById("inventory-summary")?.textContent ?? ""), null, { timeout: 60000 });
    const took = Date.now() - t0;
    const longest = await tab.evaluate(() => { clearInterval(window.__beat); return Math.round(window.__longest); });
    console.log(`    3000 paragraf: döküm ${took} ms, en uzun donma ${longest} ms`);
    check("long thesis: inventory keeps the window responsive", longest < 1000, `${longest} ms`);
    await tab.close();
  }
  // The date stands on a later mention: it moves to the selected, earlier one.
  tab = await open("", "Önce Gazzâlî der.\rSonra Gazzâlî (ö. 505/1111) yazar.", "Gazzâlî");
  await tab.click("#insert");
  await tab.waitForTimeout(400);
  check("an earlier mention offers to move the date here", (await tab.isVisible("#dated")) && (await tab.isVisible("#move-date")), await tab.textContent("#dated-text"));
  await tab.click("#move-date");
  await tab.waitForTimeout(400);
  check("Tarihi buraya taşı: one date, at the first mention", (await docText(tab)) === "Önce Gazzâlî (ö. 505/1111) der.\rSonra Gazzâlî yazar.", await docText(tab));
  await tab.close();
  // Şahıs dökümü süzgeci and name search.
  tab = await open("?view=dokum", "Gazzâlî der.\rSonra Gazali'nin eseri.\rFahreddin er-Râzî (ö. 606/1210) de.");
  await tab.waitForFunction(() => document.querySelectorAll("#inventory-list li").length === 3);
  const filterLabels = await tab.$$eval("#inventory-filter option", (o) => o.map((x) => x.textContent));
  await tab.selectOption("#inventory-filter", "undated");
  const undated = await tab.$$eval("#inventory-list li strong", (n) => n.map((x) => x.textContent));
  await tab.selectOption("#inventory-filter", "all");
  await tab.fill("#inventory-search", "gazali");
  await tab.waitForTimeout(400);
  const searched = await tab.$$eval("#inventory-list li strong", (n) => n.map((x) => x.textContent));
  check("Şahıs dökümü filters by status and searches by name", filterLabels[0] === "Tümü (3)" && undated.join() === "Gazzâlî" && searched.join() === "Gazzâlî,Gazali'nin", `${filterLabels} | ${undated} | ${searched}`);
  await tab.close();
  // One choice for every uncertain mention of the same spelling.
  tab = await open("", "Râzî der.\rSonra Râzî'nin eseri.\rRâzî yine der.", "Râzî");
  check("apply-to-all appears once a bearer is chosen among several", (await tab.isHidden("#apply-all")));
  await tab.check("#candidate-1");
  await tab.click("#apply-all");
  await tab.waitForTimeout(500);
  check("apply-to-all confirms every uncertain mention of the spelling", (await tab.textContent("#message")).includes("3 belirsiz geçişi") && Object.keys(JSON.parse(await tab.evaluate(() => window.__doc.settings["vefeyat.confirmations"] ?? "{}"))).length === 3, await tab.textContent("#message"));
  await tab.evaluate(() => window.go_view_dokum(undefined));
  await tab.waitForFunction(() => /geçiş/.test(document.getElementById("inventory-summary")?.textContent ?? ""));
  check("the inventory takes the confirmations", (await tab.textContent("#inventory-summary")).includes("0 doğrulanacak"), await tab.textContent("#inventory-summary"));
  await tab.close();
  // İsim yazımı → Tümünü değiştir: every mention of this spelling of this person, endings kept, in notes too.
  tab = await open("", "Gazzâlî der.\rSonra Gazali'nin eseri.\rGazali yine der; Gazali'ye göre.\rKAYNAKÇA\rGazali, Ebû Hâmid. el-Mustasfâ. Beyrut 1993.", "Gazali'nin", { api: ["WordApi 1.5"], notes: [{ kind: "foot", para: 1, text: "Gazali, el-Mustasfâ'da da böyle der." }] });
  await tab.waitForTimeout(300);
  if (await tab.isHidden("#details")) await tab.check("#candidate-0");
  await tab.selectOption("#spelling", "Gazzâlî");
  await tab.click("#respell-all");
  await tab.waitForSelector("#respell-confirm:not([hidden])", { timeout: 5000 }).catch(() => {});
  const plan = await tab.textContent("#respell-confirm-text");
  check("Tümünü değiştir says how many mentions change before doing it", plan.includes("belgede 5 geçiş değişecek (1 tanesi dipnot/sonnotta, 1 tanesi atıf ya da kaynakçada)") && (await docText(tab)).includes("Gazali'nin"), plan);
  await tab.click("#respell-all-go");
  await tab.waitForTimeout(500);
  const notesText = await tab.evaluate(() => window.__doc.notes.map((n) => n.text).join(" | "));
  check("Tümünü değiştir changes every mention, endings kept, citations in notes too", (await docText(tab)) === "Gazzâlî der.\rSonra Gazzâlî'nin eseri.\rGazzâlî yine der; Gazzâlî'ye göre.\rKAYNAKÇA\rGazzâlî, Ebû Hâmid. el-Mustasfâ. Beyrut 1993." && notesText.includes("Gazzâlî, el-Mustasfâ"), `${await docText(tab)} || ${notesText} || ${await tab.textContent("#message")}`);
  await tab.close();
  // Sorun bildir: versions, settings and recent messages; never the document text.
  tab = await open("?view=yardim", "Gizli tez metni: Gazzâlî der.", "Gazzâlî");
  await tab.fill("#report-note", "Tarih eklenmedi.");
  await tab.waitForTimeout(200);
  const report = await tab.inputValue("#report-text");
  check("Sorun bildir lists version, settings and the note, not the document text", report.includes("Ne oldu: Tarih eklendi".slice(0, 10)) && report.includes("Vefeyât: ") && report.includes("Ayarlar: Kısaltma") && report.includes("Tarih eklenmedi.") && !report.includes("Gizli tez metni"), report.slice(0, 300));
  await tab.click("#report-mail").catch(() => {});
  check("e-mail link goes to the beta address", (await tab.getAttribute("#report-mail", "href")).startsWith("mailto:apoyraz@cumhuriyet.edu.tr?subject="));
  await tab.close();
  // A title before the name: found without it, kept when the spelling changes.
  tab = await open("", "Gazzâlî der.\rSonra İmam Gazali'nin eseri.", "İmam Gazali'nin");
  await tab.waitForTimeout(300);
  if (await tab.isHidden("#details")) await tab.check("#candidate-0");
  await tab.selectOption("#spelling", "Gazzâlî");
  check("a title before the name: the preview keeps it", (await tab.textContent("#spelling-preview")).includes("İmam Gazali'nin → İmam Gazzâlî'nin"), await tab.textContent("#spelling-preview"));
  await tab.click("#respell");
  await tab.waitForTimeout(300);
  check("Değiştir keeps the title", (await docText(tab)) === "Gazzâlî der.\rSonra İmam Gazzâlî'nin eseri.", await docText(tab));
  await tab.close();
  // Ek kataloglar: Durkheim comes from the field catalog; Gazzâlî stays DİA's.
  tab = await open("?view=katalog", "Durkheim ve Gazzâlî der.", "");
  await tab.waitForSelector("#extras:not([hidden])", { timeout: 3000 }).catch(() => {});
  check("extra catalogs are on from the first run", (await tab.isChecked("#extras-list input")) && (await tab.textContent("#catalog-status")).includes("+ 10 kişi ek kataloglardan"), await tab.textContent("#catalog-status"));
  await tab.uncheck("#extras-list input");
  await tab.waitForTimeout(600);
  const off = (await tab.textContent("#message")).includes("kataloğu çıkarıldı") && !(await tab.textContent("#catalog-status")).includes("ek kataloglardan");
  await tab.check("#extras-list input");
  await tab.waitForTimeout(600);
  check("an extra catalog can be turned off and on again", off && (await tab.textContent("#message")).includes("Deneme alanı kataloğu eklendi") && (await tab.textContent("#catalog-status")).includes("+ 10 kişi ek kataloglardan"), `${await tab.textContent("#message")} | ${await tab.textContent("#catalog-status")}`);
  await select(tab, "Durkheim");
  const durkheim = (await tab.textContent("#view-secili")).replace(/\s+/g, " ");
  check("a Wikidata person is found, with its source and link", durkheim.includes("Émile Durkheim") && durkheim.includes("Wikidata") && (await tab.textContent("#detail-link")) === "Wikidata kaydını aç", durkheim.slice(0, 200));
  await select(tab, "Gazzâlî");
  check("DİA is authoritative where the catalogs meet", (await tab.getAttribute("#detail-link", "href")).startsWith("https://islamansiklopedisi.org.tr/") && (await tab.$$("#candidates .candidate")).length === 1);
  await tab.close();
  // A double click selects the space after the name: the date still goes right after the name.
  tab = await open("", "Gazzâlî'nin farklı eserleri.", "Gazzâlî'nin ");
  await tab.click("#insert");
  await tab.waitForTimeout(300);
  check("date goes after the name, not after the selected space", (await docText(tab)) === "Gazzâlî'nin (ö. 505/1111) farklı eserleri.", await docText(tab));
  await tab.close();
  tab = await open("", "Gazzâlî'ninfarklı eserleri.", "Gazzâlî'nin");
  await tab.click("#insert");
  await tab.waitForTimeout(300);
  check("a space is kept between the date and the next word", (await docText(tab)) === "Gazzâlî'nin (ö. 505/1111) farklı eserleri.", await docText(tab));
  await tab.close();
  // Belge sağlığı: broken spacing around a Vefeyât date (old double-click bug).
  tab = await open("", "Gazzâlî'nin farklı eserleri.", "Gazzâlî'nin");
  await tab.click("#insert");
  await tab.waitForTimeout(300);
  await tab.evaluate(() => {
    // "Gazzâlî'nin (ö. 505/1111) farklı" → "Gazzâlî'nin  (ö. 505/1111)farklı"
    const c = window.__doc.controls.find((x) => x.range.e > x.range.s);
    window.__R(c.range.e, c.range.e + 1).delete();
    window.__R(c.range.s, c.range.s).insertText(" ", "Replace");
  });
  check("spacing broken as in the report", (await docText(tab)) === "Gazzâlî'nin  (ö. 505/1111)farklı eserleri.", await docText(tab));
  await tab.evaluate(() => window.go_view_saglik(undefined));
  await tab.waitForSelector("#spacing:not([hidden])", { timeout: 3000 }).catch(() => {});
  check("Belge sağlığı reports the broken spacing", (await tab.textContent("#spacing-text")).includes("1 tarihin çevresinde boşluk bozuk"), await tab.textContent("#spacing-text"));
  await tab.click("#spacing-fix");
  await tab.waitForTimeout(500);
  check("Boşlukları düzelt restores one space on each side", (await docText(tab)) === "Gazzâlî'nin (ö. 505/1111) farklı eserleri." && await tab.isHidden("#spacing"), await docText(tab));
  await tab.close();
  // The date deleted by hand leaves an empty control (Word shows its placeholder): removed right after the edit.
  tab = await open("", "Gazzâlî der.", "Gazzâlî", { api: ["WordApi 1.6"] });
  await tab.click("#insert");
  await tab.waitForTimeout(300);
  await tab.evaluate(() => {
    const c = window.__doc.controls.find((x) => x.range.e > x.range.s);
    window.__R(c.range.s, c.range.e).delete();
    c.range.s = c.range.e = window.__doc.text.indexOf(" der");
    for (const h of window.__doc.handlers.changed) h({ uniqueLocalIds: ["p0"] });
  });
  await tab.waitForTimeout(800);
  check("empty date control is removed right after the edit", (await tab.evaluate(() => window.__doc.controls.filter((c) => c.range.s >= 0).length)) === 0 && (await docText(tab)) === "Gazzâlî der.", await docText(tab));
  await tab.close();
  tab = await open("", "Gazzâlî (ö. 505/1111) der.", "Gazzâlî");
  await tab.evaluate(() => { window.__doc.settings["vefeyat.prefix"] = "o"; });
  await select(tab, "Gazzâlî");
  await tab.click("#insert");
  await tab.waitForTimeout(300);
  check("same date: says it is already current", (await tab.textContent("#message")).includes("zaten katalogdakiyle aynı") && await tab.isHidden("#existing"));
  await tab.close();
  tab = await open("?view=kisayol", "Metin.", "", { api: ["KeyboardShortcuts 1.1"] });
  await tab.waitForSelector("#shortcuts:not([hidden])", { timeout: 3000 }).catch(() => {});
  const keyOptions = await tab.$$eval("#shortcuts-list select[data-action=go_action_insert] option", (o) => o.map((x) => x.value));
  check("Klavye kısayolları lists conflicts and only free keys", (await tab.textContent("#shortcuts-summary")).includes("1 kısayol kullanılamıyor") && !keyOptions.includes("Alt+Shift+A") && !keyOptions.includes("Alt+Shift+Y") && keyOptions.includes("Alt+Shift+C"), keyOptions.slice(0, 5).join(","));
  check("another command's default key is never offered", !keyOptions.includes("Alt+Shift+B") && keyOptions.includes("Alt+Shift+C"), keyOptions.slice(0, 5).join(","));
  await tab.selectOption("#shortcuts-list select[data-action=go_action_insert]", "Alt+Shift+C");
  await tab.click("#shortcuts-save");
  await tab.waitForTimeout(300);
  check("chosen shortcut is saved in Word", (await tab.evaluate(() => window.__keys?.go_action_insert)) === "Alt+Shift+C");
  await tab.click("#shortcuts-reset");
  await tab.waitForTimeout(300);
  check("reset returns every command to its default", (await tab.evaluate(() => Object.values(window.__keys).every((v) => v === null) && Object.keys(window.__keys).length === 6)));
  await tab.close();
  tab = await open("?view=tercihler", "Metin.");
  check("no shortcut section without shortcut support", await tab.isHidden("#shortcuts"));
  await tab.close();
  tab = await open("?action=insert", "Râzî der ki.", "Râzî");
  check("ribbon insert asks for a choice with homonyms", (await tab.textContent("#message")).includes("Birden fazla aday") && (await docText(tab)) === "Râzî der ki.");
  await tab.close();
  tab = await open("?action=insert", "Gazzâlî önce. Sonra Gazzâlî.", "Gazzâlî", { last: true });
  check("ribbon insert still warns about an earlier mention", (await tab.isVisible("#earlier")) && (await docText(tab)) === "Gazzâlî önce. Sonra Gazzâlî.");
  await tab.close();
  // A date written after an ambiguous name picks its bearer in the selected-name window.
  tab = await open("", "Bu konuda Râzî (ö. 313/925) farklı düşünür.", "Râzî");
  await tab.waitForSelector(".candidate.chosen");
  check("written date picks and marks the homonym", (await tab.textContent(".candidate.chosen")).includes("Ebû Bekir") && (await tab.textContent(".candidate.chosen")).includes("metindeki ölüm tarihine göre"), await tab.textContent(".candidate.chosen"));
  await tab.close();
  // Another person with the same short name earlier is no earlier mention.
  tab = await open("?action=insert", "Fahreddin er-Râzî tefsir yazdı. Ebû Bekir er-Râzî tıpta meşhurdur.", "Ebû Bekir er-Râzî");
  check("an earlier homonym does not count as an earlier mention", !(await tab.isVisible("#earlier")) && (await docText(tab)) === "Fahreddin er-Râzî tefsir yazdı. Ebû Bekir er-Râzî (ö. 313/925) tıpta meşhurdur.", await docText(tab));
  await tab.close();
  tab = await open("?tool=dizin");
  check("Teslim menu opens the tools window", (await tab.isVisible("#view-teslim")) && (await tab.getAttribute("#tool-dizin", "class")) === "focus" && (await tab.textContent("#view-title")) === "Teslim araçları");
  await tab.click("#run-tutarlilik");
  await tab.waitForSelector("#consistency-out:not([hidden])");
  check("Tutarlılık raporu runs on the document", ((await tab.textContent("#consistency-out")) ?? "").length > 0);
  await tab.click("#back");
  check("back returns to the selected name", await tab.isVisible("#view-secili"));
  await tab.close();
  tab = await open("?view=ara");
  await tab.fill("#search-query", "Râzî");
  await tab.click("#search-form button");
  await tab.waitForTimeout(200);
  const hits = await tab.$$eval("#search-results li strong", (n) => n.map((x) => x.textContent));
  check("Şahıs ara lists catalog matches", hits.length === 2 && hits[0].startsWith("Fahreddin er-Râzî"), hits.join(" | "));
  await tab.close();
  tab = await open("?view=tercihler");
  check("Tercihler window opens", (await tab.textContent("#view-title")) === "Tercihler" && (await tab.isVisible("#view-tercihler")));
  await tab.fill("#prefs-text", JSON.stringify({ format: "vefeyat-tercihler", version: 1, settings: { prefix: "v" }, spellings: {}, exclusions: ["deneme"], confirmations: {} }));
  await tab.click("#prefs-apply");
  await tab.waitForFunction(() => window.__doc.settings["vefeyat.prefix"] === "v");
  check("Tercihleri al applies settings and lists", (await tab.evaluate(() => window.__doc.settings["vefeyat.exclusions"])) === '["deneme"]');
  await tab.close();
  tab = await open("?view=saglik");
  await tab.waitForFunction(() => !document.getElementById("health-summary").textContent.startsWith("Denetleniyor"));
  check("Belge sağlığı reports the document's dates", (await tab.isVisible("#view-saglik")) && (await tab.textContent("#health-summary")).length > 0, await tab.textContent("#health-summary"));
  await tab.close();
  // A document from the VBA add-in: its ADDIN date field becomes a Vefeyât date.
  {
    const text = "Bu konuda Gazzâlî (ö. 505/1111) der ki.";
    const at = text.indexOf(" (ö.");
    tab = await open("?view=saglik", text, "", { fields: [{ code: " ADDIN Vefeyat.Date.D0123.P1 ", s: at, e: at + " (ö. 505/1111)".length }], api: ["WordApi 1.4"] });
    await tab.waitForSelector("#legacy:not([hidden])");
    await tab.click("#legacy-convert");
    await tab.waitForFunction(() => window.__doc.fields.every((f) => f.gone));
    const result = await tab.evaluate(() => ({ text: window.__doc.text, tags: window.__doc.controls.map((c) => c.tag + "=" + window.__doc.text.slice(c.range.s, c.range.e)) }));
    check("VBA dates are converted to Vefeyât dates", result.text === text && result.tags.length === 1 && result.tags[0].endsWith("= (ö. 505/1111)") && result.tags[0].startsWith("vefeyat:date:"), JSON.stringify(result));
  }
  await tab.close();
  // Toplu düzeltme inserts the first-mention date; history and health follow it.
  tab = await open("?view=toplu", "Gazzâlî der ki.\rSonra Gazzâlî yine der.");
  await tab.waitForFunction(() => !document.getElementById("bulk-summary").textContent.includes("hazırlanıyor"));
  const bulkItems = await tab.locator("#bulk-list li").count();
  await tab.click("#bulk-apply");
  await tab.waitForFunction(() => window.__doc.text.includes("(ö. 505/1111)"));
  check("Toplu düzeltme adds only the first-mention date", bulkItems === 1 && (await docText(tab)) === "Gazzâlî (ö. 505/1111) der ki.\rSonra Gazzâlî yine der.", JSON.stringify(await docText(tab)));
  await tab.evaluate(() => { history.replaceState(null, "", "?view=gecmis"); });
  await tab.reload();
  await tab.waitForSelector("#view-gecmis:not([hidden])");
  check("İşlem geçmişi lists the change after reopening", ((await tab.textContent("#history-list")) ?? "").includes("Gazzâlî (ö. 505/1111)"));
  await tab.close();
  // Toplu düzeltme offers a wrong written date unchecked and corrects it once checked.
  tab = await open("?view=toplu", "Gazzâlî (ö. 506/1112) der ki.");
  await tab.waitForFunction(() => !document.getElementById("bulk-summary").textContent.includes("hazırlanıyor"));
  const fixUnchecked = !(await tab.isChecked("#bulk-list li input"));
  await tab.check("#bulk-list li input");
  await tab.click("#bulk-apply");
  await tab.waitForFunction(() => window.__doc.text.includes("505/1111"));
  check("a wrong written date is offered unchecked and corrected", fixUnchecked && (await docText(tab)) === "Gazzâlî (ö. 505/1111) der ki.", await docText(tab));
  await tab.close();
  // İlk geçiş bakımı: a new earlier mention receives the date, the later one loses it (with approval).
  tab = await open("?view=toplu", "Gazzâlî burada anılır.\rSonra Gazzâlî (ö. 505/1111) der ki.");
  await tab.waitForFunction(() => !document.getElementById("bulk-summary").textContent.includes("hazırlanıyor"));
  const moveOffered = (await tab.textContent("#bulk-summary")).includes("İlk geçiş değişikliği: 1") && (await tab.isChecked("#bulk-list li input"));
  await tab.click("#bulk-apply");
  await tab.waitForFunction(() => window.__doc.text.startsWith("Gazzâlî (ö."));
  check("a date moves to the new first mention once approved", moveOffered && (await docText(tab)) === "Gazzâlî (ö. 505/1111) burada anılır.\rSonra Gazzâlî der ki.", await docText(tab));
  await tab.close();
  // After an edit the window tells that a first mention moved (paragraph events, WordApi 1.6).
  tab = await open("", "Sonra Gazzâlî (ö. 505/1111) der ki.", "Gazzâlî", { api: ["WordApi 1.6"] });
  await tab.evaluate(() => window.__fireSelection?.());
  await tab.waitForTimeout(2600);
  await tab.evaluate(() => {
    window.__doc.text = "Gazzâlî burada anılır.\rSonra Gazzâlî (ö. 505/1111) der ki.";
    for (const h of window.__doc.handlers.added) h({ uniqueLocalIds: ["p0"] });
    window.__fireSelection?.();
  });
  await tab.waitForSelector("#firstuse-notice:not([hidden])", { timeout: 6000 }).catch(() => {});
  check("an edit that moves a first mention is reported", ((await tab.textContent("#firstuse-text")) ?? "").includes("Gazzâlî") && (await docText(tab)).includes("Sonra Gazzâlî (ö. 505/1111)"), await tab.textContent("#firstuse-text"));
  await tab.close();
  // Footnotes and endnotes (WordApi 1.5): read, selected, kept apart, and dated in their own right.
  {
    const notes = [
      { kind: "foot", para: 0, text: "Fahreddin er-Râzî de böyle der." },
      { kind: "end", para: 1, text: "Gazzâlî'nin görüşü için bk. Gazzâlî, el-Mustasfâ, I, 45." },
    ];
    tab = await open("?view=dokum", "Gazzâlî der.\rSonra başka bir konu.", "", { api: ["WordApi 1.5"], notes });
    await tab.waitForFunction(() => document.querySelectorAll("#inventory-list li").length === 3);
    const summary = await tab.textContent("#inventory-summary");
    const badges = await tab.$$eval("#inventory-list .badge", (b) => b.map((x) => x.textContent));
    check("footnotes and endnotes are read and labelled", summary.includes("okunan: 1 dipnot, 1 sonnot") && badges.join(",") === "Metin,Dipnot,Sonnot" && (await tab.textContent("#inventory-list")).includes("Gazzâlî'nin"), `${summary} | ${badges}`);
    await tab.click("#inventory-list li:nth-child(2)");
    await tab.waitForTimeout(300);
    const picked = await tab.evaluate(() => (window.__doc.selStory === window.__doc ? "main" : window.__doc.selStory.kind) + ":" + window.__doc.selStory.text.slice(...window.__doc.sel));
    check("a note row selects the name inside the note", picked === "foot:Fahreddin er-Râzî", picked);
    await tab.close();
  }
  {
    const notes = [
      { kind: "foot", para: 0, text: "Gazzâlî der." },
      { kind: "foot", para: 0, text: "Gazzâlî (ö. 505/1111) de der." },
    ];
    tab = await open("?view=toplu", "Gazzâlî (ö. 505/1111) metinde.", "", { api: ["WordApi 1.5"], notes });
    await tab.waitForFunction(() => !document.getElementById("bulk-summary").textContent.includes("hazırlanıyor"));
    const offered = (await tab.textContent("#bulk-summary")).includes("İlk geçiş değişikliği: 1");
    await tab.click("#bulk-apply");
    await tab.waitForFunction(() => window.__doc.notes[0].text.includes("505/1111"), null, { timeout: 5000 }).catch(() => {});
    const texts = await tab.evaluate(() => [window.__doc.text, ...window.__doc.notes.map((n) => n.text)]);
    check("notes keep their own first mention: the date moves inside the notes", offered && texts.join(" | ") === "Gazzâlî (ö. 505/1111) metinde. | Gazzâlî (ö. 505/1111) der. | Gazzâlî de der.", texts.join(" | "));
    await tab.close();
  }
  {
    const notes = [
      { kind: "foot", para: 0, text: "Gazzâlî der." },
      { kind: "foot", para: 0, text: "Sonra Gazzâlî de der." },
    ];
    tab = await open("", "Giriş paragrafı.", "", { api: ["WordApi 1.5"], notes });
    await tab.evaluate(() => {
      const note = window.__doc.notes[1];
      window.__doc.selStory = note;
      window.__doc.sel = [note.text.indexOf("Gazzâlî"), note.text.indexOf("Gazzâlî") + "Gazzâlî".length];
      window.__fireSelection();
    });
    await tab.waitForSelector("#details:not([hidden])");
    await tab.click("#insert");
    await tab.waitForTimeout(600);
    check("in a note, an earlier note mention is an earlier mention", (await tab.isVisible("#earlier")) && !(await tab.evaluate(() => window.__doc.notes[1].text.includes("("))));
    await tab.close();
  }
  {
    // A main-text mention before does not count for the first note mention: the date goes into the note.
    tab = await open("", "Gazzâlî metinde.", "", { api: ["WordApi 1.5"], notes: [{ kind: "foot", para: 0, text: "Gazzâlî der." }] });
    await tab.evaluate(() => {
      const note = window.__doc.notes[0];
      window.__doc.selStory = note;
      window.__doc.sel = [0, "Gazzâlî".length];
      window.__fireSelection();
    });
    await tab.waitForSelector("#details:not([hidden])");
    await tab.click("#insert");
    await tab.waitForFunction(() => window.__doc.notes[0].text.includes("505/1111"), null, { timeout: 5000 }).catch(() => {});
    check("the first note mention is dated although the main text named him before", !(await tab.isVisible("#earlier")) && (await tab.evaluate(() => window.__doc.notes[0].text)) === "Gazzâlî (ö. 505/1111) der.", await tab.evaluate(() => window.__doc.notes[0].text));
    await tab.close();
  }
  // Şahıs değil: the phrase leaves the inventory and appears in Belge istisnaları.
  tab = await open("", "Gazzâlî der ki.", "Gazzâlî");
  await tab.waitForSelector("#not-person:not([hidden])");
  await tab.click("#not-person");
  await tab.waitForSelector("#excluded-note:not([hidden])");
  await tab.evaluate(() => history.replaceState(null, "", "?view=istisna"));
  await tab.reload();
  await tab.waitForSelector("#view-istisna:not([hidden])");
  // The list is filled after the document settings are read: wait for it, not only for the view.
  await tab.waitForFunction(() => (document.querySelector("#exclusion-list")?.textContent ?? "").includes("gazzali"), null, { timeout: 5000 }).catch(() => {});
  check("Şahıs değil is stored as a document exclusion", ((await tab.textContent("#exclusion-list")) ?? "").includes("gazzali"), await tab.textContent("#exclusion-list"));
  await tab.evaluate(() => history.replaceState(null, "", "?view=kapsam"));
  await tab.reload();
  await tab.waitForFunction(() => !document.getElementById("coverage-list").textContent.includes("Okunuyor"));
  check("Denetim kapsamı reports scanned parts", ((await tab.textContent("#coverage-list")) ?? "").includes("Ana metin: 1 paragraf"), await tab.textContent("#coverage-list"));
  await tab.close();
  tab = await open("?action=dia", "Gazzâlî der ki.", "Gazzâlî");
  check("DİA maddesini aç offers the article link", (await tab.getAttribute("#detail-link", "href")) === "https://islamansiklopedisi.org.tr/gazzali");
  await tab.close();
  // Identity confirmations are remembered in the document.
  tab = await open("", "Giriş paragrafı.\rBu konuda Râzî de katılır.", "Râzî");
  await tab.check("#candidate-1");
  await tab.click("#insert");
  await tab.waitForTimeout(300);
  check("chosen homonym gets its date", (await docText(tab)).includes("Râzî (ö. 313/925) de"), await docText(tab));
  check("choice saved in the document", (await tab.evaluate(() => window.__doc.settings["vefeyat.confirmations"] || "")).includes("Ebû Bekir er-Râzî"));
  await select(tab, "Giriş");
  await select(tab, "Râzî");
  check("confirmed choice is preselected next time", (await tab.isChecked("#candidate-1")) && (await tab.textContent("#candidates")).includes("bu geçiş için onayladınız"));
  await tab.close();

  // Şahıs dökümü and Tüm belgeyi denetle on a whole document.
  tab = await open("?view=dokum", "Gazzâlî (ö. 506/1112) der ki.");
  await tab.waitForFunction(() => document.querySelectorAll("#inventory-list li").length > 0);
  check("a wrong written date is reported", (await tab.textContent("#inventory-summary")).includes("1 tarih uyarısı") && (await tab.textContent("#inventory-list")).includes("katalogdan farklı"), await tab.textContent("#inventory-summary"));
  await tab.close();
  // With paragraph events, a changed paragraph is read again by itself.
  tab = await open("?view=dokum", "Gazzâlî der.\rFahreddin er-Râzî der.", "", { api: ["WordApi 1.6"] });
  await tab.waitForFunction(() => document.querySelectorAll("#inventory-list li").length === 2);
  const runsBefore = await tab.evaluate(() => window.__doc.runs);
  await tab.click("#inventory-refresh");
  await tab.waitForTimeout(300);
  check("an unchanged document is not read again", (await tab.evaluate(() => window.__doc.runs)) === runsBefore, `${runsBefore} → ${await tab.evaluate(() => window.__doc.runs)}`);
  await tab.evaluate(() => {
    window.__doc.text = "Gazzâlî der.\rFahreddin er-Râzî ve Ebû Bekir er-Râzî der.";
    for (const h of window.__doc.handlers.changed) h({ uniqueLocalIds: ["p1"] });
  });
  await tab.click("#inventory-refresh");
  await tab.waitForFunction(() => document.querySelectorAll("#inventory-list li").length === 3);
  check("a changed paragraph is read again", ((await tab.textContent("#inventory-list")) ?? "").includes("Ebû Bekir er-Râzî"));
  await tab.evaluate(() => {
    window.__doc.text = "Yeni giriş.\rGazzâlî der.\rFahreddin er-Râzî ve Ebû Bekir er-Râzî der.\rSonra Gazzâlî de.";
    for (const h of window.__doc.handlers.added) h({ uniqueLocalIds: ["p3"] });
  });
  await tab.click("#inventory-refresh");
  await tab.waitForFunction(() => document.querySelectorAll("#inventory-list li").length === 4);
  check("an added paragraph means a full read", ((await tab.textContent("#inventory-list")) ?? "").includes("Sonra") || (await tab.$$eval("#inventory-list li", (l) => l.length)) === 4);
  await tab.close();
  const book = "Gazali bu görüşü savunur.\rGazzâlî (ö. 505/1111) ise farklı düşünür. Râzî de katılır.\rSonra Gazzâlî'nin talebeleri gelir.";
  tab = await open("?view=dokum", book);
  await tab.waitForFunction(() => document.querySelectorAll("#inventory-list li").length > 0);
  const listed = await tab.$$eval("#inventory-list li strong", (n) => n.map((x) => x.textContent));
  check("Şahıs dökümü lists every mention", listed.join("|") === "Gazali|Gazzâlî|Râzî|Gazzâlî'nin", listed.join("|"));
  check("inventory summary", (await tab.textContent("#inventory-summary")).startsWith("4 geçiş · 1 kesin şahıs · 1 doğrulanacak"), await tab.textContent("#inventory-summary"));
  {
    const [download] = await Promise.all([tab.waitForEvent("download", { timeout: 5000 }), tab.click("#inventory-xlsx")]);
    const saved = await download.path();
    const bytes = saved ? (await readFile(saved)).subarray(0, 2).toString() : "";
    check("Şahıs dökümü saves as an Excel file", download.suggestedFilename() === "Adsiz-belge-sahis-dokumu.xlsx" && bytes === "PK", download.suggestedFilename());
  }
  await tab.click("#inventory-list li:nth-child(3)");
  await tab.waitForTimeout(200);
  check("clicking a row selects it in the document", (await tab.evaluate(() => window.__doc.text.slice(...window.__doc.sel))) === "Râzî");
  await tab.click("#inventory-act");
  await tab.waitForTimeout(400);
  check("row action opens the selected-name window", (await tab.isVisible("#view-secili")) && (await tab.textContent("#selected-name")) === "Râzî");
  await tab.close();
  tab = await open("?view=denetim", book);
  await tab.waitForTimeout(400);
  check("audit visits mentions in order", (await tab.isVisible("#audit-bar")) && (await tab.textContent("#audit-position")).startsWith("Denetim: 1 / 2") && (await tab.textContent("#selected-name")) === "Gazali" && (await tab.textContent("#audit-position")).includes("kimlik belgedeki diğer geçişlerden"), await tab.textContent("#audit-position"));
  await tab.click("#audit-next");
  await tab.waitForTimeout(400);
  // The dated Gazzâlî is a repeat of the person already met as "Gazali": skipped under "Yalnız ilk kullanım".
  check("audit moves to the next mention", (await tab.textContent("#selected-name")) === "Râzî" && (await tab.textContent("#audit-position")).startsWith("Denetim: 2 / 2"));
  await tab.click("#audit-next");
  await tab.waitForTimeout(400);
  check("first-use audit skips the verified repeat", (await tab.textContent("#selected-name")) === "Râzî" && (await tab.textContent("#audit-next")) === "Bitir");
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
  check("reset returns to the published catalog", (await tab.textContent("#catalog-status")).includes("5 şahıs · eklentiyle gelen katalog"), await tab.textContent("#catalog-status"));
  await tab.close();

  // SMOKE_VIEWS=<folder>: a screenshot of every window at pane width, to look for cramped layouts.
  if (process.env.SMOKE_VIEWS) {
    const text = "Gazzâlî der.\rSonra Gazali'nin eseri ve Râzî.\rFahreddin er-Râzî (ö. 606/1210) de böyle der.";
    for (const query of ["?view=dokum", "?view=toplu", "?view=saglik", "?view=teslim", "?view=katalog", "?view=tercihler", "?view=yardim", "?view=ara", "?view=gecmis", "?view=kimlik", "?view=istisna", "?view=kapsam"]) {
      const shot = await open(query, text, "Râzî", { api: ["WordApi 1.6", "KeyboardShortcuts 1.1"] });
      await shot.waitForTimeout(800);
      await shot.screenshot({ path: `${process.env.SMOKE_VIEWS}/${query.slice(6)}.png`, fullPage: true });
      await shot.close();
    }
    console.log("views written");
  }
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
