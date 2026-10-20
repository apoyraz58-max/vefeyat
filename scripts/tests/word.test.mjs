import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";
import vm from "node:vm";

const source = readFileSync(new URL("../../static/vefeyat/word.js", import.meta.url), "utf8");

// These mocks exercise the taskpane's decisions and plain-text rendering.
// Windows Word is still required to verify the real Office bridge and Undo.
class Element {
  constructor(tag = "div") {
    this.tagName = tag;
    this.children = [];
    this.dataset = {};
    this.attributes = {};
    this.listeners = new Map();
    this.textContent = "";
    this.value = "";
    this.hidden = false;
    this.disabled = false;
  }
  addEventListener(event, callback) {
    if (!this.listeners.has(event)) this.listeners.set(event, []);
    this.listeners.get(event).push(callback);
  }
  async dispatch(event) {
    if (this.disabled && event === "click") return;
    for (const callback of this.listeners.get(event) || []) {
      await callback({ preventDefault() {} });
    }
    await settle();
  }
  appendChild(child) {
    if (child.tagName === "fragment") this.children.push(...child.children);
    else this.children.push(child);
    return child;
  }
  replaceChildren(...children) { this.children = children; }
  setAttribute(key, value) { this.attributes[key] = value; }
}

async function settle() {
  for (let i = 0; i < 4; i++) await new Promise((resolve) => setImmediate(resolve));
}

const example = {
  id: 42,
  short_name: "İbn Sînâ",
  full_name: "Ebû Alî el-Hüseyn b. Abdillâh b. Alî b. Sînâ",
  original_name: "ابن سينا",
  date_text: "(ö. 428/1037)",
  description: "İslâm filozofu, hekim.",
  source_url: "https://example.org/ibn-sina",
};

function createApp({ office = true, selectedText = "İbn Sînâ", paragraphText, people = [example], apiError = false,
  supported = true, compareError = false, trackingError = false } = {}) {
  const ids = ["status", "get-selection", "insert-date", "search-button", "more", "insert-help", "selection-hint",
    "chosen-name", "chosen-date", "insert-panel", "results", "query", "search-form", "host-notice"];
  const nodes = Object.fromEntries(ids.map((id) => [id, new Element()]));
  let currentText = selectedText;
  let currentLocation = 1;
  let paragraph = paragraphText ?? `${selectedText} hakkında bir cümle.`;
  const inserts = [];
  const requests = [];
  const errors = [];
  const tracked = new Set();
  let runCount = 0;
  const sandbox = {
    document: {
      getElementById: (id) => nodes[id],
      createElement: (tag) => new Element(tag),
      createDocumentFragment: () => new Element("fragment"),
    },
    window: { setTimeout, clearTimeout },
    AbortController,
    URL,
    URLSearchParams,
    fetch: async (url, options) => {
      requests.push({ url, options });
      if (apiError) throw new Error("Network unavailable");
      return {
        ok: true,
        json: async () => ({ query: "ibn", count: people.length, page: 1, page_size: 20, results: people, has_next: false }),
      };
    },
    console: { error: (...args) => errors.push(args) },
  };
  if (office) {
    sandbox.Office = {
      HostType: { Word: "Word" },
      onReady: (callback) => callback({ host: "Word" }),
      context: { requirements: { isSetSupported: (name, version) => supported && name === "WordApi" && version === "1.3" } },
    };
    sandbox.Word = {
      InsertLocation: { end: "End" },
      LocationRelation: { equal: "Equal" },
      run: async (objectOrCallback, maybeCallback) => {
        runCount++;
        const callback = maybeCallback || objectOrCallback;
        if (maybeCallback && !tracked.has(objectOrCallback)) throw new Error("InvalidObjectPath");
        const selection = {
          location: currentLocation,
          get text() { return currentText; },
          load() {},
          paragraphs: { items: [{ text: paragraph }], load() {} },
          compareLocationWith(other) {
            if (compareError) throw new Error("Cannot compare selection locations");
            return { value: this.location === other.location ? "Equal" : "Before" };
          },
          insertText(text, location) {
            inserts.push({ text, location });
            paragraph = paragraph.replace(currentText, currentText + text);
            return { select() { currentText = text; } };
          },
        };
        return callback({ document: { getSelection: () => selection }, sync: async () => {}, trackedObjects: {
          add(range) {
            if (trackingError) throw new Error("Cannot track selection");
            tracked.add(range);
          },
          remove(range) { tracked.delete(range); },
        } });
      },
    };
  }
  vm.runInNewContext(source, sandbox, { filename: "word.js" });
  return {
    nodes,
    requests,
    inserts,
    errors,
    get runCount() { return runCount; },
    get trackedCount() { return tracked.size; },
    setSelection(text, location = currentLocation) { currentText = text; currentLocation = location; },
    async capture() { await nodes["get-selection"].dispatch("click"); },
    async search(term) {
      nodes.query.value = term;
      await nodes.query.dispatch("input");
      await nodes["search-form"].dispatch("submit");
    },
    async choose(index = 0) {
      const radio = nodes.results.children[index].children[0].children[0];
      await radio.dispatch("change");
    },
    async insert() { await nodes["insert-date"].dispatch("click"); },
  };
}

test("valid Word selection appends the original date and disables repeated insertion", async () => {
  const app = createApp();
  await app.capture();
  assert.equal(app.trackedCount, 1);
  assert.equal(app.requests.length, 1);
  assert.equal(new URL(app.requests[0].url, "https://example.org").searchParams.get("q"), "İbn Sînâ");
  await app.choose();
  assert.equal(app.nodes["insert-date"].disabled, false);
  await app.insert();
  assert.deepEqual(app.inserts, [{ text: ` ${example.date_text}`, location: "End" }]);
  assert.equal(app.nodes.status.dataset.level, "success");
  assert.equal(app.nodes["insert-date"].disabled, true);
  assert.equal(app.trackedCount, 0);
  await app.insert();
  assert.equal(app.inserts.length, 1);
});

test("a changed Word selection prevents writing and requires a new capture", async () => {
  const app = createApp();
  await app.capture();
  await app.choose();
  app.setSelection("İbn Rüşd");
  await app.insert();
  assert.equal(app.inserts.length, 0);
  assert.match(app.nodes.status.textContent, /seçim değişmiş/u);
  assert.equal(app.nodes["insert-date"].disabled, true);
});

test("the same name selected in another document location cannot receive the date", async () => {
  const app = createApp();
  await app.capture();
  await app.choose();
  app.setSelection("İbn Sînâ", 2);
  await app.insert();
  assert.equal(app.inserts.length, 0);
  assert.match(app.nodes.status.textContent, /seçim değişmiş/u);
  assert.equal(app.nodes["insert-date"].disabled, true);
  assert.equal(app.trackedCount, 0);
});

test("a location comparison failure never falls back to writing based on text alone", async () => {
  const app = createApp({ compareError: true });
  await app.capture();
  await app.choose();
  await app.insert();
  assert.equal(app.inserts.length, 0);
  assert.equal(app.nodes.status.dataset.level, "error");
  assert.equal(app.nodes["insert-date"].disabled, true);
  assert.equal(app.trackedCount, 0);
});

test("failed range tracking never enables document insertion", async () => {
  const app = createApp({ trackingError: true });
  await app.capture();
  assert.equal(app.inserts.length, 0);
  assert.equal(app.requests.length, 0);
  assert.equal(app.nodes["insert-date"].disabled, true);
  assert.equal(app.nodes.status.dataset.level, "error");
});

test("Word without WordApi 1.3 keeps search usable and document writing disabled", async () => {
  const app = createApp({ supported: false });
  assert.match(app.nodes["host-notice"].textContent, /WordApi 1\.3/u);
  assert.equal(app.nodes["get-selection"].disabled, true);
  await app.search("İbn Sînâ");
  await app.choose();
  assert.equal(app.nodes.results.children.length, 1);
  assert.equal(app.nodes["insert-date"].disabled, true);
  assert.equal(app.runCount, 0);
});

test("empty Word selection does not search or write", async () => {
  const app = createApp({ selectedText: "  " });
  await app.capture();
  assert.equal(app.requests.length, 0);
  assert.equal(app.inserts.length, 0);
  assert.match(app.nodes.status.textContent, /Önce Word/u);
});

test("an existing date immediately after a unique name is not inserted twice", async () => {
  const app = createApp({ paragraphText: `İbn Sînâ ${example.date_text} hakkında bir cümle.` });
  await app.capture();
  await app.choose();
  await app.insert();
  assert.equal(app.inserts.length, 0);
  assert.match(app.nodes.status.textContent, /zaten var/u);
});

test("a date already in the selected text is not inserted again", async () => {
  const app = createApp({ selectedText: `İbn Sînâ ${example.date_text}` });
  await app.capture();
  await app.choose();
  await app.insert();
  assert.equal(app.inserts.length, 0);
  assert.match(app.nodes.status.textContent, /zaten var/u);
});

test("manual search works without Office and document insertion stays disabled", async () => {
  const app = createApp({ office: false });
  assert.equal(app.nodes["get-selection"].disabled, true);
  await app.search("İbni Sina");
  await app.choose();
  assert.equal(app.nodes.results.children.length, 1);
  assert.equal(app.nodes["chosen-date"].textContent, example.date_text);
  assert.equal(app.nodes["insert-date"].disabled, true);
  assert.equal(app.runCount, 0);
});

test("editing the search box invalidates an earlier Word selection", async () => {
  const app = createApp();
  await app.capture();
  await app.choose();
  app.nodes.query.value = "İbn Rüşd";
  await app.nodes.query.dispatch("input");
  await app.insert();
  assert.equal(app.inserts.length, 0);
  assert.equal(app.nodes["insert-date"].disabled, true);
  assert.equal(app.trackedCount, 0);
});

test("API failure is announced without leaving the taskpane busy", async () => {
  const app = createApp({ apiError: true });
  await app.search("İbn Sînâ");
  assert.match(app.nodes.status.textContent, /ulaşılamadı/u);
  assert.equal(app.nodes.status.dataset.level, "error");
  assert.equal(app.nodes.results.attributes["aria-busy"], "false");
  assert.equal(app.nodes["insert-date"].disabled, true);
});

test("short queries clear previously chosen people, including a new one-letter Word selection", async () => {
  const app = createApp();
  await app.capture();
  await app.choose();
  app.setSelection("A");
  await app.capture();
  assert.equal(app.nodes.results.children.length, 0);
  assert.equal(app.nodes["insert-panel"].hidden, true);
  assert.equal(app.nodes["insert-date"].disabled, true);
  assert.equal(app.requests.length, 1);
});

test("result data is plain text and unsafe source links are omitted", async () => {
  const malicious = { ...example, short_name: '<img src=x onerror="alert(1)">', source_url: "javascript:alert(1)" };
  const app = createApp({ office: false, people: [malicious] });
  await app.search("ibn");
  const result = app.nodes.results.children[0];
  assert.equal(result.children[0].children[1].textContent, malicious.short_name);
  assert.equal(result.children.some((node) => node.tagName === "a"), false);
});

test("approximate and non-death date expressions are appended without reinterpretation", async () => {
  const person = { ...example, date_text: "(gaybeti 260/874)" };
  const app = createApp({ people: [person] });
  await app.capture();
  await app.choose();
  await app.insert();
  assert.equal(app.inserts[0].text, " (gaybeti 260/874)");
});
