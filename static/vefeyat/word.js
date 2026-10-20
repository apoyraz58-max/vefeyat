/* global Office, Word */
"use strict";

(() => {
  const el = (id) => document.getElementById(id);
  const state = {
    wordReady: false,
    selectionText: null,
    selectionRange: null,
    person: null,
    query: "",
    page: 0,
    count: 0,
    requestId: 0,
    controller: null,
    searching: false,
    writing: false,
    reading: false,
  };

  function status(message, level = "info") {
    el("status").textContent = message;
    el("status").dataset.level = level;
  }

  function normalized(text) {
    return String(text || "").replace(/\s+/gu, " ").trim();
  }

  function updateControls() {
    el("get-selection").disabled = !state.wordReady || state.reading || state.writing;
    el("insert-date").disabled = !state.wordReady || !state.selectionText || !state.selectionRange ||
      !state.person || !normalized(state.person.date_text) || state.writing || state.searching || state.reading;
    el("search-button").disabled = state.writing;
    el("query").disabled = state.writing || state.reading;
    el("more").disabled = state.searching || state.writing;
    if (!state.wordReady) {
      el("insert-help").textContent = "Belgeye eklemek için bu sayfayı Word eklentisi olarak açın.";
    } else if (!state.selectionText) {
      el("insert-help").textContent = "Belgeye eklemek için şahıs adını Word’de seçip “Seçili adı al” düğmesine basın.";
    } else if (state.person && !normalized(state.person.date_text)) {
      el("insert-help").textContent = "Bu kayıtta eklenecek bir tarih ifadesi bulunmuyor.";
    } else {
      el("insert-help").textContent = "Şahıs adı korunur; sonuna bir boşluk ve yukarıdaki tarih ifadesi eklenir.";
    }
  }

  function clearSelection() {
    const range = state.selectionRange;
    state.selectionRange = null;
    state.selectionText = null;
    el("selection-hint").textContent = "Belgeye eklemek için şahıs adını Word’den yeniden alın.";
    updateControls();
    if (!range) return Promise.resolve();
    // Release tracked host objects after their last use, including when a
    // manual query invalidates the captured selection. A failed cleanup never
    // restores insertion permission or causes a document write.
    return Word.run(range, async (context) => {
      context.trackedObjects.remove(range);
      await context.sync();
    }).catch(() => {});
  }

  function safeSourceUrl(value) {
    try {
      const url = new URL(value);
      return ["https:", "http:"].includes(url.protocol) ? url.href : null;
    } catch (_) {
      return null;
    }
  }

  function addText(parent, tag, value, className) {
    if (!value) return null;
    const node = document.createElement(tag);
    node.textContent = value;
    if (className) node.className = className;
    parent.appendChild(node);
    return node;
  }

  function showPerson(person) {
    state.person = person;
    el("chosen-name").textContent = person.short_name || person.full_name || "Adsız kayıt";
    el("chosen-date").textContent = person.date_text || "Tarih belirtilmemiş.";
    el("insert-panel").hidden = false;
    updateControls();
  }

  function renderPeople(people) {
    const fragment = document.createDocumentFragment();
    for (const person of people) {
      const item = document.createElement("li");
      item.className = "person";
      const label = document.createElement("label");
      const radio = document.createElement("input");
      radio.type = "radio";
      radio.name = "person";
      radio.value = String(person.id);
      radio.addEventListener("change", () => showPerson(person));
      label.appendChild(radio);
      addText(label, "span", person.short_name || person.full_name || "Adsız kayıt");
      item.appendChild(label);
      addText(item, "p", person.date_text || "Tarih belirtilmemiş.", "date");
      if (person.full_name && person.full_name !== person.short_name) {
        addText(item, "p", person.full_name, "muted");
      }
      const original = addText(item, "p", person.original_name, "muted");
      if (original) original.dir = "auto";
      addText(item, "p", person.description, "description");
      const source = safeSourceUrl(person.source_url);
      if (source) {
        const link = addText(item, "a", "Kaynağı aç ↗");
        link.href = source;
        link.target = "_blank";
        link.rel = "noopener noreferrer";
      }
      fragment.appendChild(item);
    }
    el("results").appendChild(fragment);
  }

  async function search(query, append = false) {
    const term = query.trim();
    if (state.controller) state.controller.abort();
    const requestId = ++state.requestId;
    if (term.length < 2) {
      state.controller = null;
      state.searching = false;
      state.person = null;
      state.query = "";
      state.page = 0;
      el("results").replaceChildren();
      el("results").setAttribute("aria-busy", "false");
      el("insert-panel").hidden = true;
      el("more").hidden = true;
      status("Aramak için en az iki harf yazın.", "error");
      updateControls();
      return;
    }
    const controller = new AbortController();
    state.controller = controller;
    state.searching = true;
    if (!append) {
      state.person = null;
      state.query = term;
      state.page = 0;
      el("results").replaceChildren();
      el("insert-panel").hidden = true;
    }
    el("more").hidden = true;
    el("results").setAttribute("aria-busy", "true");
    updateControls();
    status(append ? "Diğer sonuçlar getiriliyor…" : "Şahıslar aranıyor…");
    const timeout = window.setTimeout(() => controller.abort(), 20000);
    try {
      const params = new URLSearchParams({ q: term, page: String(state.page + 1), page_size: "20" });
      const response = await fetch(`/api/people/?${params}`, {
        headers: { Accept: "application/json" },
        credentials: "same-origin",
        signal: controller.signal,
      });
      if (!response.ok) throw new Error("api");
      const data = await response.json();
      if (!Array.isArray(data.results) || !Number.isInteger(data.page) || !Number.isFinite(data.count)) {
        throw new Error("format");
      }
      if (requestId !== state.requestId) return;
      state.page = data.page;
      state.count = data.count;
      renderPeople(data.results);
      el("more").hidden = !data.has_next;
      const shown = el("results").children.length;
      status(data.count ? `${data.count.toLocaleString("tr-TR")} kayıt bulundu. ${shown} kayıt gösteriliyor; doğru şahsı seçin.` :
        "Sonuç bulunamadı. İsmin başka bir yazılışını veya daha kısa bir bölümünü deneyin.");
    } catch (error) {
      if (requestId !== state.requestId) return;
      status(error.name === "AbortError" ? "Arama zaman aşımına uğradı. Yeniden deneyin." :
        "Arama servisine ulaşılamadı. Bağlantınızı kontrol edip yeniden deneyin.", "error");
      // Keep an existing results page available if fetching the next page failed.
      if (append) el("more").hidden = false;
    } finally {
      window.clearTimeout(timeout);
      if (requestId === state.requestId) {
        state.searching = false;
        state.controller = null;
        el("results").setAttribute("aria-busy", "false");
        updateControls();
      }
    }
  }

  async function getSelection() {
    if (!state.wordReady || state.reading || state.writing) return;
    state.reading = true;
    await clearSelection();
    status("Word’deki seçim alınıyor…");
    try {
      const captured = await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();
        const text = selection.text;
        if (!normalized(text) || text.trim().length > 200 || /[\r\n]/u.test(text)) {
          return { text, range: null };
        }
        // Persist the exact range across the asynchronous search. Word adjusts
        // tracked ranges when surrounding document content changes.
        context.trackedObjects.add(selection);
        await context.sync();
        return { text, range: selection };
      });
      const { text } = captured;
      if (!normalized(text)) {
        clearSelection();
        status("Önce Word belgesinde şahıs adını seçin. Boş seçime metin eklenmez.", "error");
        return;
      }
      if (text.trim().length > 200 || /[\r\n]/u.test(text)) {
        clearSelection();
        status("Yalnızca aramak istediğiniz şahıs adını seçin; paragraf veya uzun metin seçmeyin.", "error");
        return;
      }
      state.selectionText = text;
      state.selectionRange = captured.range;
      el("selection-hint").textContent = `Word’den alınan ad: “${text}”`;
      el("query").value = text.trim();
      await search(text);
    } catch (_) {
      clearSelection();
      status("Word’deki seçim okunamadı. Belgeye dönüp şahıs adını yeniden seçin.", "error");
    } finally {
      state.reading = false;
      updateControls();
    }
  }

  function alreadyFollowedByDate(paragraphText, selectedText, dateText) {
    // Only inspect a paragraph when the
    // selected text occurs exactly once, so another occurrence cannot block it.
    const index = paragraphText.indexOf(selectedText);
    if (index < 0 || paragraphText.indexOf(selectedText, index + selectedText.length) !== -1) return false;
    const following = normalized(paragraphText.slice(index + selectedText.length));
    const date = normalized(dateText);
    return following.startsWith(date);
  }

  async function insertDate() {
    if (!state.wordReady || !state.selectionText || !state.selectionRange || !state.person || state.writing || state.searching || state.reading) return;
    const originalSelection = state.selectionText;
    const originalRange = state.selectionRange;
    const date = String(state.person.date_text || "");
    if (!normalized(date)) return;
    state.writing = true;
    updateControls();
    status("Seçim doğrulanıyor…");
    try {
      const result = await Word.run(originalRange, async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        const location = selection.compareLocationWith(originalRange);
        const paragraphs = selection.paragraphs;
        paragraphs.load("items/text");
        await context.sync();
        // Both the exact text and its document location must still match. A
        // second occurrence of the same name cannot become the write target.
        if (!normalized(selection.text) || selection.text !== originalSelection ||
            location.value !== Word.LocationRelation.equal) return "changed";
        if (normalized(selection.text).includes(normalized(date)) ||
            (paragraphs.items.length === 1 && alreadyFollowedByDate(paragraphs.items[0].text, selection.text, date))) {
          return "duplicate";
        }
        // A single insert preserves the selected name and Word's normal Undo.
        const inserted = selection.insertText(` ${date}`, Word.InsertLocation.end);
        inserted.select();
        await context.sync();
        return "inserted";
      });
      await clearSelection();
      if (result === "changed") {
        status("Word’deki seçim değişmiş. Hiçbir şey eklenmedi. Şahıs adını yeniden seçip “Seçili adı al” düğmesine basın.", "error");
      } else if (result === "duplicate") {
        status("Bu tarih seçili metinde veya hemen sonrasında zaten var. Tekrar eklenmedi.", "error");
      } else {
        status(`“${date}” eklendi. Gerekirse Word’de Geri Al komutunu kullanın.`, "success");
      }
    } catch (_) {
      await clearSelection();
      status("Word işlemi tamamlanamadı. Belgeyi kontrol edin; ekleme gerçekleştiyse yeniden eklemeyin. Korunan belgelerde düzenleme izni gerekir.", "error");
    } finally {
      state.writing = false;
      updateControls();
    }
  }

  el("search-form").addEventListener("submit", (event) => {
    event.preventDefault();
    if (!state.writing) search(el("query").value);
  });
  el("query").addEventListener("input", clearSelection);
  el("get-selection").addEventListener("click", getSelection);
  el("insert-date").addEventListener("click", insertDate);
  el("more").addEventListener("click", () => search(state.query, true));
  updateControls();

  if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
    Office.onReady((info) => {
      if (info.host !== Office.HostType.Word) return;
      if (!Office.context.requirements.isSetSupported("WordApi", "1.3")) {
        el("host-notice").textContent = "Bu Word sürümü güvenli seçim doğrulaması için gerekli WordApi 1.3 desteğini sunmuyor. Arama yapabilirsiniz; belgeye eklemek için Word’ü güncelleyin.";
        return;
      }
      state.wordReady = true;
      el("host-notice").textContent = "Word bağlantısı hazır. Belgedeki şahıs adını seçerek başlayın.";
      updateControls();
    });
  }
})();
