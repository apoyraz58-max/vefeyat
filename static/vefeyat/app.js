(() => {
  "use strict";

  const byId = (id) => document.getElementById(id);
  const elements = {
    searchForm: byId("search-form"), searchInput: byId("search-input"),
    results: byId("results-list"), resultCount: byId("result-count"),
    resultError: byId("result-error"), emptyState: byId("empty-state"),
    loadMore: byId("load-more"), detailPanel: byId("detail-panel"),
    detailPlaceholder: byId("detail-placeholder"), detailStatus: byId("detail-status"),
    detailContent: byId("detail-content"), dialog: byId("contribution-dialog"),
    contributionForm: byId("contribution-form"), contributionError: byId("contribution-error"),
    contributionSubmit: byId("contribution-submit"), contributionSuccess: byId("contribution-success"),
  };
  const state = {
    query: "", page: 0, hasNext: false, loading: false,
    selected: null, searchController: null, detailController: null,
    debounceTimer: null, contributionKind: "new_person", contributionPerson: null,
    submitting: false,
  };
  const formatNumber = (value) => new Intl.NumberFormat("tr-TR").format(value);
  const fieldLabels = {
    short_name: "Şahıs adı", date_text: "Tarih ifadesi", description: "Kısa tanıtım",
    source_url: "Kaynak bağlantısı", note: "Açıklamanız", kind: "Katkı türü",
    person: "İlgili şahıs", person_id: "İlgili şahıs", __all__: "Öneri",
  };

  function node(tag, className, text) {
    const item = document.createElement(tag);
    if (className) item.className = className;
    if (text !== undefined && text !== null) item.textContent = String(text);
    return item;
  }

  function show(element, visible) { element.hidden = !visible; }

  function safeSourceUrl(value) {
    try {
      const url = new URL(value);
      return ["https:", "http:"].includes(url.protocol) ? url.href : null;
    } catch { return null; }
  }

  async function request(url, options = {}) {
    const response = await fetch(url, {
      credentials: "same-origin",
      ...options,
      headers: { Accept: "application/json", ...options.headers },
    });
    let data;
    try { data = await response.json(); }
    catch (error) {
      if (error.name === "AbortError") throw error;
      throw new Error("Sunucudan beklenen yanıt alınamadı. Lütfen yeniden deneyin.");
    }
    if (!response.ok) {
      let message = typeof data.message === "string" ? data.message : "";
      if (!message && typeof data.detail === "string") message = data.detail;
      if (!message && typeof data.error === "string") message = data.error;
      if (!message && data.errors && typeof data.errors === "object") {
        const messages = Object.values(data.errors).flat().filter((value) => typeof value === "string");
        message = messages.join(" ");
      }
      if (data.fields && typeof data.fields === "object") {
        const fieldMessages = Object.entries(data.fields).map(([field, errors]) => {
          const messages = (Array.isArray(errors) ? errors : [errors])
            .filter((value) => typeof value === "string");
          return messages.length ? `${fieldLabels[field] || "Alan"}: ${messages.join(" ")}` : "";
        }).filter(Boolean);
        message = [message, ...fieldMessages].filter(Boolean).join("\n");
      }
      const error = new Error(message || "İşlem tamamlanamadı. Lütfen bilgileri kontrol edip yeniden deneyin.");
      error.name = "ApiError";
      throw error;
    }
    return data;
  }

  function renderSkeletons() {
    const fragment = document.createDocumentFragment();
    for (let index = 0; index < 5; index += 1) {
      const card = node("div", "skeleton-card");
      card.setAttribute("aria-hidden", "true");
      for (let line = 0; line < 3; line += 1) card.append(node("div", "skeleton-line"));
      fragment.append(card);
    }
    elements.results.replaceChildren(fragment);
  }

  function renderPerson(person) {
    const card = node("button", "person-card");
    card.type = "button";
    card.dataset.personId = person.id;
    card.setAttribute("aria-pressed", String(state.selected?.id === person.id));
    card.setAttribute("aria-controls", "detail-panel");
    const top = node("span", "person-card-top");
    top.append(node("span", "person-name", person.short_name));
    if (person.date_text) top.append(node("span", "person-date", person.date_text));
    card.append(top);
    if (person.description) card.append(node("span", "person-description", person.description));
    if (person.original_name) {
      const original = node("span", "person-original", person.original_name);
      original.dir = "auto";
      card.append(original);
    }
    const arrow = node("span", "person-arrow icon-arrow");
    arrow.setAttribute("aria-hidden", "true");
    card.append(arrow);
    card.addEventListener("click", () => selectPerson(person, true));
    return card;
  }

  function resetDetail() {
    state.detailController?.abort();
    state.selected = null;
    elements.detailPanel.setAttribute("aria-busy", "false");
    show(elements.detailPlaceholder, true);
    show(elements.detailContent, false);
    show(elements.detailStatus, false);
  }

  async function loadPeople(query, append = false) {
    state.searchController?.abort();
    const controller = new AbortController();
    state.searchController = controller;
    const nextPage = append ? state.page + 1 : 1;
    state.loading = true;
    elements.results.setAttribute("aria-busy", "true");
    elements.loadMore.disabled = true;
    show(elements.resultError, false);
    show(elements.emptyState, false);
    if (!append) {
      state.query = query;
      state.page = 0;
      state.hasNext = false;
      resetDetail();
      show(elements.loadMore, false);
      renderSkeletons();
      elements.resultCount.textContent = "Kayıtlar yükleniyor";
    } else elements.loadMore.textContent = "Kayıtlar yükleniyor…";

    try {
      const params = new URLSearchParams({ q: query, page: String(nextPage), page_size: "20" });
      const data = await request(`/api/people/?${params}`, { signal: controller.signal });
      if (controller.signal.aborted) return;
      const fragment = document.createDocumentFragment();
      data.results.forEach((person) => fragment.append(renderPerson(person)));
      if (append) elements.results.append(fragment);
      else elements.results.replaceChildren(fragment);
      state.page = data.page;
      state.hasNext = data.has_next;
      elements.resultCount.textContent = data.approximate
        ? `Bunu mu demek istediniz? ${formatNumber(data.count)} sonuç`
        : query ? `${formatNumber(data.count)} sonuç bulundu` : `${formatNumber(data.count)} kayıt`;
      byId("results-heading").textContent = data.approximate
        ? "Yakın isim eşleşmeleri" : query ? "Arama sonuçları" : "Şahıslar";
      show(elements.emptyState, data.count === 0);
      show(elements.loadMore, data.has_next);
    } catch (error) {
      if (error.name === "AbortError") return;
      if (!append) {
        elements.results.replaceChildren();
        elements.resultCount.textContent = "Kayıtlar yüklenemedi";
      }
      elements.resultError.textContent = error.name === "TypeError"
        ? "Kayıtlar yüklenemedi. Bağlantınızı kontrol edip yeniden arayın."
        : error.message || "Kayıtlar yüklenemedi. Lütfen yeniden deneyin.";
      show(elements.resultError, true);
      if (append) show(elements.loadMore, true);
    } finally {
      if (state.searchController === controller) {
        state.loading = false;
        elements.results.setAttribute("aria-busy", "false");
        elements.loadMore.disabled = false;
        elements.loadMore.textContent = "Daha fazla kayıt göster ↓";
      }
    }
  }

  function populateDetail(person) {
    byId("detail-name").textContent = person.short_name;
    byId("detail-original").textContent = person.original_name || "";
    show(byId("detail-original"), Boolean(person.original_name));
    byId("detail-date").textContent = person.date_text || "Tarih bilgisi belirtilmemiş.";
    byId("detail-full-name").textContent = person.full_name || "";
    show(byId("full-name-section"), Boolean(person.full_name));
    byId("detail-description").textContent = person.description || "Kısa tanıtım belirtilmemiş.";
    const aliases = Array.isArray(person.aliases) ? person.aliases : [];
    byId("detail-aliases").replaceChildren(...aliases.map((alias) => node("span", "alias", alias)));
    show(byId("aliases-section"), aliases.length > 0);
    byId("citation-text").value = [person.short_name, person.date_text].filter(Boolean).join(" ");
    byId("copy-status").textContent = "";
    const sourceUrl = safeSourceUrl(person.source_url);
    const source = byId("detail-source");
    if (sourceUrl) source.href = sourceUrl;
    else source.removeAttribute("href");
    show(source, Boolean(sourceUrl));
    show(elements.detailContent, true);
  }

  async function selectPerson(person, userInitiated) {
    state.detailController?.abort();
    const controller = new AbortController();
    state.detailController = controller;
    state.selected = person;
    elements.results.querySelectorAll(".person-card").forEach((card) => {
      card.setAttribute("aria-pressed", String(card.dataset.personId === person.id));
    });
    show(elements.detailPlaceholder, false);
    show(elements.detailContent, false);
    elements.detailStatus.textContent = "Şahıs bilgileri yükleniyor…";
    show(elements.detailStatus, true);
    elements.detailPanel.setAttribute("aria-busy", "true");
    if (userInitiated && window.matchMedia("(max-width: 580px)").matches) {
      elements.detailPanel.scrollIntoView({ block: "start", behavior: "auto" });
    }
    try {
      const detail = await request(`/api/people/${encodeURIComponent(person.id)}/`, { signal: controller.signal });
      if (controller.signal.aborted) return;
      state.selected = detail;
      populateDetail(detail);
      show(elements.detailStatus, false);
      if (userInitiated && window.matchMedia("(max-width: 580px)").matches) {
        byId("detail-name").focus({ preventScroll: true });
      }
    } catch (error) {
      if (error.name === "AbortError") return;
      elements.detailStatus.textContent = "Şahıs bilgileri yüklenemedi. Yeniden denemek için kaydı tekrar seçin.";
    } finally {
      if (state.detailController === controller) elements.detailPanel.setAttribute("aria-busy", "false");
    }
  }

  function submitSearch() {
    clearTimeout(state.debounceTimer);
    loadPeople(elements.searchInput.value.trim());
  }

  elements.searchForm.addEventListener("submit", (event) => { event.preventDefault(); submitSearch(); });
  elements.searchInput.addEventListener("input", () => {
    clearTimeout(state.debounceTimer);
    state.debounceTimer = setTimeout(submitSearch, 350);
  });
  document.querySelectorAll("[data-search]").forEach((button) => {
    button.addEventListener("click", () => {
      elements.searchInput.value = button.dataset.search;
      submitSearch();
    });
  });
  elements.loadMore.addEventListener("click", () => {
    if (!state.loading) loadPeople(state.query, true);
  });

  byId("copy-citation").addEventListener("click", async () => {
    const input = byId("citation-text");
    try {
      await navigator.clipboard.writeText(input.value);
      byId("copy-status").textContent = "Ad ve tarih bilgisi kopyalandı.";
    } catch {
      input.focus();
      input.select();
      byId("copy-status").textContent = "Metin seçildi. Kopyalamak için Ctrl+C tuşlarına basın.";
    }
  });

  function openContribution(kind) {
    if (!state.submitting) {
      elements.contributionForm.reset();
      state.contributionKind = kind;
      state.contributionPerson = kind === "correction" ? state.selected : null;
      show(elements.contributionForm, true);
      show(elements.contributionSuccess, false);
      show(elements.contributionError, false);
      byId("contribution-heading").textContent = kind === "correction" ? "Düzeltme önerin" : "Yeni şahıs önerin";
      if (state.contributionPerson) {
        const person = state.contributionPerson;
        byId("contribution-name").value = person.short_name || "";
        byId("contribution-date").value = person.date_text || "";
        byId("contribution-description").value = person.description || "";
        byId("contribution-source").value = person.source_url || "";
      }
    }
    elements.dialog.showModal();
    document.body.classList.add("dialog-open");
  }

  function closeContribution() { elements.dialog.close(); }

  document.querySelectorAll("[data-contribute]").forEach((button) => {
    button.addEventListener("click", () => openContribution(button.dataset.contribute));
  });
  byId("correct-person").addEventListener("click", () => {
    if (state.selected) openContribution("correction");
  });
  byId("close-dialog").addEventListener("click", closeContribution);
  byId("finish-contribution").addEventListener("click", closeContribution);
  elements.dialog.addEventListener("close", () => { document.body.classList.remove("dialog-open"); });
  elements.dialog.addEventListener("click", (event) => {
    if (event.target !== elements.dialog) return;
    const box = elements.dialog.getBoundingClientRect();
    if (event.clientX < box.left || event.clientX > box.right || event.clientY < box.top || event.clientY > box.bottom) closeContribution();
  });

  function csrfToken() {
    const cookie = document.cookie.split(";").find((value) => value.trim().startsWith("csrftoken="));
    return cookie ? decodeURIComponent(cookie.trim().slice("csrftoken=".length)) : "";
  }

  elements.contributionForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    if (state.submitting || !elements.contributionForm.reportValidity()) return;
    state.submitting = true;
    elements.contributionSubmit.disabled = true;
    elements.contributionSubmit.textContent = "Gönderiliyor…";
    show(elements.contributionError, false);
    const data = Object.fromEntries(new FormData(elements.contributionForm));
    for (const key of Object.keys(data)) data[key] = data[key].trim();
    data.kind = state.contributionKind;
    data.person_id = state.contributionPerson?.id || null;
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 20000);
    try {
      const result = await request("/api/contributions/", {
        method: "POST", signal: controller.signal,
        headers: { "Content-Type": "application/json", "X-CSRFToken": csrfToken() },
        body: JSON.stringify(data),
      });
      show(elements.contributionForm, false);
      byId("contribution-success-message").textContent = result.message || "Öneriniz editör incelemesini bekliyor. Onaylandıktan sonra yayımlanır.";
      show(elements.contributionSuccess, true);
      if (elements.dialog.open) byId("finish-contribution").focus();
    } catch (error) {
      elements.contributionError.textContent = error.name === "AbortError"
        ? "Sunucu yanıtı zamanında alınamadı. Öneri alınmış olabilir; yeniden göndermeden önce bağlantınızı kontrol edin."
        : error.name === "TypeError" ? "Sunucuya bağlanılamadı. Bağlantınızı kontrol edip yeniden deneyin."
          : error.message || "Öneriniz gönderilemedi. Lütfen yeniden deneyin.";
      show(elements.contributionError, true);
    } finally {
      clearTimeout(timeout);
      state.submitting = false;
      elements.contributionSubmit.disabled = false;
      elements.contributionSubmit.textContent = "İncelemeye gönder →";
    }
  });

  request("/api/stats/").then((data) => {
    if (Number.isInteger(data.people_count) && data.people_count >= 0) {
      byId("people-count").textContent = formatNumber(data.people_count);
    }
  }).catch(() => { byId("people-count").textContent = "Vefeyât"; });
  loadPeople("");
})();
