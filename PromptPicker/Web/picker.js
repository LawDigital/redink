/* Offline, host-validated model. Uses textContent only for untrusted library data. */
(() => {
  "use strict";
  let prompts = [];
  let visible = [];
  let selectedId = null;
  let committing = false;
  const el = id => document.getElementById(id);
  const post = data => window.chrome?.webview?.postMessage(data);
  const search = el("search");
  const category = el("category");
  const results = el("results");
  const insert = el("insert");
  const saveSelection = el("saveSelection");
  const chosen = () => visible.find(p => p.Id === selectedId) || null;
  const workspace = el("workspace") || document.querySelector(".workspace");
  const splitter = el("splitter");
  let previewPercent = 55;
  let movingDivider = false;
  let lastPointer = null;
  const within = (number, lower, upper) => Math.max(lower, Math.min(upper, number));

  function resizePanels(value, report = false) {
    const available = Math.max(1, workspace.clientWidth - splitter.offsetWidth);
    // At narrow widths keep a minimum usable region for both columns.
    const minList = Math.min(195, available * .42);
    const minPreview = Math.min(230, available * .45);
    const minPreviewPercent = Math.max(32, Math.ceil(100 * minPreview / available));
    const maxPreviewPercent = Math.min(68, Math.floor(100 * (available - minList) / available));
    previewPercent = within(Math.round(value), minPreviewPercent, maxPreviewPercent);
    const left = available * (100 - previewPercent) / 100;
    workspace.style.gridTemplateColumns = `${left}px ${splitter.offsetWidth}px minmax(0, 1fr)`;
    splitter.setAttribute("aria-valuenow", String(previewPercent));
    splitter.setAttribute("aria-valuetext", `Vorschau ${previewPercent} Prozent`);
    if (report) post({type:"split",previewPercent});
  }
  splitter.addEventListener("pointerdown", e => {
    if (e.button !== 0) return;
    movingDivider = true;
    splitter.setPointerCapture(e.pointerId);
    e.preventDefault();
  });
  splitter.addEventListener("pointermove", e => {
    if (!movingDivider) return;
    const bounds = workspace.getBoundingClientRect();
    const leftPercent = 100 * (e.clientX - bounds.left - splitter.offsetWidth / 2) / (workspace.clientWidth - splitter.offsetWidth);
    resizePanels(100 - leftPercent);
  });
  function endResize() {
    if (!movingDivider) return;
    movingDivider = false;
    resizePanels(previewPercent, true);
  }
  splitter.addEventListener("pointerup", endResize);
  splitter.addEventListener("pointercancel", endResize);
  splitter.addEventListener("lostpointercapture", endResize);
  splitter.addEventListener("keydown", e => {
    if (!["ArrowLeft","ArrowRight","Home","End"].includes(e.key)) return;
    e.preventDefault();
    const value = e.key === "Home" ? 68 : e.key === "End" ? 32 : previewPercent + (e.key === "ArrowLeft" ? 2 : -2);
    resizePanels(value, true);
  });
  new ResizeObserver(() => resizePanels(previewPercent)).observe(workspace);


  function categories() {
    const previous = category.value;
    category.replaceChildren(new Option("Alle Kategorien", "*"));
    const names = new Map();
    for (const p of prompts) {
      const label = String(p.Category || "");
      const key = label.toLocaleLowerCase();
      if (!names.has(key)) names.set(key, label);
    }
    for (const name of [...names.values()].sort((a,b) => a.localeCompare(b))) category.add(new Option(name, name));
    category.value = [...category.options].some(x => x.value === previous) ? previous : "*";
  }
  function paintPreview() {
    const p = chosen();
    el("previewTitle").textContent = p?.Title || "Prompt auswählen";
    el("previewCategory").textContent = p?.Category || "";
    el("previewCategory").hidden = !p;
    el("previewText").textContent = p ? p.Prompt.replace(/\\n/g, "\n") : "Links einen Prompt auswählen.";
    el("previewSource").textContent = p ? (p.Source === "local" ? "Lokal" : "Zentral") : "";
    insert.disabled = !p || committing;
  }
  function choose(id) {
    if (!visible.some(x => x.Id === id)) return;
    selectedId = id;
    paintPreview();
    for (const node of results.children) {
      const active = Number(node.dataset.id) === id;
      node.classList.toggle("active", active);
      node.setAttribute("aria-selected", active ? "true" : "false");
    }
  }
  function paint() {
    const term = search.value.trim().toLocaleLowerCase();
    const cat = category.value;
    visible = prompts.filter(p =>
      (cat === "*" || p.Category.toLocaleLowerCase() === cat.toLocaleLowerCase()) &&
      (!term || [p.Title, p.Prompt, p.Category].some(t => t.toLocaleLowerCase().includes(term))));
    if (!visible.some(p => p.Id === selectedId)) selectedId = visible.length ? visible[0].Id : null;
    const fragment = document.createDocumentFragment();
    for (const p of visible) {
      const item = document.createElement("button");
      item.type = "button";
      item.className = "item";
      item.dataset.id = String(p.Id);
      item.setAttribute("role", "option");
      item.setAttribute("aria-selected", "false");
      item.title = p.Title;
      const heading = document.createElement("span");
      heading.className = "itemTitle";
      heading.textContent = p.Title;
      const meta = document.createElement("span");
      meta.className = "itemMeta";
      const tag = document.createElement("span");
      tag.className = "tag";
      tag.textContent = p.Category === "(Uncategorized)" ? "" : p.Category;
      const origin = document.createElement("span");
      origin.className = "origin";
      origin.textContent = p.Source === "local" ? "Lokal" : "";
      meta.append(tag, origin);
      item.append(heading, meta);
      fragment.appendChild(item);
    }
    results.replaceChildren(fragment);
    el("count").textContent = String(visible.length);
    el("empty").hidden = visible.length !== 0;
    paintPreview();
    if (selectedId !== null) choose(selectedId);
  }
  function commit() {
    const p = chosen();
    if (!p || committing) return;
    committing = true;
    insert.disabled = true;
    post({type: "select", id: p.Id});
  }
  // WebView2 hosts may produce click/dblclick with different detail semantics.
  // Pointer-up independently recognizes two presses on the same row, without
  // assuming a native dblclick event is delivered to the page. All routes
  // converge on a single guarded commit message.
  results.addEventListener("pointerup", e => {
    const row = e.target.closest(".item");
    if (!row || !results.contains(row) || committing || e.button !== 0) return;
    const id = Number(row.dataset.id);
    const t = performance.now();
    choose(id);
    if (lastPointer && lastPointer.id === id && (t - lastPointer.time) < 550 &&
        Math.abs(e.clientX - lastPointer.x) < 9 && Math.abs(e.clientY - lastPointer.y) < 9) {
      lastPointer = null;
      commit();
    } else {
      lastPointer = {id, time:t, x:e.clientX, y:e.clientY};
    }
  });
  results.addEventListener("click", e => {
    const row = e.target.closest(".item");
    if (!row || !results.contains(row) || committing) return;
    choose(Number(row.dataset.id));
    if (e.detail >= 2) commit();
  });
  results.addEventListener("dblclick", e => {
    const row = e.target.closest(".item");
    if (!row || !results.contains(row) || committing) return;
    choose(Number(row.dataset.id));
    commit();
  });
  window.chrome?.webview?.addEventListener("message", event => {
    const model = event.data;
    if (!model || model.type !== "model" || !Array.isArray(model.items)) return;
    prompts = model.items;
    committing = false;
    if (Number.isInteger(model.previewPercent)) resizePanels(model.previewPercent);
    if (selectedId === null || !prompts.some(p => p.Id === selectedId)) selectedId = prompts.length ? prompts[0].Id : null;
    saveSelection.disabled = !model.canSave;
    saveSelection.title = !model.hasLocal ? "Lokales Speichern ist nicht eingerichtet" : !model.hasSelection ? "Zuerst Text in der Zielanwendung markieren" : model.canSave ? "Markierung als Prompt speichern" : "Lokale Bibliothek ist nicht beschreibbar";
    el("saveHint").textContent = !model.hasLocal ? "Lokales Speichern nicht verfügbar" : !model.hasSelection ? "Text markieren, um einen Prompt zu speichern" : model.canSave ? "Markierung als Prompt speichern" : "Lokales Speichern nicht verfügbar";
    categories();
    paint();
    focusSearch();
  });
  function focusSearch() {
    if (!movingDivider) search.focus({preventScroll:true});
  }
  // WebView2's top-level HWND can be active while no DOM input owns the focus.
  // Reapply input focus whenever the picker itself regains keyboard activation.
  window.addEventListener("focus", () => requestAnimationFrame(focusSearch));
  search.addEventListener("input", paint);
  category.addEventListener("change", paint);
  insert.addEventListener("click", commit);
  el("cancel").addEventListener("click", () => post({type:"close"}));
  el("close").addEventListener("click", () => post({type:"close"}));
  saveSelection.addEventListener("click", () => { if (!saveSelection.disabled) post({type:"save", category:category.value}); });
  document.addEventListener("keydown", e => {
    if (e.key === "Escape") {e.preventDefault(); post({type:"close"}); return;}
    if (e.ctrlKey && e.key.toLowerCase() === "k") {e.preventDefault(); search.focus(); search.select(); return;}
    if ((e.key === "ArrowDown" || e.key === "ArrowUp") && !e.altKey && !e.ctrlKey && !e.metaKey &&
        document.activeElement !== category && document.activeElement !== splitter) {
      e.preventDefault();
      if (!visible.length) return;
      const position = visible.findIndex(p => p.Id === selectedId);
      const delta = e.key === "ArrowDown" ? 1 : -1;
      const next = visible[within(position + delta, 0, visible.length - 1)];
      if (next) {
        choose(next.Id);
        results.querySelector(`[data-id="${next.Id}"]`)?.scrollIntoView({block:"nearest"});
      }
      return;
    }
    if (e.key === "Enter" && !e.shiftKey && !e.ctrlKey && !e.altKey && !e.metaKey &&
        document.activeElement !== category && document.activeElement !== splitter &&
        document.activeElement !== saveSelection && document.activeElement !== el("cancel") && document.activeElement !== el("close")) {
      e.preventDefault(); commit();
    }
  });
})();
