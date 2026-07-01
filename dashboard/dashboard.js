const state = {
  payload: null,
  search: "",
  date: "all",
  status: "all",
  expandedRows: new Set(),
};

const statusOrder = ["unchecked", "relevant", "not_relevant", "analyzing", "submitting"];

const els = {
  content: document.getElementById("content"),
  generatedAt: document.getElementById("generatedAt"),
  statNew: document.getElementById("statNew"),
  statUnchecked: document.getElementById("statUnchecked"),
  statNotRelevant: document.getElementById("statNotRelevant"),
  statInWork: document.getElementById("statInWork"),
  searchInput: document.getElementById("searchInput"),
  dateFilter: document.getElementById("dateFilter"),
  statusFilter: document.getElementById("statusFilter"),
  refreshBtn: document.getElementById("refreshBtn"),
  cleanupBtn: document.getElementById("cleanupBtn"),
  toast: document.getElementById("toast"),
};

function make(tag, className, text) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}

function formatDateTime(value) {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleString("ru-RU", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function showToast(message) {
  els.toast.textContent = message;
  els.toast.classList.add("is-visible");
  window.clearTimeout(showToast.timer);
  showToast.timer = window.setTimeout(() => els.toast.classList.remove("is-visible"), 4200);
}

async function fetchJson(url, options = {}) {
  const response = await fetch(url, {
    headers: { "Content-Type": "application/json" },
    ...options,
  });
  const payload = await response.json();
  if (!response.ok) {
    throw new Error(payload.error || `HTTP ${response.status}`);
  }
  return payload;
}

async function loadData() {
  els.generatedAt.textContent = "Загрузка";
  const payload = await fetchJson("/api/items");
  state.payload = payload;
  syncDateFilter();
  render();
}

function syncDateFilter() {
  const previous = els.dateFilter.value || state.date;
  els.dateFilter.replaceChildren();
  els.dateFilter.append(new Option("Все даты", "all"));
  for (const date of state.payload.dates || []) {
    els.dateFilter.append(new Option(date, date));
  }
  const available = new Set(["all", ...(state.payload.dates || [])]);
  state.date = available.has(previous) ? previous : "all";
  els.dateFilter.value = state.date;
}

function updateStats() {
  const stats = state.payload.stats || {};
  els.statNew.textContent = stats.new || 0;
  els.statUnchecked.textContent = stats.unchecked || 0;
  els.statNotRelevant.textContent = stats.notRelevant || 0;
  els.statInWork.textContent = stats.inWork || 0;
  els.generatedAt.textContent = `Обновлено: ${formatDateTime(state.payload.generatedAt)}. Всего: ${stats.total || 0}`;
}

function itemMatchesFilters(item) {
  if (state.date !== "all" && item.entryDate !== state.date) return false;
  if (state.status === "new" && !item.isNew) return false;
  if (state.status === "in_work" && !["analyzing", "submitting"].includes(item.status)) return false;
  if (!["all", "new", "in_work"].includes(state.status) && item.status !== state.status) return false;

  const query = state.search.trim().toLowerCase();
  if (!query) return true;
  const haystack = [
    item.title,
    item.entryDate,
    item.statusLabel,
    ...Object.values(item.columns || {}),
    ...(item.links || []),
    ...(item.files || []).map((file) => file.path),
  ].join(" ").toLowerCase();
  return haystack.includes(query);
}

function groupByDate(items) {
  const groups = new Map();
  for (const item of items) {
    const date = item.entryDate || "без даты";
    if (!groups.has(date)) groups.set(date, []);
    groups.get(date).push(item);
  }
  return [...groups.entries()].sort((a, b) => b[0].localeCompare(a[0]));
}

function render() {
  updateStats();
  const items = (state.payload.items || []).filter(itemMatchesFilters);
  els.content.replaceChildren();
  if (!items.length) {
    els.content.append(make("div", "empty", "Нет закупок под выбранные фильтры"));
    return;
  }
  const groups = groupByDate(items);
  groups.forEach(([date, groupItems], index) => {
    els.content.append(renderDateSection(date, groupItems, index === 0 && state.date === "all"));
  });
}

function renderDateSection(date, items, isMain) {
  const section = make("section", "date-section");
  const head = make("div", "section-head");
  const titleWrap = make("div", "section-title");
  titleWrap.append(make("h2", "", isMain ? `Основная таблица: ${date}` : date));
  titleWrap.append(make("span", "section-count", `${items.length} строк`));
  head.append(titleWrap);
  section.append(head);

  const wrap = make("div", "table-wrap");
  const table = make("table");
  table.append(renderTableHead());
  const body = make("tbody");
  for (const item of items) {
    body.append(renderRow(item));
  }
  table.append(body);
  wrap.append(table);
  section.append(wrap);
  return section;
}

function renderTableHead() {
  const thead = make("thead");
  const tr = make("tr");
  const statusTh = make("th", "sticky-col", "Решение");
  tr.append(statusTh);
  tr.append(make("th", "", "Дата входа"));
  for (const header of state.payload.headers || []) {
    tr.append(make("th", "", header));
  }
  thead.append(tr);
  return thead;
}

function renderRow(item) {
  const tr = make("tr", rowClass(item));
  const statusTd = make("td", "sticky-col");
  statusTd.append(renderStatusCell(item));
  tr.append(statusTd);
  tr.append(textCell(item.entryDate));
  for (const header of state.payload.headers || []) {
    tr.append(renderValueCell(item, header));
  }
  return tr;
}

function rowClass(item) {
  const classes = [];
  if (item.isNew) classes.push("row-new");
  if (item.status === "relevant") classes.push("row-relevant");
  if (item.status === "analyzing") classes.push("row-analyzing");
  if (item.status === "submitting") classes.push("row-submitting");
  if (item.status === "not_relevant") classes.push("row-not_relevant");
  if (state.expandedRows.has(item.id)) classes.push("row-expanded");
  return classes.join(" ");
}

function renderStatusCell(item) {
  const cell = make("div", "status-cell");
  const toggle = make("button", "row-toggle", state.expandedRows.has(item.id) ? "Свернуть" : "Раскрыть");
  toggle.type = "button";
  toggle.dataset.itemId = item.id;
  cell.append(toggle);
  const statuses = state.payload.statuses || {};
  for (const status of statusOrder) {
    const label = make("label", `status-option${item.status === status ? " is-active" : ""}`);
    const input = document.createElement("input");
    input.type = "radio";
    input.name = `status-${item.id}`;
    input.value = status;
    input.dataset.itemId = item.id;
    input.className = "status-radio";
    input.checked = item.status === status;
    label.append(input);
    label.append(document.createTextNode(statuses[status]?.label || status));
    cell.append(label);
  }
  const meta = make("div", "status-meta");
  const badges = [];
  if (item.isNew) badges.push(["badge badge-new", "новая"]);
  if (item.bitrixLeadId) badges.push(["badge badge-bitrix", `Bitrix #${item.bitrixLeadId}`]);
  if (item.filesDeletedAt) badges.push(["badge", `файлы удалены ${formatDateTime(item.filesDeletedAt)}`]);
  if (item.bitrixError) badges.push(["badge badge-error", `Bitrix: ${item.bitrixError}`]);
  for (const [className, text] of badges) {
    meta.append(make("span", className, text));
    meta.append(document.createTextNode(" "));
  }
  if (!badges.length && item.decisionUpdatedAt) {
    meta.textContent = `обновлено ${formatDateTime(item.decisionUpdatedAt)}`;
  }
  cell.append(meta);
  return cell;
}

function renderValueCell(item, header) {
  const rawValue = item.columns?.[header] || "";
  if (header === "Тема" || header === "Объект" || header === "Название объекта") {
    const td = textCell(rawValue);
    td.classList.add("row-title");
    return td;
  }
  if (header === "Папка материалов") {
    return linkCell(item.folder ? [item.folder] : []);
  }
  if (header.toLowerCase().includes("папка") && looksLikeLocalPath(rawValue)) {
    return linkCell([localPathLink(rawValue, "папка материалов")]);
  }
  if (header.toLowerCase().includes("excel") && looksLikeLocalPath(rawValue)) {
    return linkCell([localPathLink(rawValue, "excel-отчет")]);
  }
  if (header === "Исходные файлы") {
    if (item.compact) return textCell("Файлы удалены");
    return linkCell(item.files || []);
  }
  if (header === "Ссылки" || header === "Ссылки в интернете") {
    const links = (item.links || []).map((url, index) => ({
      name: index === 0 ? "ссылка на закупку" : `ссылка на закупку ${index + 1}`,
      url,
      path: "Открыть закупку",
    }));
    return linkCell(links);
  }
  if (extractUrls(rawValue).length) {
    return linkCell(extractUrls(rawValue).map((url, index) => ({
      name: index === 0 ? "ссылка на закупку" : `ссылка на закупку ${index + 1}`,
      url,
      path: "Открыть закупку",
    })));
  }
  return textCell(rawValue);
}

function extractUrls(value) {
  const matches = String(value || "").match(/https?:\/\/[^\s;]+/g);
  return matches ? [...new Set(matches.map((url) => url.replace(/[.,)]$/, "")))] : [];
}

function looksLikeLocalPath(value) {
  return /^[a-z]:\\/i.test(value || "") || /^\\\\/.test(value || "");
}

function pathToFileUrl(value) {
  const path = String(value || "");
  if (path.startsWith("\\\\")) {
    return `file://///${path.replace(/^\\\\/, "").replaceAll("\\", "/")}`;
  }
  return `file:///${path.replaceAll("\\", "/")}`;
}

function pathToDashboardUrl(value, kind = "file") {
  const path = String(value || "");
  if (state.payload?.fileLinkMode === "download") {
    const route = kind === "folder" ? "/api/browse" : "/api/download";
    return `${route}?path=${encodeURIComponent(path)}`;
  }
  return pathToFileUrl(path);
}

function localPathLink(value, name) {
  const isFolder = String(name || "").toLowerCase().includes("папка");
  return {
    name,
    url: pathToDashboardUrl(value, isFolder ? "folder" : "file"),
    path: value,
    local: state.payload?.fileLinkMode !== "download",
    download: state.payload?.fileLinkMode === "download" && !isFolder,
  };
}

function textCell(value) {
  const td = make("td");
  const text = value || "—";
  const content = make("div", "cell-content", text);
  td.append(content);
  td.title = text;
  if (!value) td.classList.add("cell-muted");
  return td;
}

function linkCell(links) {
  const td = make("td");
  if (!links.length) {
    td.append(make("div", "cell-content", "—"));
    td.classList.add("cell-muted");
    return td;
  }
  const content = make("div", "cell-content");
  const wrap = make("div", "link-list");
  const limit = 12;
  links.slice(0, limit).forEach((link) => {
    const anchor = document.createElement("a");
    const openPath = !link.download && (link.local || looksLikeLocalPath(link.path)) ? link.path : "";
    anchor.href = link.url || link.path;
    anchor.textContent = link.name || link.path || link.url;
    anchor.title = link.path || link.url;
    if (openPath) {
      anchor.dataset.openPath = openPath;
      anchor.classList.add("local-open-link");
    } else {
      anchor.target = "_blank";
      anchor.rel = "noreferrer";
    }
    wrap.append(anchor);
  });
  if (links.length > limit) {
    wrap.append(make("span", "cell-muted", `еще ${links.length - limit}`));
  }
  content.append(wrap);
  td.append(content);
  return td;
}

async function updateStatus(itemId, status) {
  const payload = await fetchJson(`/api/items/${encodeURIComponent(itemId)}/status`, {
    method: "POST",
    body: JSON.stringify({ status }),
  });
  state.payload = payload;
  syncDateFilter();
  render();
  if (payload.bitrix?.status === "sent") {
    showToast(`Лид Bitrix создан: #${payload.bitrix.leadId}`);
  } else if (payload.bitrix?.status === "not_configured") {
    showToast("Статус сохранен. Bitrix webhook не задан в конфиге.");
  } else if (payload.bitrix?.status === "error") {
    showToast(`Bitrix не принял лид: ${payload.bitrix.message}`);
  }
}

els.content.addEventListener("change", async (event) => {
  const input = event.target;
  if (!input.classList.contains("status-radio")) return;
  input.disabled = true;
  try {
    await updateStatus(input.dataset.itemId, input.value);
  } catch (error) {
    showToast(error.message);
    await loadData();
  } finally {
    input.disabled = false;
  }
});

els.content.addEventListener("click", (event) => {
  const button = event.target;
  if (!button.classList.contains("row-toggle")) return;
  const itemId = button.dataset.itemId;
  if (state.expandedRows.has(itemId)) {
    state.expandedRows.delete(itemId);
  } else {
    state.expandedRows.add(itemId);
  }
  render();
});

els.content.addEventListener("click", async (event) => {
  const anchor = event.target.closest?.("a[data-open-path]");
  if (!anchor) return;
  event.preventDefault();
  try {
    await fetchJson("/api/open-path", {
      method: "POST",
      body: JSON.stringify({ path: anchor.dataset.openPath }),
    });
    showToast("Открываю файл или папку");
  } catch (error) {
    showToast(error.message);
    if (anchor.href && anchor.href !== "#") {
      window.open(anchor.href, "_blank", "noopener");
    }
  }
});

els.searchInput.addEventListener("input", () => {
  state.search = els.searchInput.value;
  render();
});

els.dateFilter.addEventListener("change", () => {
  state.date = els.dateFilter.value;
  render();
});

els.statusFilter.addEventListener("change", () => {
  state.status = els.statusFilter.value;
  render();
});

els.refreshBtn.addEventListener("click", async () => {
  try {
    await loadData();
    showToast("Dashboard обновлен");
  } catch (error) {
    showToast(error.message);
  }
});

els.cleanupBtn.addEventListener("click", async () => {
  const confirmed = window.confirm("Удалить скачанные файлы старше срока хранения для закупок без статуса «Подаемся»?");
  if (!confirmed) return;
  try {
    const result = await fetchJson("/api/cleanup", { method: "POST", body: "{}" });
    await loadData();
    showToast(`Удалено файлов: ${result.deletedCount}`);
  } catch (error) {
    showToast(error.message);
  }
});

loadData().catch((error) => {
  els.content.replaceChildren(make("div", "empty", error.message));
  els.generatedAt.textContent = "Ошибка загрузки";
});
