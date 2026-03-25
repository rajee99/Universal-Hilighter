
"use strict";

const PRESET_COLORS = {
  yellow: { solid: "#ffd42e" },
  green: { solid: "#33c3ac" },
  blue: { solid: "#4f9fe8" },
  pink: { solid: "#ff7eb6" },
  orange: { solid: "#ff9f43" }
};

const state = {
  enabled: true,
  highlights: {},
  favorites: [],
  categories: {},
  domainAssignments: {},
  domainAssignmentCutoffs: {},
  notebookItems: [],
  highlightAttachments: {},
  notebookObjectUrls: new Map(),
  timeline: [],
  view: "all",
  search: "",
  colorFilter: "all",
  domainFilter: "all",
  categoryFilter: "all",
  sortOrder: "date_desc"
};

const refs = {
  statusBadge: document.getElementById("statusBadge"),
  statusText: document.getElementById("statusText"),
  powerButtonText: document.getElementById("powerButtonText"),
  btnUploadNotebookFile: document.getElementById("btnUploadNotebookFile"),
  btnDashboardPower: document.getElementById("btnDashboardPower"),
  btnToggleFullscreen: document.getElementById("btnToggleFullscreen"),
  btnRefreshDashboard: document.getElementById("btnRefreshDashboard"),
  searchInput: document.getElementById("searchInput"),
  domainChips: document.getElementById("domainChips"),
  colorFilter: document.getElementById("colorFilter"),
  sortSelect: document.getElementById("sortSelect"),
  categoryFilter: document.getElementById("categoryFilter"),
  categoryList: document.getElementById("categoryList"),
  dashboardContent: document.getElementById("dashboardContent"),
  summaryText: document.getElementById("summaryText"),
  selectionInfo: document.getElementById("selectionInfo"),
  emptyState: document.getElementById("emptyState"),
  navButtons: Array.from(document.querySelectorAll(".hmdb-nav-btn")),
  btnNewCategory: document.getElementById("btnNewCategory"),
  btnNewNotebookNote: document.getElementById("btnNewNotebookNote"),
  btnExport: document.getElementById("btnExport"),
  btnImport: document.getElementById("btnImport"),
  btnRestore: document.getElementById("btnRestore"),
  importFile: document.getElementById("importFile"),
  notebookFileInput: document.getElementById("notebookFileInput"),
  highlightFileInput: document.getElementById("highlightFileInput"),
  dashboardVersion: document.getElementById("dashboardVersion"),
  noteModal: document.getElementById("noteModal"),
  noteEditor: document.getElementById("noteEditor"),
  noteContext: document.getElementById("noteContext"),
  noteSaveState: document.getElementById("noteSaveState"),
  btnCloseNoteModal: document.getElementById("btnCloseNoteModal"),
  toast: document.getElementById("toast")
};

let toastTimer = null;
let searchDebounce = null;
let activeNoteTarget = null;
let noteSaveTimer = null;
let lastSavedNoteValue = "";
let realtimeRefreshTimer = null;
let pendingRealtimeRefresh = false;
let pendingAttachmentTarget = null;
let realtimeRefreshInFlight = false;

const NOTEBOOK_META_KEY = "dashboardNotebookItemsV1";
const HIGHLIGHT_ATTACHMENT_META_KEY = "dashboardHighlightAttachmentsV1";
const NOTEBOOK_DB_NAME = "highlightMasterDashboardFiles";
const NOTEBOOK_DB_VERSION = 1;
const NOTEBOOK_STORE = "files";

function sendMsg(msg) {
  return new Promise((resolve) => {
    chrome.runtime.sendMessage(msg, (res) => {
      if (chrome.runtime.lastError) {
        resolve({ ok: false, error: chrome.runtime.lastError.message });
        return;
      }
      resolve(res || {});
    });
  });
}

function esc(str) {
  const div = document.createElement("div");
  div.textContent = String(str == null ? "" : str);
  return div.innerHTML;
}

function toast(message) {
  if (!refs.toast) return;
  refs.toast.textContent = message;
  refs.toast.classList.add("visible");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => refs.toast.classList.remove("visible"), 2200);
}

function normalizeId(id) { return id == null ? "" : String(id).trim(); }

function normalizeHighlightsMap(raw) {
  const out = {};
  for (const [domain, list] of Object.entries(raw || {})) {
    if (!Array.isArray(list)) continue;
    const seen = new Set();
    out[domain] = [];
    for (const item of list) {
      const id = normalizeId(item && item.id);
      if (!id || seen.has(id)) continue;
      seen.add(id);
      out[domain].push({ ...item, id, domain: item.domain || domain });
    }
  }
  return out;
}

function getItemLabel(item) {
  if (!item) return "";
  if (item.type === "shape-cover") return "Cover Rectangle";
  if (item.type === "shape-rect") return item.label || item.ocrText || item.text || "Rectangle Highlight";
  return item.text || "Text highlight";
}

function getColorBucket(colorName) {
  if (!colorName) return "yellow";
  if (String(colorName).startsWith("custom:")) return "custom";
  return PRESET_COLORS[colorName] ? colorName : "custom";
}

function flatHighlights() {
  const list = [];
  for (const [domain, items] of Object.entries(state.highlights)) {
    for (const item of (items || [])) list.push({ ...item, domain: item.domain || domain });
  }
  return list;
}

function assignInfo(domain) {
  const categoryId = typeof state.domainAssignments[domain] === "string" ? state.domainAssignments[domain] : "";
  const cutoffRaw = Number(state.domainAssignmentCutoffs[domain]);
  const cutoff = Number.isFinite(cutoffRaw) && cutoffRaw > 0 ? cutoffRaw : 0;
  return { categoryId, cutoff };
}

function categoryIdFor(item) {
  const info = assignInfo(item && item.domain ? item.domain : "");
  if (!info.categoryId || !state.categories[info.categoryId] || !info.cutoff) return "";
  if ((Number(item && item.timestamp) || 0) > info.cutoff) return "";
  return info.categoryId;
}

function categoryLabelFor(item) {
  const id = categoryIdFor(item);
  return id && state.categories[id] ? (state.categories[id].name || "Category") : "Uncategorized";
}

function formatDateTime(ts) {
  if (!ts) return "Unknown time";
  try {
    return new Date(ts).toLocaleString(undefined, { month: "short", day: "numeric", hour: "2-digit", minute: "2-digit" });
  } catch (_err) {
    return "Unknown time";
  }
}

function shortText(str, max = 220) {
  const clean = String(str || "").replace(/\s+/g, " ").trim();
  return clean.length > max ? clean.slice(0, max - 3).trimEnd() + "..." : clean;
}

function formatFileSize(size) {
  const value = Number(size) || 0;
  if (value < 1024) return value + " B";
  if (value < 1024 * 1024) return (value / 1024).toFixed(1).replace(/\.0$/, "") + " KB";
  if (value < 1024 * 1024 * 1024) return (value / (1024 * 1024)).toFixed(1).replace(/\.0$/, "") + " MB";
  return (value / (1024 * 1024 * 1024)).toFixed(1).replace(/\.0$/, "") + " GB";
}

function storageLocalGet(keys) {
  return new Promise((resolve) => {
    chrome.storage.local.get(keys, (result) => {
      if (chrome.runtime.lastError) {
        resolve({});
        return;
      }
      resolve(result || {});
    });
  });
}

function storageLocalSet(payload) {
  return new Promise((resolve) => {
    chrome.storage.local.set(payload, () => resolve());
  });
}

function openNotebookDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(NOTEBOOK_DB_NAME, NOTEBOOK_DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(NOTEBOOK_STORE)) {
        db.createObjectStore(NOTEBOOK_STORE);
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error || new Error("indexeddb_open_failed"));
  });
}

async function idbPutFile(id, fileBlob) {
  const db = await openNotebookDb();
  await new Promise((resolve, reject) => {
    const tx = db.transaction(NOTEBOOK_STORE, "readwrite");
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error || new Error("indexeddb_put_failed"));
    tx.objectStore(NOTEBOOK_STORE).put(fileBlob, id);
  });
  db.close();
}

async function idbGetFile(id) {
  const db = await openNotebookDb();
  const blob = await new Promise((resolve, reject) => {
    const tx = db.transaction(NOTEBOOK_STORE, "readonly");
    tx.onerror = () => reject(tx.error || new Error("indexeddb_get_failed"));
    const req = tx.objectStore(NOTEBOOK_STORE).get(id);
    req.onsuccess = () => resolve(req.result || null);
    req.onerror = () => reject(req.error || new Error("indexeddb_get_failed"));
  });
  db.close();
  return blob;
}

async function idbDeleteFile(id) {
  const db = await openNotebookDb();
  await new Promise((resolve, reject) => {
    const tx = db.transaction(NOTEBOOK_STORE, "readwrite");
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error || new Error("indexeddb_delete_failed"));
    tx.objectStore(NOTEBOOK_STORE).delete(id);
  });
  db.close();
}

function sanitizeNotebookItems(rawItems) {
  if (!Array.isArray(rawItems)) return [];
  return rawItems
    .filter((item) => item && typeof item === "object" && normalizeId(item.id))
    .map((item) => ({
      id: normalizeId(item.id),
      kind: item.kind === "text-note" ? "text-note" : "file",
      name: String(item.name || (item.kind === "text-note" ? "Notebook Note" : "Untitled File")),
      mimeType: String(item.mimeType || ""),
      size: Number(item.size) || 0,
      createdAt: Number(item.createdAt) || Date.now(),
      updatedAt: Number(item.updatedAt) || Number(item.createdAt) || Date.now(),
      note: String(item.note || ""),
      emoji: String(item.emoji || "")
    }))
    .sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0));
}

async function loadNotebookItems() {
  const stored = await storageLocalGet([NOTEBOOK_META_KEY]);
  state.notebookItems = sanitizeNotebookItems(stored[NOTEBOOK_META_KEY] || []);
}

async function saveNotebookItems() {
  await storageLocalSet({ [NOTEBOOK_META_KEY]: state.notebookItems });
}

function makeNotebookId(prefix) {
  return prefix + "_" + Date.now() + "_" + Math.random().toString(36).slice(2, 8);
}

function ownerKeyFor(domain, id) {
  return normalizeId(domain) + "::" + normalizeId(id);
}

function parseOwnerKey(ownerKey) {
  const raw = String(ownerKey || "");
  const sepIndex = raw.indexOf("::");
  if (sepIndex <= 0) return null;
  const domain = normalizeId(raw.slice(0, sepIndex));
  const id = normalizeId(raw.slice(sepIndex + 2));
  if (!domain || !id) return null;
  return { ownerKey: domain + "::" + id, domain, id };
}

function sanitizeHighlightAttachmentItems(items, ownerKey) {
  if (!Array.isArray(items)) return [];
  const owner = parseOwnerKey(ownerKey);
  if (!owner) return [];
  const seen = new Set();
  const out = [];
  for (const item of items) {
    const id = normalizeId(item && item.id);
    if (!id || seen.has(id)) continue;
    seen.add(id);
    out.push({
      id,
      ownerKey: owner.ownerKey,
      ownerDomain: owner.domain,
      ownerId: owner.id,
      name: String(item && item.name ? item.name : "Attachment"),
      mimeType: String(item && item.mimeType ? item.mimeType : ""),
      size: Number(item && item.size) || 0,
      createdAt: Number(item && item.createdAt) || Date.now(),
      updatedAt: Number(item && item.updatedAt) || Number(item && item.createdAt) || Date.now()
    });
  }
  out.sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0));
  return out;
}

function sanitizeHighlightAttachmentMap(rawMap) {
  const out = {};
  if (!rawMap || typeof rawMap !== "object") return out;
  for (const [ownerKey, items] of Object.entries(rawMap)) {
    const cleanItems = sanitizeHighlightAttachmentItems(items, ownerKey);
    if (cleanItems.length > 0) {
      out[parseOwnerKey(ownerKey).ownerKey] = cleanItems;
    }
  }
  return out;
}

async function loadHighlightAttachments() {
  const stored = await storageLocalGet([HIGHLIGHT_ATTACHMENT_META_KEY]);
  state.highlightAttachments = sanitizeHighlightAttachmentMap(stored[HIGHLIGHT_ATTACHMENT_META_KEY] || {});
}

async function saveHighlightAttachments() {
  await storageLocalSet({ [HIGHLIGHT_ATTACHMENT_META_KEY]: state.highlightAttachments || {} });
}

function isImageMime(type) {
  const value = String(type || "").toLowerCase();
  return value.startsWith("image/");
}

function isPdfMime(type, name) {
  const t = String(type || "").toLowerCase();
  const n = String(name || "").toLowerCase();
  return t === "application/pdf" || n.endsWith(".pdf");
}

function matchesNotebookSearch(item) {
  if (!state.search) return true;
  const hay = [
    item.name || "",
    item.note || "",
    item.mimeType || "",
    item.emoji || "",
    item.kind === "text-note" ? "note notebook memo emoji" : "file attachment upload"
  ].join(" ").toLowerCase();
  return hay.includes(state.search);
}

function getFilteredNotebookItems() {
  return state.notebookItems.filter(matchesNotebookSearch);
}

function clearNotebookObjectUrls() {
  if (!(state.notebookObjectUrls instanceof Map)) return;
  for (const url of state.notebookObjectUrls.values()) {
    try { URL.revokeObjectURL(url); } catch (_err) {}
  }
  state.notebookObjectUrls.clear();
}
function setStatusUi() {
  refs.statusText.textContent = state.enabled ? "Active" : "Paused";
  refs.powerButtonText.textContent = state.enabled ? "Pause System" : "Enable System";
  refs.statusBadge.classList.toggle("offline", !state.enabled);
}

function setVersionUi() {
  if (!chrome.runtime || !chrome.runtime.getManifest) return;
  refs.dashboardVersion.textContent = "HighlightMaster v" + (chrome.runtime.getManifest().version || "-");
}

function compareItems(a, b) {
  const aLabel = getItemLabel(a).toLowerCase();
  const bLabel = getItemLabel(b).toLowerCase();
  if (state.sortOrder === "date_asc") return (a.timestamp || 0) - (b.timestamp || 0);
  if (state.sortOrder === "alpha_asc") return aLabel.localeCompare(bLabel) || ((b.timestamp || 0) - (a.timestamp || 0));
  if (state.sortOrder === "alpha_desc") return bLabel.localeCompare(aLabel) || ((b.timestamp || 0) - (a.timestamp || 0));
  return (b.timestamp || 0) - (a.timestamp || 0);
}

function matchesSearch(item) {
  if (!state.search) return true;
  const attachmentNames = getHighlightAttachmentsForItem(item)
    .map((attachment) => attachment && attachment.name ? attachment.name : "")
    .join(" ");
  const hay = [
    getItemLabel(item),
    item.note || "",
    item.domain || "",
    item.pageTitle || "",
    item.ocrText || "",
    categoryLabelFor(item),
    attachmentNames
  ].join(" ").toLowerCase();
  return hay.includes(state.search);
}

function matchesCategory(item) {
  if (state.categoryFilter === "all") return true;
  if (state.categoryFilter === "uncategorized") return !categoryIdFor(item);
  if (state.categoryFilter === "with_notes") return String(item.note || "").trim().length > 0;
  if (state.categoryFilter.startsWith("cat:")) return categoryIdFor(item) === state.categoryFilter.slice(4);
  return true;
}

function matchesFilters(item) {
  if (state.view === "pinned" && !state.favorites.includes(item.id)) return false;
  if (state.domainFilter !== "all" && item.domain !== state.domainFilter) return false;
  if (state.colorFilter !== "all" && getColorBucket(item.color) !== state.colorFilter) return false;
  if (!matchesCategory(item)) return false;
  if (!matchesSearch(item)) return false;
  return true;
}

function buildDomainChips() {
  const counts = {};
  for (const item of flatHighlights()) counts[item.domain] = (counts[item.domain] || 0) + 1;
  const names = Object.keys(counts).sort((a, b) => (counts[b] - counts[a]) || a.localeCompare(b));

  const html = ['<button class="hmdb-chip' + (state.domainFilter === "all" ? ' active' : '') + '" data-domain="all">All Sources</button>'];
  for (const domain of names) {
    html.push('<button class="hmdb-chip' + (state.domainFilter === domain ? ' active' : '') + '" data-domain="' + esc(domain) + '">' + esc(domain) + '</button>');
  }
  refs.domainChips.innerHTML = html.join("");
}

function buildCategoryUi() {
  const options = [
    { value: "all", label: "All categories" },
    { value: "uncategorized", label: "Uncategorized" },
    { value: "with_notes", label: "With notes" }
  ];
  const roots = Object.entries(state.categories).filter(([, c]) => c && !c.parentId).sort((a, b) => String(a[1].name || "").localeCompare(String(b[1].name || "")));
  for (const [id, cat] of roots) options.push({ value: "cat:" + id, label: cat.name || "Category" });
  refs.categoryFilter.innerHTML = options.map((opt) => '<option value="' + esc(opt.value) + '">' + esc(opt.label) + '</option>').join("");
  if (!options.some((opt) => opt.value === state.categoryFilter)) state.categoryFilter = "all";
  refs.categoryFilter.value = state.categoryFilter;

  const side = [];
  side.push('<button class="hmdb-category-btn' + (state.categoryFilter === "all" ? ' active' : '') + '" data-cat-filter="all"><span>All categories</span><span class="hmdb-category-count">' + flatHighlights().length + '</span></button>');
  for (const [id, cat] of roots) {
    let count = 0;
    for (const item of flatHighlights()) if (categoryIdFor(item) === id) count += 1;
    side.push('<button class="hmdb-category-btn' + (state.categoryFilter === ("cat:" + id) ? ' active' : '') + '" data-cat-filter="cat:' + esc(id) + '"><span>' + esc(cat.name || "Category") + '</span><span class="hmdb-category-count">' + count + '</span></button>');
  }
  refs.categoryList.innerHTML = side.join("");
}

function notebookFileExt(name) {
  const value = String(name || "");
  const idx = value.lastIndexOf(".");
  if (idx <= 0 || idx === value.length - 1) return "FILE";
  return value.slice(idx + 1, idx + 5).toUpperCase();
}

function renderHighlightAttachmentCells(item) {
  const ownerKey = getHighlightOwnerKey(item);
  if (!ownerKey) return "";
  const attachments = getHighlightAttachmentsForItem(item);
  if (!attachments.length) return "";
  const html = [];
  html.push('<div class="hmdb-attachment-grid">');
  for (const attachment of attachments) {
    const ext = notebookFileExt(attachment.name);
    const fileMeta = (attachment.mimeType ? attachment.mimeType : "file") + " \u2022 " + formatFileSize(attachment.size || 0);
    html.push('<article class="hmdb-attachment-cell" data-ha-id="' + esc(attachment.id) + '">');
    html.push('<div class="hmdb-attachment-preview" data-ha-preview="' + esc(attachment.id) + '" data-ha-owner="' + esc(ownerKey) + '">');
    html.push('<span class="hmdb-attachment-preview-icon">' + esc(ext) + '</span>');
    html.push('</div>');
    html.push('<div class="hmdb-attachment-main">');
    html.push('<div class="hmdb-attachment-copy">');
    html.push('<button class="hmdb-attachment-open" data-ha-action="open" data-ha-id="' + esc(attachment.id) + '" data-ha-owner="' + esc(ownerKey) + '" title="Open attachment">' + esc(shortText(attachment.name || "Attachment", 54)) + '</button>');
    html.push('<span class="hmdb-attachment-meta">' + esc(fileMeta) + '</span>');
    html.push('</div>');
    html.push('</div>');
    html.push('<div class="hmdb-attachment-actions">');
    html.push('<button class="hmdb-card-btn mini" data-ha-action="download" data-ha-id="' + esc(attachment.id) + '" data-ha-owner="' + esc(ownerKey) + '" title="Download"><svg viewBox="0 0 24 24"><path d="M12 3v12"/><path d="m7 10 5 5 5-5"/><path d="M4 21h16"/></svg></button>');
    html.push('<button class="hmdb-card-btn mini danger" data-ha-action="delete" data-ha-id="' + esc(attachment.id) + '" data-ha-owner="' + esc(ownerKey) + '" title="Delete"><svg viewBox="0 0 24 24"><path d="M18 6 6 18M6 6l12 12"/></svg></button>');
    html.push('</div>');
    html.push('</article>');
  }
  html.push('</div>');
  return html.join("");
}

function renderNotebookCardsHtml(items) {
  if (!items.length) return "";
  const html = [];
  html.push("<section class=\"hmdb-notebook-block\">");
  html.push("<header class=\"hmdb-notebook-head\"><h3>Recent Uploads</h3><span class=\"hmdb-notebook-meta\">" + items.length + " item" + (items.length === 1 ? "" : "s") + " in notebook</span></header>");
  html.push("<div class=\"hmdb-notebook-grid\">");
  for (const item of items) {
    const note = shortText(item.note || "", 180);
    if (item.kind === "text-note") {
      html.push("<article class=\"hmdb-file-card hmdb-note-card\" data-notebook-id=\"" + esc(item.id) + "\">");
      html.push("<div class=\"hmdb-file-body\">");
      html.push("<p class=\"hmdb-file-name\">" + esc((item.emoji || "") + " " + item.name) + "</p>");
      html.push("<div class=\"hmdb-file-meta\">Notebook note \u2022 " + esc(formatDateTime(item.updatedAt || item.createdAt)) + "</div>");
      html.push("<div class=\"hmdb-file-note\">" + esc(note || "(empty note)") + "</div>");
      html.push("<div class=\"hmdb-file-actions\">");
      html.push("<button class=\"hmdb-card-btn note\" data-nb-action=\"edit-note\" data-nb-id=\"" + esc(item.id) + "\" title=\"Edit note\"><svg viewBox=\"0 0 24 24\"><path d=\"M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z\"/></svg><span>Edit</span></button>");
      html.push("<button class=\"hmdb-card-btn danger\" data-nb-action=\"delete\" data-nb-id=\"" + esc(item.id) + "\" title=\"Delete note\"><svg viewBox=\"0 0 24 24\"><path d=\"M18 6 6 18M6 6l12 12\"/></svg></button>");
      html.push("</div></div></article>");
      continue;
    }

    const ext = notebookFileExt(item.name);
    const fileMeta = (item.mimeType ? item.mimeType : "file") + " \u2022 " + formatFileSize(item.size || 0);
    html.push("<article class=\"hmdb-file-card\" data-notebook-id=\"" + esc(item.id) + "\">");
    html.push("<div class=\"hmdb-file-preview\" data-nb-preview=\"" + esc(item.id) + "\">");
    html.push("<span class=\"hmdb-file-preview-icon\">" + esc(ext) + "</span>");
    html.push("</div>");
    html.push("<div class=\"hmdb-file-body\">");
    html.push("<p class=\"hmdb-file-name\">" + esc(shortText(item.name || "Attachment", 90)) + "</p>");
    html.push("<div class=\"hmdb-file-meta\">" + esc(fileMeta) + "</div>");
    if (note) html.push("<div class=\"hmdb-file-note\">" + esc(note) + "</div>");
    html.push("<div class=\"hmdb-file-actions\">");
    html.push("<button class=\"hmdb-card-btn\" data-nb-action=\"open\" data-nb-id=\"" + esc(item.id) + "\" title=\"Open file\"><svg viewBox=\"0 0 24 24\"><path d=\"M7 17 17 7\"/><path d=\"M7 7h10v10\"/></svg></button>");
    html.push("<button class=\"hmdb-card-btn\" data-nb-action=\"download\" data-nb-id=\"" + esc(item.id) + "\" title=\"Download file\"><svg viewBox=\"0 0 24 24\"><path d=\"M12 3v12\"/><path d=\"m7 10 5 5 5-5\"/><path d=\"M4 21h16\"/></svg></button>");
    html.push("<button class=\"hmdb-card-btn note\" data-nb-action=\"edit-note\" data-nb-id=\"" + esc(item.id) + "\" title=\"Edit note\"><svg viewBox=\"0 0 24 24\"><path d=\"M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z\"/></svg><span>Note</span></button>");
    html.push("<button class=\"hmdb-card-btn danger\" data-nb-action=\"delete\" data-nb-id=\"" + esc(item.id) + "\" title=\"Delete file\"><svg viewBox=\"0 0 24 24\"><path d=\"M18 6 6 18M6 6l12 12\"/></svg></button>");
    html.push("</div></div></article>");
  }
  html.push("</div></section>");
  return html.join("");
}

async function hydrateNotebookPreviews() {
  const nodes = Array.from(document.querySelectorAll("[data-nb-preview]"));
  for (const node of nodes) {
    const itemId = node.getAttribute("data-nb-preview") || "";
    const item = state.notebookItems.find((n) => n.id === itemId);
    if (!item || item.kind !== "file") continue;
    const canShowImage = isImageMime(item.mimeType);
    const canShowPdf = isPdfMime(item.mimeType, item.name);
    if (!canShowImage && !canShowPdf) continue;
    try {
      let blobUrl = state.notebookObjectUrls.get(itemId);
      if (!blobUrl) {
        const blob = await idbGetFile(itemId);
        if (!blob) continue;
        blobUrl = URL.createObjectURL(blob);
        state.notebookObjectUrls.set(itemId, blobUrl);
      }
      if (canShowImage) {
        node.innerHTML = "<img src=\"" + esc(blobUrl) + "\" alt=\"" + esc(item.name || "Attachment preview") + "\">";
      } else {
        const pdfUrl = blobUrl + "#toolbar=0&navpanes=0&scrollbar=0&page=1&view=FitH";
        node.innerHTML = "<iframe src=\"" + esc(pdfUrl) + "\" title=\"" + esc(item.name || "PDF preview") + "\" loading=\"lazy\"></iframe>";
      }
    } catch (_err) {}
  }
}

async function hydrateHighlightAttachmentPreviews() {
  const nodes = Array.from(document.querySelectorAll("[data-ha-preview][data-ha-owner]"));
  for (const node of nodes) {
    const attachmentId = node.getAttribute("data-ha-preview") || "";
    const ownerKey = node.getAttribute("data-ha-owner") || "";
    const found = findHighlightAttachment(attachmentId, ownerKey);
    if (!found || !found.item) continue;
    const item = found.item;
    const canShowImage = isImageMime(item.mimeType);
    const canShowPdf = isPdfMime(item.mimeType, item.name);
    if (!canShowImage && !canShowPdf) continue;
    try {
      let blobUrl = state.notebookObjectUrls.get(item.id);
      if (!blobUrl) {
        const blob = await idbGetFile(item.id);
        if (!blob) continue;
        blobUrl = URL.createObjectURL(blob);
        state.notebookObjectUrls.set(item.id, blobUrl);
      }
      if (canShowImage) {
        node.innerHTML = "<img src=\"" + esc(blobUrl) + "\" alt=\"" + esc(item.name || "Attachment preview") + "\">";
      } else {
        const pdfUrl = blobUrl + "#toolbar=0&navpanes=0&scrollbar=0&page=1&view=FitH";
        node.innerHTML = "<iframe src=\"" + esc(pdfUrl) + "\" title=\"" + esc(item.name || "PDF preview") + "\" loading=\"lazy\"></iframe>";
      }
    } catch (_err) {}
  }
}

function renderNotebookOnly() {
  const items = getFilteredNotebookItems();
  if (!items.length) {
    refs.dashboardContent.innerHTML = "";
    refs.emptyState.hidden = false;
    return;
  }
  refs.emptyState.hidden = true;
  refs.dashboardContent.innerHTML = renderNotebookCardsHtml(items);
  hydrateNotebookPreviews();
  hydrateHighlightAttachmentPreviews();
}

function renderHighlights() {
  const items = flatHighlights().filter(matchesFilters).sort(compareItems);
  const notebookItems = state.view === "all" ? getFilteredNotebookItems() : [];

  if (!items.length && !notebookItems.length) {
    refs.dashboardContent.innerHTML = "";
    refs.emptyState.hidden = false;
    return;
  }
  refs.emptyState.hidden = true;

  const groups = new Map();
  for (const item of items) {
    if (!groups.has(item.domain)) groups.set(item.domain, []);
    groups.get(item.domain).push(item);
  }
  const names = Array.from(groups.keys()).sort((a, b) => (groups.get(b).length - groups.get(a).length) || a.localeCompare(b));

  const html = [];
  if (notebookItems.length) {
    html.push(renderNotebookCardsHtml(notebookItems));
  }
  for (const domain of names) {
    const list = groups.get(domain);
    html.push('<section class="hmdb-domain-block"><header class="hmdb-domain-head"><h3 class="hmdb-domain-name">' + esc(domain) + '</h3><span class="hmdb-domain-count">' + list.length + ' items</span></header><div class="hmdb-grid">');
    for (const item of list) {
      const color = getColorBucket(item.color);
      const text = shortText(getItemLabel(item), 260);
      const note = shortText(item.note || "", 170);
      const attachmentCount = getHighlightAttachmentsForItem(item).length;
      const pinned = state.favorites.includes(item.id);
      const kind = item.type === "shape-rect" ? "Rectangle" : (item.type === "shape-cover" ? "Cover" : "Highlight");
      html.push('<article class="hmdb-card" data-color="' + color + '" data-domain="' + esc(item.domain) + '" data-id="' + esc(item.id) + '">');
      html.push('<p class="hmdb-card-text">' + esc(text) + '</p>');
      if (note) html.push('<div class="hmdb-card-note">' + esc(note) + '</div>');
      html.push('<div class="hmdb-meta-row"><span class="hmdb-badge">' + esc(kind) + '</span><span class="hmdb-meta">' + esc(categoryLabelFor(item)) + '</span><span class="hmdb-meta">' + esc(formatDateTime(item.timestamp)) + '</span><span class="hmdb-meta">' + attachmentCount + ' attachment' + (attachmentCount === 1 ? '' : 's') + '</span></div>');
      html.push('<div class="hmdb-card-actions">');
      html.push('<button class="hmdb-card-btn" data-action="copy" data-domain="' + esc(item.domain) + '" data-id="' + esc(item.id) + '" title="Copy"><svg viewBox="0 0 24 24"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg></button>');
      html.push('<button class="hmdb-card-btn' + (pinned ? ' active' : '') + '" data-action="pin" data-domain="' + esc(item.domain) + '" data-id="' + esc(item.id) + '" title="Pin"><svg viewBox="0 0 24 24"><path d="M12 17v5"/><path d="M8 10V3h10.2a1.8 1.8 0 0 1 1.8 1.8v8.8L17 11l-3 3-2-2-2 2z"/></svg></button>');
      html.push('<button class="hmdb-card-btn" data-action="attach" data-domain="' + esc(item.domain) + '" data-id="' + esc(item.id) + '" title="Attach files"><svg viewBox="0 0 24 24"><path d="M21.44 11.05 12.25 20.24a6 6 0 0 1-8.49-8.49l9.19-9.19a4 4 0 0 1 5.66 5.66l-9.2 9.19a2 2 0 0 1-2.83-2.83l8.49-8.48"/></svg></button>');
      html.push('<button class="hmdb-card-btn note" data-action="note" data-domain="' + esc(item.domain) + '" data-id="' + esc(item.id) + '" title="Note"><svg viewBox="0 0 24 24"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg><span>Note</span></button>');
      html.push('<button class="hmdb-card-btn danger" data-action="delete" data-domain="' + esc(item.domain) + '" data-id="' + esc(item.id) + '" title="Delete"><svg viewBox="0 0 24 24"><path d="M18 6 6 18M6 6l12 12"/></svg></button>');
      html.push('</div>');
      html.push(renderHighlightAttachmentCells(item));
      html.push('</article>');
    }
    html.push('</div></section>');
  }
  refs.dashboardContent.innerHTML = html.join("");
  hydrateNotebookPreviews();
  hydrateHighlightAttachmentPreviews();
}

function renderTimeline() {
  let rows = (state.timeline || []).filter((entry) => {
    if (state.domainFilter !== "all") {
      const changed = Array.isArray(entry.changedDomains) ? entry.changedDomains : [];
      if (entry.domain !== state.domainFilter && !changed.includes(state.domainFilter)) return false;
    }
    if (!state.search) return true;
    const hay = [entry.label || "", entry.action || "", entry.domain || "", entry.pageUrl || ""].join(" ").toLowerCase();
    return hay.includes(state.search);
  });
  rows = rows.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));
  if (!rows.length) {
    refs.dashboardContent.innerHTML = "";
    refs.emptyState.hidden = false;
    return;
  }
  refs.emptyState.hidden = true;
  const html = ['<div class="hmdb-timeline-list">'];
  for (const row of rows) {
    const count = Array.isArray(row.changes) ? row.changes.length : 0;
    html.push('<article class="hmdb-timeline-item"><div class="hmdb-timeline-title">' + esc(row.label || "Change") + '</div><div class="hmdb-timeline-meta"><span>' + esc(formatDateTime(row.timestamp)) + '</span><span>' + count + ' items</span>' + (row.domain ? ('<span>' + esc(row.domain) + '</span>') : '') + '<button class="hmdb-timeline-undo" data-action="undo-timeline" data-entry-id="' + esc(row.id || "") + '">Undo</button></div></article>');
  }
  html.push('</div>');
  refs.dashboardContent.innerHTML = html.join("");
}

function render() {
  buildDomainChips();
  buildCategoryUi();
  const cbtns = refs.colorFilter.querySelectorAll("button[data-color]");
  cbtns.forEach((btn) => btn.classList.toggle("active", (btn.getAttribute("data-color") || "all") === state.colorFilter));
  setStatusUi();
  if (state.view === "timeline") {
    renderTimeline();
  } else if (state.view === "notebook") {
    renderNotebookOnly();
  } else {
    renderHighlights();
  }

  if (state.view === "timeline") {
    const visible = refs.dashboardContent.querySelectorAll(".hmdb-timeline-item").length;
    refs.summaryText.textContent = visible + " timeline entries";
    refs.selectionInfo.textContent = "Timeline view";
  } else if (state.view === "notebook") {
    const visible = refs.dashboardContent.querySelectorAll(".hmdb-file-card").length;
    refs.summaryText.textContent = visible + " notebook item" + (visible === 1 ? "" : "s");
    refs.selectionInfo.textContent = "Notebook mode";
  } else {
    const visibleHighlights = refs.dashboardContent.querySelectorAll(".hmdb-card").length;
    const visibleNotebook = refs.dashboardContent.querySelectorAll(".hmdb-file-card").length;
    const visibleTotal = visibleHighlights + visibleNotebook;
    refs.summaryText.textContent = visibleTotal + " results \u00b7 " + flatHighlights().length + " highlights \u2022 " + state.notebookItems.length + " notebook";
    refs.selectionInfo.textContent = state.favorites.length + " pinned \u2022 " + countHighlightAttachmentItems() + " attachments";
  }
}

function hasHighlightOwner(ownerKey) {
  const parsed = parseOwnerKey(ownerKey);
  if (!parsed) return false;
  const list = state.highlights[parsed.domain] || [];
  return list.some((item) => normalizeId(item.id) === parsed.id);
}

async function pruneOrphanHighlightAttachments() {
  let changed = false;
  const next = {};
  for (const [ownerKey, items] of Object.entries(state.highlightAttachments || {})) {
    if (!hasHighlightOwner(ownerKey)) {
      changed = true;
      for (const item of items || []) {
        const id = normalizeId(item && item.id);
        if (!id) continue;
        const existingUrl = state.notebookObjectUrls.get(id);
        if (existingUrl) {
          try { URL.revokeObjectURL(existingUrl); } catch (_err) {}
          state.notebookObjectUrls.delete(id);
        }
        await idbDeleteFile(id).catch(() => {});
      }
      continue;
    }
    const clean = sanitizeHighlightAttachmentItems(items, ownerKey);
    if (clean.length > 0) {
      next[ownerKey] = clean;
      if (clean.length !== (Array.isArray(items) ? items.length : 0)) changed = true;
    } else {
      changed = true;
    }
  }
  if (changed) {
    state.highlightAttachments = next;
    await saveHighlightAttachments();
  } else {
    state.highlightAttachments = sanitizeHighlightAttachmentMap(state.highlightAttachments);
  }
}

async function loadDataAndRender() {
  clearNotebookObjectUrls();
  await Promise.all([
    loadNotebookItems(),
    loadHighlightAttachments()
  ]);
  const [enabledRes, highlightsRes, favoritesRes, categoriesRes, timelineRes] = await Promise.all([
    sendMsg({ action: "getEnabled" }),
    sendMsg({ action: "getAllHighlights" }),
    sendMsg({ action: "getFavorites" }),
    sendMsg({ action: "getCategories" }),
    sendMsg({ action: "getTimelineForPage", pageUrl: "", domain: "", limit: 200 })
  ]);
  state.enabled = enabledRes && enabledRes.enabled !== false;
  state.highlights = normalizeHighlightsMap(highlightsRes.highlights || {});
  state.favorites = Array.isArray(favoritesRes.favorites) ? Array.from(new Set(favoritesRes.favorites.map(normalizeId).filter(Boolean))) : [];
  state.categories = categoriesRes.categories || {};
  state.domainAssignments = categoriesRes.domainAssignments || {};
  state.domainAssignmentCutoffs = categoriesRes.domainAssignmentCutoffs || {};
  state.timeline = Array.isArray(timelineRes.entries) ? timelineRes.entries : [];
  await pruneOrphanHighlightAttachments();
  render();
}
function lockButton(button, fn) {
  if (!button || button.dataset.busy === "1") return;
  button.dataset.busy = "1";
  Promise.resolve().then(fn).finally(() => { delete button.dataset.busy; });
}

function getHighlight(domain, id) {
  const list = state.highlights[domain] || [];
  return list.find((item) => normalizeId(item.id) === normalizeId(id)) || null;
}

function getNotebookItem(id) {
  const target = normalizeId(id);
  return state.notebookItems.find((item) => normalizeId(item.id) === target) || null;
}

function getHighlightOwnerKey(item) {
  if (!item) return "";
  const domain = normalizeId(item.domain);
  const id = normalizeId(item.id);
  if (!domain || !id) return "";
  return ownerKeyFor(domain, id);
}

function getHighlightAttachmentsForItem(item) {
  const key = getHighlightOwnerKey(item);
  if (!key) return [];
  return Array.isArray(state.highlightAttachments[key]) ? state.highlightAttachments[key] : [];
}

function countHighlightAttachmentItems() {
  let total = 0;
  for (const items of Object.values(state.highlightAttachments || {})) {
    if (!Array.isArray(items)) continue;
    total += items.length;
  }
  return total;
}

function findHighlightAttachment(attachmentId, ownerKeyHint) {
  const wantedId = normalizeId(attachmentId);
  if (!wantedId) return null;

  const tryOwner = parseOwnerKey(ownerKeyHint);
  if (tryOwner) {
    const items = state.highlightAttachments[tryOwner.ownerKey] || [];
    const index = items.findIndex((item) => normalizeId(item.id) === wantedId);
    if (index >= 0) {
      return { ownerKey: tryOwner.ownerKey, index, item: items[index] };
    }
  }

  for (const [ownerKey, items] of Object.entries(state.highlightAttachments || {})) {
    const index = Array.isArray(items)
      ? items.findIndex((item) => normalizeId(item.id) === wantedId)
      : -1;
    if (index >= 0) return { ownerKey, index, item: items[index] };
  }
  return null;
}

function copyText(text) {
  const value = String(text || "");
  if (!value) return;
  navigator.clipboard.writeText(value).then(() => toast("Copied")).catch(() => {
    const ta = document.createElement("textarea");
    ta.value = value;
    ta.style.position = "fixed";
    ta.style.opacity = "0";
    document.body.appendChild(ta);
    ta.select();
    document.execCommand("copy");
    document.body.removeChild(ta);
    toast("Copied");
  });
}

async function onCardAction(action, domain, id) {
  const item = getHighlight(domain, id);
  if (!item) {
    toast("Item not found");
    await loadDataAndRender();
    return;
  }
  if (action === "copy") { copyText(getItemLabel(item)); return; }
  if (action === "attach") {
    pendingAttachmentTarget = { domain: item.domain || domain, id: item.id };
    if (refs.highlightFileInput) refs.highlightFileInput.click();
    return;
  }
  if (action === "pin") {
    await sendMsg({ action: "toggleFavorite", id: item.id });
    if (state.favorites.includes(item.id)) state.favorites = state.favorites.filter((fav) => fav !== item.id);
    else state.favorites = state.favorites.concat(item.id);
    render();
    return;
  }
  if (action === "delete") {
    const res = await sendMsg({ action: "deleteHighlight", domain, id });
    if (!res || !res.ok) { toast("Could not delete"); return; }
    toast("Deleted");
    await loadDataAndRender();
    return;
  }
  if (action === "note") {
    openNoteModal(item, { type: "highlight", id: item.id, domain: item.domain });
  }
}

function openNoteModal(item, target) {
  if (!item || !target) return;
  activeNoteTarget = target;
  if (target.type === "notebook") {
    refs.noteContext.textContent = "Notebook \u2022 " + shortText(item.name || "Attachment", 120);
  } else {
    refs.noteContext.textContent = item.domain + " \u00b7 " + shortText(getItemLabel(item), 120);
  }
  refs.noteEditor.value = String(item.note || "");
  lastSavedNoteValue = refs.noteEditor.value;
  refs.noteSaveState.textContent = "Saved";
  refs.noteModal.hidden = false;
  refs.noteEditor.focus();
}

async function saveOpenNote() {
  if (!activeNoteTarget) return;
  const next = refs.noteEditor.value;
  if (next === lastSavedNoteValue) { refs.noteSaveState.textContent = "Saved"; return; }
  refs.noteSaveState.textContent = "Saving...";

  if (activeNoteTarget.type === "notebook") {
    const notebookItem = getNotebookItem(activeNoteTarget.id);
    if (!notebookItem) {
      refs.noteSaveState.textContent = "Item missing";
      return;
    }
    notebookItem.note = next;
    notebookItem.updatedAt = Date.now();
    await saveNotebookItems();
    lastSavedNoteValue = next;
    refs.noteSaveState.textContent = "Saved";
    render();
    return;
  }

  const item = getHighlight(activeNoteTarget.domain, activeNoteTarget.id);
  if (!item) {
    refs.noteSaveState.textContent = "Item missing";
    return;
  }

  const res = await sendMsg({
    action: "saveHighlight",
    highlight: { ...item, domain: item.domain || activeNoteTarget.domain, note: next },
    historyMeta: { action: "update", label: next.trim() ? "Updated highlight note" : "Cleared highlight note" }
  });
  if (!res || !res.ok) {
    refs.noteSaveState.textContent = "Save failed";
    return;
  }
  item.note = next;
  lastSavedNoteValue = next;
  refs.noteSaveState.textContent = "Saved";
  render();
}

function scheduleNoteSave() {
  clearTimeout(noteSaveTimer);
  refs.noteSaveState.textContent = "Typing...";
  noteSaveTimer = setTimeout(() => { saveOpenNote(); }, 320);
}

async function closeNoteModal() {
  clearTimeout(noteSaveTimer);
  await saveOpenNote();
  activeNoteTarget = null;
  refs.noteModal.hidden = true;
  if (pendingRealtimeRefresh) {
    scheduleRealtimeRefresh();
  }
}

function scheduleRealtimeRefresh() {
  pendingRealtimeRefresh = true;
  if (!refs.noteModal.hidden) return;
  if (realtimeRefreshInFlight) return;
  clearTimeout(realtimeRefreshTimer);
  realtimeRefreshTimer = setTimeout(async () => {
    if (realtimeRefreshInFlight) return;
    realtimeRefreshInFlight = true;
    pendingRealtimeRefresh = false;
    try {
      await loadDataAndRender();
    } catch (_err) {
      // Ignore transient refresh failures.
    } finally {
      realtimeRefreshInFlight = false;
      realtimeRefreshTimer = null;
      if (pendingRealtimeRefresh) {
        scheduleRealtimeRefresh();
      }
    }
  }, 160);
}

function downloadJson(payload, name) {
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  a.click();
  URL.revokeObjectURL(url);
}

async function togglePower() {
  const next = !state.enabled;
  const res = await sendMsg({ action: "setEnabled", enabled: next });
  if (res && res.ok !== false) {
    state.enabled = next;
    setStatusUi();
    toast(next ? "System enabled" : "System paused");
  } else {
    toast("Power change failed");
  }
}

async function toggleFullscreen() {
  if (!chrome.windows || !chrome.windows.getCurrent || !chrome.windows.update) {
    toast("Fullscreen unavailable");
    return;
  }
  const win = await new Promise((resolve) => chrome.windows.getCurrent({}, (w) => resolve(chrome.runtime.lastError ? null : w || null)));
  if (!win || typeof win.id !== "number") { toast("Window unavailable"); return; }
  const next = win.state === "fullscreen" ? "normal" : "fullscreen";
  await new Promise((resolve) => chrome.windows.update(win.id, { state: next }, () => resolve()));
}

async function exportAll() {
  const res = await sendMsg({ action: "exportAll" });
  if (!res || !res.ok || !res.payload) { toast("Export failed"); return; }
  downloadJson(res.payload, "highlightmaster-dashboard-export.json");
  toast("Exported");
}

async function importAll(file) {
  if (!file) return;
  let payload;
  try { payload = JSON.parse(await file.text()); } catch (_err) { toast("Invalid JSON file"); return; }
  const res = await sendMsg({ action: "importAll", payload });
  if (!res || !res.ok) { toast("Import failed"); return; }
  toast("Imported " + (res.imported || 0) + " items");
  await loadDataAndRender();
}

async function restoreLatestDeleted() {
  const deleted = await sendMsg({ action: "getDeletedAuditLog", limit: 1 });
  const rows = deleted && Array.isArray(deleted.rows) ? deleted.rows : [];
  if (!rows.length || !rows[0].rowKey) { toast("No deleted items"); return; }
  const res = await sendMsg({ action: "restoreDeletedAuditRows", rowIds: [rows[0].rowKey] });
  if (!res || !res.ok) { toast("Restore failed"); return; }
  toast("Restored latest item");
  await loadDataAndRender();
}

async function undoTimeline(entryId) {
  if (!entryId) return;
  const res = await sendMsg({ action: "undoTimelineEntry", entryId });
  if (!res || !res.ok) { toast("Undo failed"); return; }
  toast("Undo complete");
  await loadDataAndRender();
}

async function uploadNotebookFiles(files) {
  const list = Array.from(files || []).filter((file) => file && file.size >= 0);
  if (!list.length) return;

  let uploaded = 0;
  for (const file of list) {
    const id = makeNotebookId("nb_file");
    try {
      await idbPutFile(id, file);
      state.notebookItems.push({
        id,
        kind: "file",
        name: file.name || "Attachment",
        mimeType: file.type || "",
        size: Number(file.size) || 0,
        createdAt: Date.now(),
        updatedAt: Date.now(),
        note: "",
        emoji: ""
      });
      uploaded += 1;
    } catch (_err) {}
  }

  if (!uploaded) {
    toast("Could not attach files");
    return;
  }

  state.notebookItems.sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0));
  await saveNotebookItems();
  render();
  toast(uploaded + " file" + (uploaded === 1 ? "" : "s") + " attached");
}

async function uploadHighlightAttachments(target, files) {
  if (!target || !target.domain || !target.id) {
    if (Array.isArray(files) && files.length) toast("Select a highlight first");
    return;
  }
  const list = Array.from(files || []).filter((file) => file && file.size >= 0);
  if (!list.length) return;
  const ownerKey = ownerKeyFor(target.domain, target.id);
  const current = Array.isArray(state.highlightAttachments[ownerKey])
    ? state.highlightAttachments[ownerKey].slice()
    : [];

  let uploaded = 0;
  for (const file of list) {
    const id = makeNotebookId("ha_file");
    try {
      await idbPutFile(id, file);
      current.push({
        id,
        ownerKey,
        ownerDomain: target.domain,
        ownerId: target.id,
        name: file.name || "Attachment",
        mimeType: file.type || "",
        size: Number(file.size) || 0,
        createdAt: Date.now(),
        updatedAt: Date.now()
      });
      uploaded += 1;
    } catch (_err) {}
  }

  if (!uploaded) {
    toast("Could not attach files");
    return;
  }

  current.sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0));
  state.highlightAttachments[ownerKey] = current;
  await saveHighlightAttachments();
  render();
  toast(uploaded + " attachment" + (uploaded === 1 ? "" : "s") + " added");
}

async function openHighlightAttachment(item, asDownload) {
  const blob = await idbGetFile(item.id);
  if (!blob) {
    toast("File missing");
    return;
  }
  const blobUrl = URL.createObjectURL(blob);
  if (asDownload) {
    const anchor = document.createElement("a");
    anchor.href = blobUrl;
    anchor.download = item.name || "attachment";
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(blobUrl), 2000);
    return;
  }
  window.open(blobUrl, "_blank", "noopener,noreferrer");
  setTimeout(() => URL.revokeObjectURL(blobUrl), 30000);
}

async function deleteHighlightAttachment(attachmentId, ownerKeyHint) {
  const found = findHighlightAttachment(attachmentId, ownerKeyHint);
  if (!found || !found.item) {
    toast("Attachment not found");
    return;
  }
  await idbDeleteFile(found.item.id).catch(() => {});
  const existingUrl = state.notebookObjectUrls.get(found.item.id);
  if (existingUrl) {
    try { URL.revokeObjectURL(existingUrl); } catch (_err) {}
    state.notebookObjectUrls.delete(found.item.id);
  }
  const list = Array.isArray(state.highlightAttachments[found.ownerKey])
    ? state.highlightAttachments[found.ownerKey].slice()
    : [];
  const next = list.filter((row) => normalizeId(row.id) !== normalizeId(found.item.id));
  if (next.length > 0) state.highlightAttachments[found.ownerKey] = next;
  else delete state.highlightAttachments[found.ownerKey];
  await saveHighlightAttachments();
  render();
  toast("Attachment removed");
}

async function handleHighlightAttachmentAction(action, attachmentId, ownerKey) {
  const found = findHighlightAttachment(attachmentId, ownerKey);
  if (!found || !found.item) {
    toast("Attachment not found");
    return;
  }
  if (action === "open") {
    await openHighlightAttachment(found.item, false);
    return;
  }
  if (action === "download") {
    await openHighlightAttachment(found.item, true);
    return;
  }
  if (action === "delete") {
    await deleteHighlightAttachment(attachmentId, ownerKey);
  }
}

async function openNotebookFile(item, asDownload) {
  const blob = await idbGetFile(item.id);
  if (!blob) {
    toast("File missing");
    return;
  }
  const blobUrl = URL.createObjectURL(blob);
  if (asDownload) {
    const anchor = document.createElement("a");
    anchor.href = blobUrl;
    anchor.download = item.name || "attachment";
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(blobUrl), 2000);
    return;
  }
  window.open(blobUrl, "_blank", "noopener,noreferrer");
  setTimeout(() => URL.revokeObjectURL(blobUrl), 30000);
}

async function deleteNotebookItem(itemId) {
  const item = getNotebookItem(itemId);
  if (!item) {
    toast("Item not found");
    return;
  }
  if (item.kind === "file") {
    await idbDeleteFile(item.id).catch(() => {});
  }
  state.notebookItems = state.notebookItems.filter((row) => row.id !== item.id);
  const url = state.notebookObjectUrls.get(item.id);
  if (url) {
    try { URL.revokeObjectURL(url); } catch (_err) {}
    state.notebookObjectUrls.delete(item.id);
  }
  await saveNotebookItems();
  render();
  toast("Notebook item deleted");
}

async function handleNotebookAction(action, itemId) {
  const item = getNotebookItem(itemId);
  if (!item) {
    toast("Item not found");
    return;
  }

  if (action === "open" && item.kind === "file") {
    await openNotebookFile(item, false);
    return;
  }
  if (action === "download" && item.kind === "file") {
    await openNotebookFile(item, true);
    return;
  }
  if (action === "edit-note") {
    openNoteModal(item, { type: "notebook", id: item.id });
    return;
  }
  if (action === "delete") {
    await deleteNotebookItem(item.id);
    return;
  }
}

function createNotebookNote() {
  const title = window.prompt("Notebook note title");
  if (title == null) return;
  const cleanTitle = String(title || "").trim() || "Notebook Note";
  const emoji = window.prompt("Emoji (optional)", "\ud83d\udca1");
  const id = makeNotebookId("nb_note");
  state.notebookItems.unshift({
    id,
    kind: "text-note",
    name: cleanTitle,
    mimeType: "text/plain",
    size: 0,
    createdAt: Date.now(),
    updatedAt: Date.now(),
    note: "",
    emoji: emoji == null ? "" : String(emoji)
  });
  saveNotebookItems().then(() => {
    render();
    const added = getNotebookItem(id);
    if (added) {
      openNoteModal(added, { type: "notebook", id });
    }
  });
}

function createCategory() {
  const name = window.prompt("Category name");
  if (name == null) return;
  const clean = String(name).trim();
  if (!clean) { toast("Category name is required"); return; }
  const id = "cat_" + Date.now() + "_" + Math.random().toString(36).slice(2, 7);
  sendMsg({ action: "createCategory", id, name: clean, parentId: null }).then((res) => {
    if (!res || res.ok === false) { toast("Category create failed"); return; }
    toast("Category created");
    loadDataAndRender();
  });
}
function attachEvents() {
  refs.btnUploadNotebookFile.addEventListener("click", () => refs.notebookFileInput.click());
  refs.btnDashboardPower.addEventListener("click", () => lockButton(refs.btnDashboardPower, togglePower));
  refs.btnToggleFullscreen.addEventListener("click", () => lockButton(refs.btnToggleFullscreen, toggleFullscreen));
  refs.btnRefreshDashboard.addEventListener("click", () => lockButton(refs.btnRefreshDashboard, async () => { await loadDataAndRender(); toast("Dashboard refreshed"); }));

  refs.searchInput.addEventListener("input", () => {
    const next = String(refs.searchInput.value || "").trim().toLowerCase();
    clearTimeout(searchDebounce);
    searchDebounce = setTimeout(() => { state.search = next; render(); }, 90);
  });

  refs.domainChips.addEventListener("click", (event) => {
    const btn = event.target.closest("button[data-domain]");
    if (!btn) return;
    state.domainFilter = btn.getAttribute("data-domain") || "all";
    render();
  });

  refs.colorFilter.addEventListener("click", (event) => {
    const btn = event.target.closest("button[data-color]");
    if (!btn) return;
    state.colorFilter = btn.getAttribute("data-color") || "all";
    render();
  });

  refs.sortSelect.addEventListener("change", () => {
    state.sortOrder = refs.sortSelect.value || "date_desc";
    render();
  });

  refs.categoryFilter.addEventListener("change", () => {
    state.categoryFilter = refs.categoryFilter.value || "all";
    render();
  });

  refs.categoryList.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-cat-filter]");
    if (!button) return;
    state.categoryFilter = button.getAttribute("data-cat-filter") || "all";
    refs.categoryFilter.value = state.categoryFilter;
    render();
  });

  for (const nav of refs.navButtons) {
    nav.addEventListener("click", () => {
      state.view = nav.getAttribute("data-view") || "all";
      refs.navButtons.forEach((b) => b.classList.toggle("active", b === nav));
      render();
    });
  }

  refs.dashboardContent.addEventListener("click", (event) => {
    const undoBtn = event.target.closest("[data-action='undo-timeline']");
    if (undoBtn) {
      lockButton(undoBtn, async () => { await undoTimeline(undoBtn.getAttribute("data-entry-id") || ""); });
      return;
    }
    const notebookBtn = event.target.closest("[data-nb-action][data-nb-id]");
    if (notebookBtn) {
      lockButton(notebookBtn, async () => {
        await handleNotebookAction(
          notebookBtn.getAttribute("data-nb-action") || "",
          notebookBtn.getAttribute("data-nb-id") || ""
        );
      });
      return;
    }
    const attachmentBtn = event.target.closest("[data-ha-action][data-ha-id]");
    if (attachmentBtn) {
      lockButton(attachmentBtn, async () => {
        await handleHighlightAttachmentAction(
          attachmentBtn.getAttribute("data-ha-action") || "",
          attachmentBtn.getAttribute("data-ha-id") || "",
          attachmentBtn.getAttribute("data-ha-owner") || ""
        );
      });
      return;
    }
    const button = event.target.closest("button[data-action][data-domain][data-id]");
    if (!button) return;
    lockButton(button, async () => {
      await onCardAction(button.getAttribute("data-action") || "", button.getAttribute("data-domain") || "", button.getAttribute("data-id") || "");
    });
  });

  refs.btnNewCategory.addEventListener("click", createCategory);
  refs.btnNewNotebookNote.addEventListener("click", () => {
    state.view = "notebook";
    refs.navButtons.forEach((b) => b.classList.toggle("active", (b.getAttribute("data-view") || "") === "notebook"));
    createNotebookNote();
  });
  refs.btnExport.addEventListener("click", () => lockButton(refs.btnExport, exportAll));
  refs.btnImport.addEventListener("click", () => refs.importFile.click());
  refs.importFile.addEventListener("change", async () => {
    const file = refs.importFile.files && refs.importFile.files[0] ? refs.importFile.files[0] : null;
    await importAll(file);
    refs.importFile.value = "";
  });
  refs.notebookFileInput.addEventListener("change", async () => {
    const files = refs.notebookFileInput.files ? Array.from(refs.notebookFileInput.files) : [];
    await uploadNotebookFiles(files);
    refs.notebookFileInput.value = "";
  });
  refs.highlightFileInput.addEventListener("change", async () => {
    const target = pendingAttachmentTarget;
    pendingAttachmentTarget = null;
    const files = refs.highlightFileInput.files ? Array.from(refs.highlightFileInput.files) : [];
    await uploadHighlightAttachments(target, files);
    refs.highlightFileInput.value = "";
  });
  refs.btnRestore.addEventListener("click", () => lockButton(refs.btnRestore, restoreLatestDeleted));

  refs.noteEditor.addEventListener("input", scheduleNoteSave);
  refs.btnCloseNoteModal.addEventListener("click", () => lockButton(refs.btnCloseNoteModal, closeNoteModal));
  refs.noteModal.addEventListener("click", (event) => {
    if (event.target === refs.noteModal) lockButton(refs.noteModal, closeNoteModal);
  });
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && !refs.noteModal.hidden) {
      event.preventDefault();
      lockButton(refs.noteModal, closeNoteModal);
    }
  });

  window.addEventListener("unload", () => {
    clearTimeout(realtimeRefreshTimer);
    clearNotebookObjectUrls();
  });

  chrome.storage.onChanged.addListener((changes, areaName) => {
    if (!changes || !areaName) return;
    const syncKeys = ["enabled", "highlights", "favorites", "categories", "domainAssignments", "domainAssignmentCutoffs", "timelineHistory"];
    if (areaName === "sync") {
      if (syncKeys.some((key) => Object.prototype.hasOwnProperty.call(changes, key))) {
        scheduleRealtimeRefresh();
      }
      return;
    }
    if (areaName === "local") {
      if (Object.prototype.hasOwnProperty.call(changes, NOTEBOOK_META_KEY) || Object.prototype.hasOwnProperty.call(changes, HIGHLIGHT_ATTACHMENT_META_KEY)) {
        scheduleRealtimeRefresh();
      }
    }
  });

  window.addEventListener("focus", () => {
    scheduleRealtimeRefresh();
  });
  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) scheduleRealtimeRefresh();
  });
}

async function init() {
  setVersionUi();
  attachEvents();
  refs.sortSelect.value = state.sortOrder;
  state.view = "all";
  refs.navButtons.forEach((b) => b.classList.toggle("active", (b.getAttribute("data-view") || "") === "all"));
  await loadDataAndRender();
}

init();
