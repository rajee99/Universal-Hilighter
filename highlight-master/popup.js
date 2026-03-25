"use strict";

/*
 * HighlightMaster - Hierarchical Popup Script
 * Features: Categories (H1, H2), Domains, Highlights, Sharing, Modern UI
 */

// ---- State ----

let allHighlights = {};
let favorites = [];
let categories = {};
let domainAssignments = {};
let domainAssignmentCutoffs = {};

let currentTab = "all";
let currentDomain = "";
let currentPageUrl = "";
let searchQuery = "";
let activeColorFilter = "all";
let activeSortOrder = "date_desc";
let activeCategoryFilter = "all";
let expandedDomains = new Set();
let currentTabId = null;
let activeDrawMode = null;
let continuousModeEnabled = false;
let continuousModeColorName = "yellow";
let continuousModeCustomHex = "#ffd42e";
let autoOcrCopyEnabled = false;
let timelineEntries = [];
let selectedHighlightIds = new Set();
let pendingRectanglesDelete = null;
let draggedDomainName = "";

// ---- DOM Refs ----

const toggleEnabled = document.getElementById("toggleEnabled");
const currentDomainText = document.getElementById("currentDomainText");
const searchInput = document.getElementById("searchInput");
const filterColor = document.getElementById("filterColor");
const sortOrder = document.getElementById("sortOrder");
const filterCategory = document.getElementById("filterCategory");
const viewControls = document.getElementById("viewControls");
const domainList = document.getElementById("domainList");
const favoritesList = document.getElementById("favoritesList");
const emptyAll = document.getElementById("emptyAll");
const emptyFav = document.getElementById("emptyFav");
const panelAll = document.getElementById("panelAll");
const panelFavorites = document.getElementById("panelFavorites");
const panelTimeline = document.getElementById("panelTimeline");
const btnExport = document.getElementById("btnExport");
const btnDeletedCsv = document.getElementById("btnDeletedCsv");
const btnRestoreDeleted = document.getElementById("btnRestoreDeleted");
const btnImport = document.getElementById("btnImport");
const importFile = document.getElementById("importFile");
const tabs = document.querySelectorAll(".hm-tab");
const toast = document.getElementById("toast");
const statsCount = document.getElementById("statsCount");
const colorStatsBar = document.getElementById("colorStatsBar");
const popupVersion = document.getElementById("popupVersion");
const btnRectMode = document.getElementById("btnRectMode");
const btnAutoOcrCopy = document.getElementById("btnAutoOcrCopy");
const btnOpenDashboard = document.getElementById("btnOpenDashboard");
const btnContinuousMode = document.getElementById("btnContinuousMode");
const continuousColorSelect = document.getElementById("continuousColor");
const continuousCustomColor = document.getElementById("continuousCustomColor");
const bulkBar = document.getElementById("bulkBar");
const bulkCount = document.getElementById("bulkCount");
const btnSelectVisible = document.getElementById("btnSelectVisible");
const btnClearSelection = document.getElementById("btnClearSelection");
const btnBulkDelete = document.getElementById("btnBulkDelete");
const btnBulkMarkdown = document.getElementById("btnBulkMarkdown");
const bulkColors = document.getElementById("bulkColors");
const timelineList = document.getElementById("timelineList");
const emptyTimeline = document.getElementById("emptyTimeline");
const timelineResults = document.getElementById("timelineResults");
const btnUndoLast = document.getElementById("btnUndoLast");
let toastTimer = null;
let searchDebounceTimer = null;

// Modal Refs
const dialogModal = document.getElementById("dialogModal");
const dialogTitle = document.getElementById("dialogTitle");
const dialogInput = document.getElementById("dialogInput");
const dialogTextarea = document.getElementById("dialogTextarea");
const dialogSelectWrap = document.getElementById("dialogSelectWrap");
const dialogSelect = document.getElementById("dialogSelect");
const dialogBtnCancel = document.getElementById("dialogBtnCancel");
const dialogBtnConfirm = document.getElementById("dialogBtnConfirm");
const btnNewFolder = document.getElementById("btnNewFolder");

let dialogCallback = null;

// ---- Preset Colors ----

const PRESET_COLORS = {
  yellow: { bg: "rgba(255, 212, 46, 0.38)", solid: "#ffd42e" },
  green:  { bg: "rgba(51, 195, 172, 0.34)", solid: "#33c3ac" },
  blue:   { bg: "rgba(79, 159, 232, 0.32)", solid: "#4f9fe8" },
  pink:   { bg: "rgba(255, 126, 182, 0.32)", solid: "#ff7eb6" },
  orange: { bg: "rgba(255, 159, 67, 0.34)", solid: "#ff9f43" }
};

const COLOR_BUCKET_ORDER = ["yellow", "green", "blue", "pink", "orange", "custom"];
const COLOR_BUCKET_LABELS = {
  yellow: "Yellow",
  green: "Green",
  blue: "Blue",
  pink: "Pink",
  orange: "Orange",
  custom: "Custom"
};
const CUSTOM_COLOR_SOLID = "#b38cff";

// ---- Helpers ----

function getColorSolid(colorName) {
  if (!colorName) return "#ffd42e";
  if (colorName.startsWith("custom:")) return colorName.replace("custom:", "");
  return PRESET_COLORS[colorName] ? PRESET_COLORS[colorName].solid : "#ffd42e";
}

function normalizeHexColor(value, fallback = "#ffd42e") {
  const raw = String(value || "").trim();
  if (/^#[0-9a-fA-F]{6}$/.test(raw)) return raw.toLowerCase();
  return fallback;
}

function normalizeContinuousColorName(colorName, bgColor) {
  const raw = String(colorName || "").trim();
  if (raw.startsWith("custom:")) {
    const fromName = normalizeHexColor(raw.slice(7), normalizeHexColor(bgColor, continuousModeCustomHex));
    return "custom:" + fromName;
  }
  if (PRESET_COLORS[raw]) return raw;
  return "yellow";
}

function resolveContinuousColorPayloadFromControls() {
  const selected = continuousColorSelect ? (continuousColorSelect.value || "yellow") : "yellow";
  if (selected === "custom") {
    const hex = normalizeHexColor(
      continuousCustomColor ? continuousCustomColor.value : continuousModeCustomHex,
      continuousModeCustomHex
    );
    continuousModeCustomHex = hex;
    return { colorName: "custom:" + hex, bgColor: hex };
  }
  const colorName = PRESET_COLORS[selected] ? selected : "yellow";
  return { colorName, bgColor: PRESET_COLORS[colorName].bg };
}

function syncContinuousColorControls(colorName, bgColor) {
  const normalized = normalizeContinuousColorName(colorName, bgColor);
  if (normalized.startsWith("custom:")) {
    const customHex = normalizeHexColor(normalized.slice(7), normalizeHexColor(bgColor, continuousModeCustomHex));
    continuousModeCustomHex = customHex;
    if (continuousColorSelect) continuousColorSelect.value = "custom";
    if (continuousCustomColor) continuousCustomColor.value = customHex;
  } else {
    if (continuousColorSelect) continuousColorSelect.value = normalized;
  }
  continuousModeColorName = normalized;
}

function updateContinuousModeUi() {
  const canUseContinuous = !!currentTabId && toggleEnabled.checked;
  if (btnContinuousMode) {
    btnContinuousMode.classList.toggle("active", continuousModeEnabled);
    btnContinuousMode.setAttribute("aria-pressed", continuousModeEnabled ? "true" : "false");
    btnContinuousMode.disabled = !canUseContinuous;
    if (!currentTabId) {
      btnContinuousMode.title = "Continuous highlight works on normal web pages only";
    } else if (!toggleEnabled.checked) {
      btnContinuousMode.title = "Enable extension first";
    } else {
      btnContinuousMode.title = continuousModeEnabled
        ? "Disable continuous highlight"
        : "Enable continuous highlight";
    }
  }
  if (continuousColorSelect) {
    continuousColorSelect.disabled = !canUseContinuous;
  }
  if (continuousCustomColor) {
    const showCustom = !!continuousColorSelect && continuousColorSelect.value === "custom";
    continuousCustomColor.classList.toggle("visible", showCustom);
    continuousCustomColor.disabled = !canUseContinuous || !showCustom;
  }
}

function getItemLabel(h) {
  if (!h) return "";
  if (h.type === "shape-cover") return "Cover Rectangle";
  if (h.type === "shape-rect") return h.label || h.ocrText || h.text || "Rectangle Highlight";
  return h.text || h.ocrText || h.label || "Text highlight";
}

function normalizeHighlightId(id) {
  if (id === null || typeof id === "undefined") return "";
  return String(id).trim();
}

function normalizeHighlightsMap(raw) {
  if (!raw || typeof raw !== "object") return {};
  const normalized = {};
  for (const [domain, items] of Object.entries(raw)) {
    if (!Array.isArray(items)) continue;
    const domainKey = typeof domain === "string" ? domain.trim() : "";
    if (!domainKey) continue;
    const nextItems = [];
    const seenIds = new Set();
    for (const item of items) {
      if (!item || typeof item !== "object") continue;
      const itemId = normalizeHighlightId(item.id);
      if (!itemId || seenIds.has(itemId)) continue;
      seenIds.add(itemId);
      nextItems.push({
        ...item,
        id: itemId,
        domain: (typeof item.domain === "string" && item.domain.trim()) ? item.domain.trim() : domainKey
      });
    }
    if (nextItems.length > 0) normalized[domainKey] = nextItems;
  }
  return normalized;
}

function formatDate(timestamp) {
  if (!timestamp) return "";
  const d = new Date(timestamp);
  const now = new Date();
  const diff = now - d;

  if (diff < 60000) return "Just now";
  if (diff < 3600000) return Math.floor(diff / 60000) + "m ago";
  if (diff < 86400000) return Math.floor(diff / 3600000) + "h ago";
  if (diff < 604800000) return Math.floor(diff / 86400000) + "d ago";

  return d.toLocaleDateString(undefined, { month: "short", day: "numeric" });
}

function wordCount(str) {
  if (!str) return 0;
  return str.trim().split(/\s+/).length;
}

function truncateWords(str, maxWords) {
  if (!str) return "";
  const words = str.trim().split(/\s+/);
  if (words.length <= maxWords) return escapeHtml(str);
  return escapeHtml(words.slice(0, maxWords).join(" ")) + '<span class="hm-dots"> ...</span>';
}

function truncatePlainText(str, maxChars) {
  if (!str) return "";
  const clean = String(str).replace(/\s+/g, " ").trim();
  if (clean.length <= maxChars) return clean;
  return clean.slice(0, Math.max(0, maxChars - 3)).trimEnd() + "...";
}

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
}

function isSavedRectangle(h) {
  return !!(h && h.type === "shape-rect");
}

function buildHighlightSearchHaystack(h, domain) {
  return [
    getItemLabel(h),
    h && h.note ? h.note : "",
    h && h.text ? h.text : "",
    h && h.ocrText ? h.ocrText : "",
    h && h.label ? h.label : "",
    h && h.pageTitle ? h.pageTitle : "",
    h && h.type ? h.type : "",
    domain || ""
  ].join(" ").toLowerCase();
}

function matchesSearch(h, domain) {
  if (!searchQuery) return true;
  return buildHighlightSearchHaystack(h, domain).includes(searchQuery);
}

function getDomainAssignmentInfo(domain) {
  const categoryId = typeof domainAssignments[domain] === "string"
    ? domainAssignments[domain]
    : "";
  const rawCutoff = domainAssignmentCutoffs && Object.prototype.hasOwnProperty.call(domainAssignmentCutoffs, domain)
    ? Number(domainAssignmentCutoffs[domain])
    : 0;
  const cutoff = Number.isFinite(rawCutoff) && rawCutoff > 0 ? rawCutoff : 0;
  return { categoryId, cutoff };
}

function isHighlightAssignedToCategory(h, domain, wantedCategoryId = "") {
  const assignment = getDomainAssignmentInfo(domain);
  if (!assignment.categoryId || !categories[assignment.categoryId]) return false;
  if (!assignment.cutoff) return false;
  const timestamp = Number(h && h.timestamp) || 0;
  if (timestamp > assignment.cutoff) return false;
  if (!wantedCategoryId) return true;
  return assignment.categoryId === wantedCategoryId;
}

function matchesCategoryFilter(h, domain) {
  if (activeCategoryFilter === "all") return true;

  if (activeCategoryFilter === "with_notes") {
    return !!normalizeNoteText(h && h.note);
  }
  if (activeCategoryFilter === "text_only") {
    return !isSavedRectangle(h);
  }
  if (activeCategoryFilter === "rectangles_only") {
    return isSavedRectangle(h);
  }

  if (activeCategoryFilter === "uncategorized") {
    return !isHighlightAssignedToCategory(h, domain);
  }

  if (activeCategoryFilter.startsWith("cat:")) {
    const wantedCategory = activeCategoryFilter.slice(4);
    return !!wantedCategory && isHighlightAssignedToCategory(h, domain, wantedCategory);
  }

  return true;
}

function buildTimelineSearchHaystack(entry) {
  const parts = [
    entry && entry.label ? entry.label : "",
    entry && entry.action ? entry.action : "",
    entry && entry.domain ? entry.domain : "",
    entry && entry.pageUrl ? entry.pageUrl : ""
  ];
  const changes = entry && Array.isArray(entry.changes) ? entry.changes : [];
  for (const change of changes) {
    const before = change && change.before ? change.before : {};
    const after = change && change.after ? change.after : {};
    parts.push(
      before.note || "",
      after.note || "",
      before.text || "",
      after.text || "",
      before.ocrText || "",
      after.ocrText || "",
      before.label || "",
      after.label || "",
      before.pageTitle || "",
      after.pageTitle || ""
    );
  }
  return parts.join(" ").toLowerCase();
}

function syncCategoryFilterOptions() {
  if (!filterCategory) return;
  const previousValue = activeCategoryFilter || filterCategory.value || "all";
  const options = [
    { value: "all", label: "All categories" },
    { value: "uncategorized", label: "Uncategorized" },
    { value: "with_notes", label: "With notes" },
    { value: "text_only", label: "Text highlights only" },
    { value: "rectangles_only", label: "Rectangles only" }
  ];

  const topLevel = Object.entries(categories || {})
    .filter(([, cat]) => cat && !cat.parentId)
    .sort((a, b) => String(a[1].name || "").localeCompare(String(b[1].name || "")));

  for (const [catId, cat] of topLevel) {
    options.push({ value: "cat:" + catId, label: cat.name });
    const children = Object.entries(categories || {})
      .filter(([, child]) => child && child.parentId === catId)
      .sort((a, b) => String(a[1].name || "").localeCompare(String(b[1].name || "")));
    for (const [childId, childCat] of children) {
      options.push({ value: "cat:" + childId, label: "- " + childCat.name });
    }
  }

  filterCategory.innerHTML = "";
  for (const option of options) {
    const el = document.createElement("option");
    el.value = option.value;
    el.textContent = option.label;
    filterCategory.appendChild(el);
  }

  const hasPrevious = options.some((option) => option.value === previousValue);
  filterCategory.value = hasPrevious ? previousValue : "all";
  activeCategoryFilter = filterCategory.value || "all";
}

function getColorBucket(colorName) {
  if (!colorName) return "yellow";
  if (typeof colorName === "string" && colorName.startsWith("custom:")) return "custom";
  return PRESET_COLORS[colorName] ? colorName : "custom";
}

function matchesColorFilter(h) {
  if (activeColorFilter === "all") return true;
  return getColorBucket(h && h.color) === activeColorFilter;
}

function compareBySortOrder(a, b) {
  if (activeSortOrder === "date_asc") {
    return (a.timestamp || 0) - (b.timestamp || 0);
  }
  if (activeSortOrder === "alpha_asc" || activeSortOrder === "alpha_desc") {
    const aLabel = (getItemLabel(a) || "").toLowerCase();
    const bLabel = (getItemLabel(b) || "").toLowerCase();
    const direction = activeSortOrder === "alpha_desc" ? -1 : 1;
    if (aLabel < bLabel) return -1 * direction;
    if (aLabel > bLabel) return 1 * direction;
    return (b.timestamp || 0) - (a.timestamp || 0);
  }
  return (b.timestamp || 0) - (a.timestamp || 0);
}

function sortByPinAndRecent(items) {
  const pinnedSet = new Set(favorites);
  return items.slice().sort((a, b) => {
    const aPinned = pinnedSet.has(a.id) ? 1 : 0;
    const bPinned = pinnedSet.has(b.id) ? 1 : 0;
    if (aPinned !== bPinned) return bPinned - aPinned;
    return compareBySortOrder(a, b);
  });
}

function updateViewControlsVisibility() {
  if (!viewControls) return;
  if (currentTab === "timeline") {
    viewControls.classList.add("hidden");
  } else {
    viewControls.classList.remove("hidden");
  }
}

function showToast(message, options = {}) {
  if (!toast) return;
  const duration = typeof options.duration === "number" ? options.duration : 2200;
  if (toastTimer) {
    clearTimeout(toastTimer);
    toastTimer = null;
  }

  toast.innerHTML = "";
  const msg = document.createElement("span");
  msg.className = "hm-toast-message";
  msg.textContent = message;
  toast.appendChild(msg);

  if (options.actionLabel && typeof options.onAction === "function") {
    const actionBtn = document.createElement("button");
    actionBtn.className = "hm-toast-action";
    actionBtn.textContent = options.actionLabel;
    actionBtn.addEventListener("click", async (e) => {
      e.stopPropagation();
      try {
        await options.onAction();
      } catch (_err) {}
      toast.classList.remove("visible");
      if (toastTimer) {
        clearTimeout(toastTimer);
        toastTimer = null;
      }
    });
    toast.appendChild(actionBtn);
  }

  toast.classList.add("visible");
  toastTimer = setTimeout(() => {
    toast.classList.remove("visible");
    toastTimer = null;
  }, Math.max(500, duration));
}

function showDeleteUndoToast(historyId, message) {
  if (!historyId) {
    showToast(message);
    return;
  }
  showToast(message, {
    duration: 10000,
    actionLabel: "Undo",
    onAction: async () => {
      const res = await sendMsg({ action: "undoTimelineEntry", entryId: historyId });
      if (!res || !res.ok) {
        showToast("Could not undo");
        return;
      }
      selectedHighlightIds.clear();
      await loadData();
      await loadTimeline();
      render();
      showToast("Restored");
    }
  });
}

async function copyToClipboard(text) {
  try {
    await navigator.clipboard.writeText(text);
    showToast("✓ Copied");
  } catch (_e) {
    const textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.style.position = "fixed";
    textarea.style.opacity = "0";
    document.body.appendChild(textarea);
    textarea.select();
    document.execCommand("copy");
    document.body.removeChild(textarea);
    showToast("✓ Copied");
  }
}

function shareText(text, url) {
  const shareContent = text + (url ? "\n\nSource: " + url : "");
  copyToClipboard(shareContent);
  showToast("✓ Copied for sharing");
}

function updateStats() {
  let count = 0;
  const colorCounts = {
    yellow: 0,
    green: 0,
    blue: 0,
    pink: 0,
    orange: 0,
    custom: 0
  };
  for (const domain of Object.keys(allHighlights)) {
    const items = allHighlights[domain] || [];
    count += items.length;
    for (const item of items) {
      const bucket = getColorBucket(item && item.color);
      if (!Object.prototype.hasOwnProperty.call(colorCounts, bucket)) continue;
      colorCounts[bucket] += 1;
    }
  }
  statsCount.textContent = count === 1 ? "1 highlight" : count + " highlights";
  if (!colorStatsBar) return;
  colorStatsBar.innerHTML = "";
  if (count === 0) {
    const empty = document.createElement("span");
    empty.className = "hm-color-segment empty";
    empty.title = "No highlights yet";
    colorStatsBar.appendChild(empty);
    return;
  }
  for (const bucket of COLOR_BUCKET_ORDER) {
    const bucketCount = colorCounts[bucket] || 0;
    if (bucketCount <= 0) continue;
    const segment = document.createElement("span");
    segment.className = "hm-color-segment";
    segment.style.flexGrow = String(bucketCount);
    segment.style.background = bucket === "custom"
      ? CUSTOM_COLOR_SOLID
      : (PRESET_COLORS[bucket] ? PRESET_COLORS[bucket].solid : CUSTOM_COLOR_SOLID);
    const pct = Math.round((bucketCount / count) * 100);
    segment.title = COLOR_BUCKET_LABELS[bucket] + ": " + bucketCount + " (" + pct + "%)";
    segment.setAttribute("aria-label", segment.title);
    colorStatsBar.appendChild(segment);
  }
}

function sendMsg(msg) {
  return new Promise((resolve) => {
    chrome.runtime.sendMessage(msg, (res) => {
      resolve(chrome.runtime.lastError ? { error: chrome.runtime.lastError.message } : (res || {}));
    });
  });
}

function getHighlightById(id) {
  const wantedId = normalizeHighlightId(id);
  if (!wantedId) return null;
  for (const [domain, items] of Object.entries(allHighlights)) {
    if (!Array.isArray(items)) continue;
    const found = items.find((h) => h && normalizeHighlightId(h.id) === wantedId);
    if (found) return { ...found, domain: found.domain || domain };
  }
  return null;
}

function getDomainFromUrl(url) {
  if (!url || typeof url !== "string") return "";
  try {
    return new URL(url).hostname || "";
  } catch (_e) {
    return "";
  }
}

function resolveRecordDomain(record) {
  if (!record || typeof record !== "object") return currentDomain || "";
  if (typeof record.domain === "string" && record.domain.trim()) {
    return record.domain.trim();
  }

  const fromUrl = getDomainFromUrl(record.url || "");
  if (fromUrl) return fromUrl;

  if (record.id) {
    const fromStore = getHighlightById(record.id);
    if (fromStore && typeof fromStore.domain === "string" && fromStore.domain.trim()) {
      return fromStore.domain.trim();
    }
  }

  return currentDomain || "";
}

function buildDeleteTarget(record) {
  const id = normalizeHighlightId(record && record.id);
  if (!id) return null;
  const domain = resolveRecordDomain(record);
  if (!domain) return null;
  return { domain, id };
}

function clearCategoryDropHighlights() {
  for (const header of document.querySelectorAll(".hm-category-header.hm-drop-active")) {
    header.classList.remove("hm-drop-active");
  }
}

function getDraggedDomainFromEvent(event) {
  if (draggedDomainName) return draggedDomainName;
  if (!event || !event.dataTransfer) return "";
  try {
    return String(event.dataTransfer.getData("text/plain") || "").trim();
  } catch (_err) {
    return "";
  }
}

async function assignDomainToCategory(domain, categoryId) {
  const normalizedDomain = String(domain || "").trim();
  if (!normalizedDomain) return;
  const targetCategoryId = categoryId || null;
  const currentCategoryId = getDomainAssignmentInfo(normalizedDomain).categoryId || null;
  if (currentCategoryId === targetCategoryId) return;

  await sendMsg({
    action: "assignDomain",
    domain: normalizedDomain,
    categoryId: targetCategoryId
  });
  await loadData();
  render();

  const categoryLabel = targetCategoryId && categories[targetCategoryId]
    ? categories[targetCategoryId].name
    : "Uncategorized";
  showToast("Moved " + normalizedDomain + " to " + categoryLabel);
}

function getSelectedRecords() {
  const result = [];
  for (const id of selectedHighlightIds) {
    const rec = getHighlightById(id);
    if (rec) {
      result.push({
        ...rec,
        domain: resolveRecordDomain(rec)
      });
    }
  }
  return result;
}

function pruneSelection() {
  const next = new Set();
  for (const id of selectedHighlightIds) {
    if (getHighlightById(id)) next.add(id);
  }
  selectedHighlightIds = next;
}

function updateBulkBar() {
  if (!bulkBar) return;
  const count = selectedHighlightIds.size;
  if (currentTab === "timeline" || count === 0) {
    bulkBar.style.display = "none";
    return;
  }
  bulkBar.style.display = "block";
  bulkCount.textContent = count === 1 ? "1 selected" : count + " selected";
}

function formatDateTime(timestamp) {
  if (!timestamp) return "";
  return new Date(timestamp).toLocaleString(undefined, {
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit"
  });
}

function formatTimelineResultsLabel(count) {
  return count === 1 ? "1 result" : count + " results";
}

function formatDeletedRowOption(row) {
  const kind = row.type === "shape-rect"
    ? "Rectangle"
    : (row.type === "shape-cover" ? "Cover" : "Highlight");
  const label = truncatePlainText(row.label || row.text || row.ocrText || row.pageTitle || row.pageUrl || row.highlightId || kind, 64);
  const deletedAt = row.deletedAt || Date.parse(row.deletedAtIso || "") || 0;
  const when = formatDateTime(deletedAt);
  const status = row.restoredAt ? " | Restored" : "";
  return [when, kind, label].filter(Boolean).join(" | ") + status;
}

function getBulkColorForRecord(record, colorName) {
  const palette = PRESET_COLORS[colorName] || PRESET_COLORS.yellow;
  if (!record) return palette.bg;
  if (record.type === "shape-rect" || record.type === "shape-cover") {
    return palette.solid;
  }
  return palette.bg;
}

async function applyBulkColor(colorName) {
  const selectedRecords = getSelectedRecords();
  if (!selectedRecords.length) return;
  const updates = [];
  for (const record of selectedRecords) {
    if (!record || !record.id) continue;
    const domain = resolveRecordDomain(record);
    if (!domain) continue;
    updates.push({
      domain,
      id: record.id,
      color: colorName,
      bgColor: getBulkColorForRecord(record, colorName)
    });
  }
  if (!updates.length) {
    showToast("Could not update selected items");
    return;
  }
  const res = await sendMsg({
    action: "bulkUpdateHighlights",
    updates,
    historyAction: "bulk_color",
    historyLabel: "Bulk color change (" + updates.length + ")",
    pageUrl: currentPageUrl || "",
    domain: currentDomain || ""
  });
  const updatedCount = res && typeof res.updated === "number" ? res.updated : 0;
  if (res && res.ok && updatedCount > 0) {
    await loadData();
    await loadTimeline();
    render();
    showToast("Updated " + updatedCount + " item(s)");
  } else if (res && res.ok && updatedCount === 0) {
    showToast("No color changes applied");
  } else {
    showToast("Could not update selected items");
  }
}

async function deleteSelectedHighlights() {
  const selectedRecords = getSelectedRecords();
  if (!selectedRecords.length) return;
  const targets = [];
  for (const record of selectedRecords) {
    const target = buildDeleteTarget(record);
    if (target) targets.push(target);
  }
  if (!targets.length) {
    showToast("Could not determine selected items");
    return;
  }
  const res = await sendMsg({
    action: "bulkDeleteHighlights",
    targets,
    historyAction: "bulk_delete",
    historyLabel: "Bulk delete (" + targets.length + ")",
    pageUrl: currentPageUrl || "",
    domain: currentDomain || ""
  });
  let removedCount = res && typeof res.removed === "number" ? res.removed : 0;
  let historyId = res && res.historyId ? res.historyId : null;
  if (removedCount === 0) {
    // Fallback to per-item delete when bulk target matching fails in older data.
    for (const target of targets) {
      const single = await sendMsg({ action: "deleteHighlight", domain: target.domain, id: target.id });
      const singleRemoved = single && typeof single.removed === "number" ? single.removed : 0;
      if (singleRemoved > 0) {
        removedCount += singleRemoved;
        if (!historyId && single.historyId) historyId = single.historyId;
      }
    }
  }
  if (removedCount > 0) {
    selectedHighlightIds.clear();
    await loadData();
    await loadTimeline();
    render();
    showDeleteUndoToast(
      historyId || null,
      "Deleted " + removedCount + " item(s)"
    );
  } else {
    await loadData();
    render();
    showToast("Could not delete selected items");
  }
}

function clearPendingRectanglesDelete(showInfoToast) {
  if (!pendingRectanglesDelete) return;
  if (pendingRectanglesDelete.intervalId) {
    clearInterval(pendingRectanglesDelete.intervalId);
  }
  if (pendingRectanglesDelete.timeoutId) {
    clearTimeout(pendingRectanglesDelete.timeoutId);
  }
  pendingRectanglesDelete = null;
  if (showInfoToast) showToast("Rectangle delete canceled");
}

async function executePendingRectanglesDelete(targets, sourceCount) {
  if (!Array.isArray(targets) || targets.length === 0) return;
  const res = await sendMsg({
    action: "bulkDeleteHighlights",
    targets,
    historyAction: "bulk_delete",
    historyLabel: "Deleted rectangle category (" + (sourceCount || targets.length) + ")",
    pageUrl: currentPageUrl || "",
    domain: currentDomain || ""
  });
  let removedCount = res && typeof res.removed === "number" ? res.removed : 0;
  if (removedCount === 0) {
    for (const target of targets) {
      const single = await sendMsg({ action: "deleteHighlight", domain: target.domain, id: target.id });
      const singleRemoved = single && typeof single.removed === "number" ? single.removed : 0;
      if (singleRemoved > 0) removedCount += singleRemoved;
    }
  }
  const removedIds = new Set(targets.map((t) => t && t.id).filter(Boolean));
  for (const id of removedIds) {
    selectedHighlightIds.delete(id);
  }
  await loadData();
  await loadTimeline();
  render();
  if (removedCount > 0) {
    showToast("Deleted " + removedCount + " rectangles");
  } else {
    showToast("Could not delete rectangles");
  }
}

function scheduleRectanglesCategoryDelete(items) {
  const targets = [];
  for (const item of (items || [])) {
    const target = buildDeleteTarget(item);
    if (!target) continue;
    targets.push(target);
  }
  if (!targets.length) {
    showToast("No rectangles to delete");
    return;
  }

  clearPendingRectanglesDelete(false);

  pendingRectanglesDelete = {
    secondsLeft: 10,
    sourceCount: targets.length,
    targets: targets.slice(),
    intervalId: null,
    timeoutId: null
  };

  pendingRectanglesDelete.intervalId = setInterval(() => {
    if (!pendingRectanglesDelete) return;
    pendingRectanglesDelete.secondsLeft -= 1;
    if (pendingRectanglesDelete.secondsLeft <= 0) return;
    render();
  }, 1000);

  pendingRectanglesDelete.timeoutId = setTimeout(async () => {
    const job = pendingRectanglesDelete;
    clearPendingRectanglesDelete(false);
    if (!job) return;
    await executePendingRectanglesDelete(job.targets, job.sourceCount);
  }, 10000);

  render();
}

function buildMarkdownForHighlights(items) {
  const groups = {};
  for (const item of items) {
    if (!item) continue;
    const url = item.url || "";
    if (!groups[url]) {
      groups[url] = {
        pageTitle: item.pageTitle || item.domain || "Page",
        domain: item.domain || "",
        url,
        items: []
      };
    }
    groups[url].items.push(item);
  }

  const lines = [];
  lines.push("# HighlightMaster Markdown Export");
  lines.push("");
  lines.push("Generated: " + new Date().toISOString());
  lines.push("");

  const orderedGroups = Object.values(groups).sort((a, b) => {
    const ta = a.items[0] ? (a.items[0].timestamp || 0) : 0;
    const tb = b.items[0] ? (b.items[0].timestamp || 0) : 0;
    return tb - ta;
  });

  for (const group of orderedGroups) {
    const title = group.pageTitle || group.domain || "Page";
    if (group.url) lines.push("## [" + title + "](" + group.url + ")");
    else lines.push("## " + title);
    lines.push("");

    group.items.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));
    for (const item of group.items) {
      const typeLabel = item.type === "shape-cover" ? "Cover" : (item.type === "shape-rect" ? "Rectangle" : "Highlight");
      const colorLabel = item.color || "yellow";
      const text = getItemLabel(item);
      const note = normalizeNoteText(item.note || "");
      lines.push("- **" + typeLabel + "** (`" + colorLabel + "`): " + text);
      if (note) lines.push("  - Note: " + note);
    }
    lines.push("");
    if (group.url) lines.push("- Source: [" + group.url + "](" + group.url + ")");
    else if (group.domain) lines.push("- Source: " + group.domain);
    lines.push("");
  }

  return lines.join("\n");
}

async function exportSelectedToMarkdown() {
  const selectedRecords = getSelectedRecords();
  if (!selectedRecords.length) return;
  const markdown = buildMarkdownForHighlights(selectedRecords);
  await copyToClipboard(markdown);
  showToast("Markdown copied (" + selectedRecords.length + ")");
}

async function loadTimeline() {
  const res = await sendMsg({
    action: "getTimelineForPage",
    pageUrl: currentPageUrl || "",
    domain: currentDomain || "",
    limit: 60
  });
  timelineEntries = (res && res.ok && Array.isArray(res.entries)) ? res.entries : [];
}

async function undoTimelineEntry(entryId, silentOnSuccess = false) {
  const res = await sendMsg({ action: "undoTimelineEntry", entryId });
  if (!res || !res.ok) {
    showToast("Nothing to undo");
    return false;
  }
  selectedHighlightIds.clear();
  await loadData();
  await loadTimeline();
  render();
  if (!silentOnSuccess) {
    showToast("Undid: " + ((res.entry && res.entry.label) || "Action"));
  }
  return true;
}

async function undoLastForCurrentPage() {
  try {
    const res = await sendMsg({
      action: "undoLastTimelineForPage",
      pageUrl: currentPageUrl || "",
      domain: currentDomain || ""
    });
    if (!res || !res.ok) {
      const reason = res && res.reason ? res.reason : "unknown";
      if (reason === "not_found") {
        showToast("No actions to undo");
      } else {
        showToast("Failed to undo: " + reason);
      }
      return;
    }
    selectedHighlightIds.clear();
    await loadData();
    await loadTimeline();
    render();
    showToast("Undid: " + ((res.entry && res.entry.label) || "Action"));
  } catch (err) {
    console.error("[HighlightMaster] undoLastForCurrentPage error:", err);
    showToast("Error undoing action");
  }
}

async function openRestoreDeletedDialog() {
  const res = await sendMsg({ action: "getDeletedAuditLog", limit: 80 });
  if (!res || !res.ok || !Array.isArray(res.rows) || res.rows.length === 0) {
    showToast("No deleted items available");
    return;
  }

  const options = res.rows.map((row) => ({
    id: row.rowKey || row.rowId,
    name: formatDeletedRowOption(row)
  })).filter((option) => option.id);

  if (!options.length) {
    showToast("No deleted items available");
    return;
  }

  openDialog({
    title: "Restore Deleted Item",
    hideInput: true,
    showSelect: true,
    selectOptions: options,
    selectPlaceholder: null,
    defaultSelectValue: options[0].id,
    confirmLabel: "Restore",
    onConfirm: async (_, rowKey) => {
      if (!rowKey) {
        showToast("Choose an item to restore");
        return;
      }

      const restoreRes = await sendMsg({ action: "restoreDeletedAuditRows", rowIds: [rowKey] });
      if (!restoreRes || !restoreRes.ok) {
        const reason = restoreRes && restoreRes.reason ? restoreRes.reason : "restore_failed";
        if (reason === "already_exists") {
          showToast("That item is already restored");
        } else {
          showToast("Restore failed");
        }
        return;
      }

      selectedHighlightIds.clear();
      await loadData();
      await loadTimeline();
      render();
      showToast("Restored deleted item", {
        duration: 10000,
        actionLabel: "Undo",
        onAction: async () => {
          if (!restoreRes.historyId) return;
          await undoTimelineEntry(restoreRes.historyId, true);
          showToast("Restore undone");
        }
      });
    }
  });
}

function updateDrawModeButtons() {
  btnRectMode.classList.toggle("active", activeDrawMode === "shape-rect");
}

function updateAutoOcrCopyUi() {
  if (!btnAutoOcrCopy) return;
  btnAutoOcrCopy.classList.toggle("active", autoOcrCopyEnabled);
  btnAutoOcrCopy.setAttribute("aria-pressed", autoOcrCopyEnabled ? "true" : "false");
  btnAutoOcrCopy.title = autoOcrCopyEnabled ? "Auto OCR copy on" : "Auto OCR copy off";
}

function isHttpTab(tab) {
  return !!(tab && tab.url && /^https?:/i.test(tab.url));
}

function getActiveTab() {
  return new Promise((resolve) => {
    chrome.tabs.query({ active: true, currentWindow: true }, (tabsResult) => {
      resolve(Array.isArray(tabsResult) && tabsResult[0] ? tabsResult[0] : null);
    });
  });
}

function sendTabMessage(tabId, msg) {
  return new Promise((resolve) => {
    if (!tabId) {
      resolve({ ok: false, error: "No active tab id" });
      return;
    }
    chrome.tabs.sendMessage(tabId, msg, (res) => {
      if (chrome.runtime.lastError) {
        resolve({ ok: false, error: chrome.runtime.lastError.message });
        return;
      }
      resolve(res || { ok: false });
    });
  });
}

function createTab(url) {
  return new Promise((resolve) => {
    chrome.tabs.create({ url }, (tab) => {
      if (chrome.runtime.lastError) {
        resolve({ ok: false, error: chrome.runtime.lastError.message, tab: null });
        return;
      }
      resolve({ ok: true, tab: tab || null });
    });
  });
}

function queryTabs(queryInfo) {
  return new Promise((resolve) => {
    chrome.tabs.query(queryInfo, (tabsResult) => {
      if (chrome.runtime.lastError) {
        resolve([]);
        return;
      }
      resolve(Array.isArray(tabsResult) ? tabsResult : []);
    });
  });
}

function activateTab(tabId, windowId) {
  return new Promise((resolve) => {
    if (typeof tabId !== "number") {
      resolve({ ok: false, reason: "invalid_tab" });
      return;
    }
    chrome.tabs.update(tabId, { active: true }, () => {
      if (chrome.runtime.lastError) {
        resolve({ ok: false, reason: chrome.runtime.lastError.message });
        return;
      }
      if (typeof windowId === "number" && chrome.windows && chrome.windows.update) {
        chrome.windows.update(windowId, { focused: true }, () => {
          if (chrome.runtime.lastError) {
            resolve({ ok: false, reason: chrome.runtime.lastError.message });
            return;
          }
          resolve({ ok: true });
        });
        return;
      }
      resolve({ ok: true });
    });
  });
}

async function openDashboardInTab() {
  const targetUrl = chrome.runtime.getURL("dashboard.html");
  const existingTabs = await queryTabs({ url: targetUrl });
  if (existingTabs.length > 0) {
    const targetTab = existingTabs[0];
    const focused = await activateTab(targetTab.id, targetTab.windowId);
    if (focused && focused.ok) {
      showToast("Dashboard opened in tab");
      return;
    }
  }
  const created = await createTab(targetUrl);
  if (!created.ok || !created.tab) {
    showToast("Could not open dashboard");
    return;
  }
  showToast("Dashboard opened in tab");
}

async function refreshActiveTabContext() {
  const tab = await getActiveTab();
  if (!isHttpTab(tab)) {
    currentTabId = null;
    currentPageUrl = "";
    currentDomain = "";
    currentDomainText.textContent = "-";
    return;
  }
  currentTabId = tab.id || null;
  currentPageUrl = tab.url || "";
  try {
    currentDomain = new URL(currentPageUrl).hostname;
  } catch (_e) {
    currentDomain = "";
  }
  currentDomainText.textContent = currentDomain || "-";
  if (currentDomain) expandedDomains.add(currentDomain);
}

async function syncDrawModeFromActiveTab() {
  const tab = await getActiveTab();
  if (!isHttpTab(tab)) {
    currentTabId = null;
    activeDrawMode = null;
    updateDrawModeButtons();
    return;
  }
  currentTabId = tab.id;
  const res = await sendTabMessage(currentTabId, { action: "getDrawMode" });
  activeDrawMode = res && res.ok ? (res.drawMode || null) : null;
  updateDrawModeButtons();
}

async function syncContinuousModeFromActiveTab() {
  const tab = await getActiveTab();
  if (!isHttpTab(tab)) {
    currentTabId = null;
    continuousModeEnabled = false;
    syncContinuousColorControls("yellow", PRESET_COLORS.yellow.bg);
    updateContinuousModeUi();
    return;
  }
  currentTabId = tab.id;
  const res = await sendTabMessage(currentTabId, { action: "getContinuousMode" });
  if (res && res.ok) {
    continuousModeEnabled = !!res.enabled;
    syncContinuousColorControls(res.colorName, res.bgColor);
  } else {
    continuousModeEnabled = false;
    syncContinuousColorControls("yellow", PRESET_COLORS.yellow.bg);
  }
  updateContinuousModeUi();
}

async function syncAutoOcrCopySetting() {
  const res = await sendMsg({ action: "getAutoOcrCopy" });
  autoOcrCopyEnabled = !!(res && res.ok && res.enabled);
  updateAutoOcrCopyUi();
}

async function toggleAutoOcrCopySetting() {
  const nextEnabled = !autoOcrCopyEnabled;
  const res = await sendMsg({ action: "setAutoOcrCopy", enabled: nextEnabled });
  if (!res || !res.ok) {
    showToast("Could not update auto OCR copy");
    return;
  }
  autoOcrCopyEnabled = !!res.enabled;
  updateAutoOcrCopyUi();
  showToast(autoOcrCopyEnabled ? "Auto OCR copy on" : "Auto OCR copy off");
}

async function applyContinuousModeToActiveTab(nextEnabled) {
  const tab = await getActiveTab();
  if (!isHttpTab(tab)) {
    currentTabId = null;
    continuousModeEnabled = false;
    updateContinuousModeUi();
    return { ok: false, reason: "unsupported_tab" };
  }
  currentTabId = tab.id;
  const payload = resolveContinuousColorPayloadFromControls();
  const res = await sendTabMessage(currentTabId, {
    action: "setContinuousMode",
    enabled: !!nextEnabled,
    colorName: payload.colorName,
    bgColor: payload.bgColor
  });
  if (res && res.ok) {
    continuousModeEnabled = !!res.enabled;
    syncContinuousColorControls(res.colorName || payload.colorName, res.bgColor || payload.bgColor);
    updateContinuousModeUi();
    return { ok: true, enabled: continuousModeEnabled };
  }
  await syncContinuousModeFromActiveTab();
  return { ok: false, reason: "message_failed" };
}

async function toggleContinuousMode() {
  if (!toggleEnabled.checked) {
    showToast("Enable extension first");
    return;
  }
  const result = await applyContinuousModeToActiveTab(!continuousModeEnabled);
  if (!result.ok) {
    showToast("Could not toggle continuous mode on this page");
    return;
  }
  showToast(result.enabled ? "Continuous highlight on" : "Continuous highlight off");
}

async function updateContinuousColorOnActiveTab() {
  const result = await applyContinuousModeToActiveTab(continuousModeEnabled);
  if (!result.ok) {
    showToast("Could not update continuous color");
  }
}

async function toggleDrawMode(mode) {
  if (!toggleEnabled.checked) {
    showToast("Enable extension first");
    return;
  }

  const tab = await getActiveTab();
  if (!isHttpTab(tab)) {
    showToast("Drawing works on normal web pages only");
    return;
  }
  currentTabId = tab.id;

  const nextMode = activeDrawMode === mode ? null : mode;
  const res = await sendTabMessage(currentTabId, { action: "setDrawMode", mode: nextMode });
  if (!res || !res.ok) {
    activeDrawMode = null;
    updateDrawModeButtons();
    showToast("Could not start draw mode on this page");
    return;
  }

  activeDrawMode = res.drawMode || null;
  updateDrawModeButtons();

  if (activeDrawMode === "shape-rect") {
    showToast("Rectangle mode active");
  } else {
    showToast("Draw mode off");
  }
}

function exportBlob(payload, filename) {
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename.replace(/\s+/g, "_");
  a.click();
  URL.revokeObjectURL(url);
  showToast("📦 Exported successfully");
}
function downloadTextFile(text, filename, mimeType) {
  const blob = new Blob([text], { type: mimeType || "text/plain;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename.replace(/\s+/g, "_");
  a.click();
  URL.revokeObjectURL(url);
}

// ---- Dialog Modal ----

function openDialog({ title, placeholder, showSelect, selectOptions, onConfirm, hideInput, selectPlaceholder, defaultSelectValue, confirmLabel, multiline, initialValue }) {
  dialogTitle.textContent = title;
  dialogBtnConfirm.textContent = confirmLabel || "Save";

  if (multiline) {
    dialogInput.style.display = "none";
    dialogTextarea.style.display = "block";
    dialogTextarea.placeholder = placeholder || "";
    dialogTextarea.value = typeof initialValue === "string" ? initialValue : "";
  } else if (hideInput) {
    dialogInput.style.display = "none";
    dialogTextarea.style.display = "none";
    dialogInput.value = "ignored";
  } else {
    dialogInput.style.display = "block";
    dialogTextarea.style.display = "none";
    dialogInput.placeholder = placeholder || "";
    dialogInput.value = typeof initialValue === "string" ? initialValue : "";
  }
  
  if (showSelect && selectOptions) {
    dialogSelectWrap.style.display = "block";
    dialogSelect.innerHTML = "";
    if (typeof selectPlaceholder === "string" && selectPlaceholder) {
      const placeholderOption = document.createElement("option");
      placeholderOption.value = "";
      placeholderOption.textContent = selectPlaceholder;
      dialogSelect.appendChild(placeholderOption);
    } else if (selectPlaceholder !== null) {
      const placeholderOption = document.createElement("option");
      placeholderOption.value = "";
      placeholderOption.textContent = "None (Top Level Uncategorized)";
      dialogSelect.appendChild(placeholderOption);
    }
    selectOptions.forEach(opt => {
      const el = document.createElement("option");
      el.value = opt.id;
      el.textContent = opt.name;
      dialogSelect.appendChild(el);
    });
    if (typeof defaultSelectValue === "string") {
      dialogSelect.value = defaultSelectValue;
    }
  } else {
    dialogSelectWrap.style.display = "none";
  }
  
  dialogCallback = onConfirm;
  dialogModal.style.display = "flex";
  if (multiline) dialogTextarea.focus();
  else if (!hideInput) dialogInput.focus();
}

function closeDialog() {
  dialogModal.style.display = "none";
  dialogCallback = null;
  dialogBtnConfirm.textContent = "Save";
  if (dialogTextarea) {
    dialogTextarea.value = "";
    dialogTextarea.style.display = "none";
  }
}

dialogBtnCancel.addEventListener("click", closeDialog);
dialogBtnConfirm.addEventListener("click", () => {
  const isMultilineInput = dialogTextarea && dialogTextarea.style.display !== "none";
  const textValue = isMultilineInput
    ? dialogTextarea.value
    : dialogInput.value.trim();
  if (dialogCallback) dialogCallback(textValue, dialogSelect.value);
  closeDialog();
});

dialogInput.addEventListener("keydown", (e) => {
  if (e.key !== "Enter") return;
  e.preventDefault();
  dialogBtnConfirm.click();
});

btnNewFolder.addEventListener("click", () => {
  openDialog({
    title: "Create Category",
    placeholder: "Category Name (e.g., Work)...",
    showSelect: false,
    onConfirm: async (name) => {
      if (!name) return;
      const id = "cat_" + Date.now() + "_" + Math.random().toString(36).slice(2, 6);
      await sendMsg({ action: "createCategory", id, name, parentId: null });
      await loadData();
      render();
    }
  });
});

// ---- Init ----

async function init() {
  const enabledRes = await sendMsg({ action: "getEnabled" });
  toggleEnabled.checked = enabledRes.enabled !== false;
  if (filterColor) filterColor.value = activeColorFilter;
  if (sortOrder) sortOrder.value = activeSortOrder;
  if (filterCategory) filterCategory.value = activeCategoryFilter;
  if (popupVersion && chrome.runtime && chrome.runtime.getManifest) {
    const manifest = chrome.runtime.getManifest();
    popupVersion.textContent = "v" + (manifest && manifest.version ? manifest.version : "-");
  }
  await refreshActiveTabContext();

  await syncDrawModeFromActiveTab();
  await syncContinuousModeFromActiveTab();
  await syncAutoOcrCopySetting();

  await loadData();
  await loadTimeline();
  render();
}

async function loadData() {
  const [hlRes, favRes, catRes] = await Promise.all([
    sendMsg({ action: "getAllHighlights" }),
    sendMsg({ action: "getFavorites" }),
    sendMsg({ action: "getCategories" })
  ]);
  
  allHighlights = normalizeHighlightsMap(hlRes.highlights || {});
  favorites = Array.isArray(favRes.favorites)
    ? Array.from(new Set(favRes.favorites.map((id) => normalizeHighlightId(id)).filter(Boolean)))
    : [];
  categories = catRes.categories || {};
  domainAssignments = catRes.domainAssignments || {};
  domainAssignmentCutoffs = catRes.domainAssignmentCutoffs || {};
  pruneSelection();
}

function collectCategorySubtreeIds(rootId, catsSnapshot) {
  const result = new Set();
  const queue = [rootId];
  while (queue.length) {
    const curr = queue.shift();
    if (!curr || result.has(curr)) continue;
    result.add(curr);
    for (const [catId, cat] of Object.entries(catsSnapshot || {})) {
      if (cat && cat.parentId === curr) queue.push(catId);
    }
  }
  return result;
}

async function deleteCategoryWithUndo(categoryId, categoryName) {
  const catsSnapshot = JSON.parse(JSON.stringify(categories || {}));
  const assignmentsSnapshot = JSON.parse(JSON.stringify(domainAssignments || {}));
  const removedIds = collectCategorySubtreeIds(categoryId, catsSnapshot);
  await sendMsg({ action: "deleteCategory", id: categoryId });
  await loadData();
  render();
  showToast('Deleted category "' + categoryName + '"', {
    duration: 10000,
    actionLabel: "Undo",
    onAction: async () => {
      const pending = new Set(Array.from(removedIds));
      let guard = 0;
      while (pending.size && guard < 80) {
        guard++;
        let progressed = false;
        for (const id of Array.from(pending)) {
          const cat = catsSnapshot[id];
          if (!cat) {
            pending.delete(id);
            continue;
          }
          if (cat.parentId && pending.has(cat.parentId)) continue;
          await sendMsg({ action: "createCategory", id, name: cat.name, parentId: cat.parentId || null });
          pending.delete(id);
          progressed = true;
        }
        if (!progressed) break;
      }
      for (const id of Array.from(pending)) {
        const cat = catsSnapshot[id];
        if (!cat) continue;
        await sendMsg({ action: "createCategory", id, name: cat.name, parentId: null });
      }
      for (const [domain, catId] of Object.entries(assignmentsSnapshot)) {
        if (!removedIds.has(catId)) continue;
        await sendMsg({ action: "assignDomain", domain, categoryId: catId });
      }
      await loadData();
      render();
      showToast("Category restored");
    }
  });
}

// ---- Render Engine ----

function render() {
  syncCategoryFilterOptions();
  updateStats();
  updateViewControlsVisibility();
  if (currentTab === "all") {
    renderAllHighlights();
  } else if (currentTab === "favorites") {
    renderFavorites();
  } else {
    renderTimeline();
  }
  updateBulkBar();
}

function renderAllHighlights() {
  domainList.innerHTML = "";
  const unassignedDomains = [];
  const domainByCat = {};
  const rectangleItems = [];
  
  let totalVisible = 0;
  
  const allDoms = Object.keys(allHighlights).sort();
  for (const dom of allDoms) {
     const allItems = allHighlights[dom] || [];
     const visibleItems = allItems.filter((h) =>
       matchesSearch(h, dom) &&
       matchesColorFilter(h) &&
       matchesCategoryFilter(h, dom)
     );
     const rectsForDomain = visibleItems.filter((h) => isSavedRectangle(h));
     const nonRectItems = visibleItems.filter((h) => !isSavedRectangle(h));

     if (rectsForDomain.length > 0) {
       rectangleItems.push(...rectsForDomain);
     }
     if (nonRectItems.length === 0) continue;

     totalVisible += nonRectItems.length;

     const assignment = getDomainAssignmentInfo(dom);
     const catId = assignment.categoryId;
     let categorizedItems = [];
     if (catId && categories[catId]) {
       categorizedItems = nonRectItems.filter((h) => isHighlightAssignedToCategory(h, dom, catId));
     }
     const unassignedItems = categorizedItems.length > 0
       ? nonRectItems.filter((h) => !isHighlightAssignedToCategory(h, dom, catId))
       : nonRectItems;

     if (categorizedItems.length > 0) {
       if (!domainByCat[catId]) domainByCat[catId] = [];
       domainByCat[catId].push({ domain: dom, items: categorizedItems });
     }
     if (unassignedItems.length > 0) {
       unassignedDomains.push({ domain: dom, items: unassignedItems });
     }
  }

  if (rectangleItems.length > 0) {
    totalVisible += rectangleItems.length;
    const rectangleGroup = renderRectangleCategory(rectangleItems);
    if (rectangleGroup) domainList.appendChild(rectangleGroup);
  }
  
  // 1. Render Categories (H1 -> H2 Hierarchy)
  const rootCats = Object.entries(categories).filter(([id, c]) => !c.parentId);
  for (const [cId, cat] of rootCats) {
      const el = renderCategoryNode(cId, cat, 0, domainByCat);
      if (el) domainList.appendChild(el);
  }
  
  // 2. Render Unassigned Domains
  for (const data of unassignedDomains) {
      domainList.appendChild(renderDomainNode(data.domain, data.items, 0));
  }
  
  emptyAll.style.display = totalVisible === 0 ? "flex" : "none";
}

function renderRectangleCategory(items) {
  const key = "__rectangles__";
  const isExpanded = expandedDomains.has(key);
  const pending = pendingRectanglesDelete;

  const group = document.createElement("div");
  group.className = "hm-domain-group";

  const header = document.createElement("div");
  header.className = "hm-domain-header" + (isExpanded ? " expanded" : "");
  header.innerHTML = `
    <div class="hm-domain-left">
      <svg class="hm-domain-icon" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><rect x="4" y="4" width="16" height="16" rx="2"/></svg>
      <span class="hm-domain-label">Rectangles</span>
      <span class="hm-domain-count">${items.length}</span>
    </div>
    <div class="hm-domain-actions">
      ${pending ? `
        <span class="hm-delete-countdown">Deleting in ${pending.secondsLeft}s</span>
        <button class="hm-delete-undo" data-undo-rect-delete title="Undo pending delete">Undo</button>
      ` : `
        <button class="hm-icon-btn danger" data-clear-rectangles title="Clear rectangles">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/><path d="M10 11v6"/><path d="M14 11v6"/><path d="M9 6V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/></svg>
        </button>
      `}
    </div>
  `;

  const itemsContainer = document.createElement("div");
  itemsContainer.className = "hm-domain-items" + (isExpanded ? "" : " collapsed");

  header.addEventListener("click", (e) => {
    if (e.target.closest("button")) return;
    if (expandedDomains.has(key)) {
      expandedDomains.delete(key);
      itemsContainer.classList.add("collapsed");
      header.classList.remove("expanded");
    } else {
      expandedDomains.add(key);
      itemsContainer.classList.remove("collapsed");
      header.classList.add("expanded");
    }
  });

  const clearRectanglesBtn = header.querySelector("[data-clear-rectangles]");
  if (clearRectanglesBtn) {
    clearRectanglesBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      scheduleRectanglesCategoryDelete(items);
    });
  }

  const undoRectDeleteBtn = header.querySelector("[data-undo-rect-delete]");
  if (undoRectDeleteBtn) {
    undoRectDeleteBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      clearPendingRectanglesDelete(true);
      render();
    });
  }

  const sorted = sortByPinAndRecent(items);
  for (const h of sorted) {
    itemsContainer.appendChild(createHighlightItem(h));
  }

  group.appendChild(header);
  group.appendChild(itemsContainer);
  return group;
}

function renderCategoryNode(cId, cat, level, domainByCat) {
   const subCats = Object.entries(categories).filter(([id, c]) => c.parentId === cId);
   const doms = domainByCat[cId] || [];
   
   // Even if empty, render category so they can delete it or visualize it
   // But hide via search logic if needed. For now, always show folder.
   if (searchQuery && subCats.length === 0 && doms.length === 0) return null;
   
   const group = document.createElement("div");
   group.className = "hm-category-group hm-indent-" + level;
   
   const isExpanded = expandedDomains.has("cat_" + cId);
   
   const header = document.createElement("div");
   header.className = "hm-category-header" + (isExpanded ? " expanded" : "");
   const deleteButton = `
         <button class="hm-icon-btn danger" data-delete-cat="${cId}" title="Delete Category">
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
         </button>
   `;
   header.innerHTML = `
      <div class="hm-domain-left">
        <svg class="hm-domain-icon" width="16" height="16" viewBox="0 0 24 24" fill="currentColor" stroke="none"><path d="M10 4H4c-1.1 0-1.99.9-1.99 2L2 18c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V8c0-1.1-.9-2-2-2h-8l-2-2z"/></svg>
        <span class="hm-domain-label">${escapeHtml(cat.name)}</span>
      </div>
      <div class="hm-domain-actions">
         <button class="hm-icon-btn" data-share-cat="${cId}" title="Share / Export Category">
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
         </button>
         ${deleteButton}
      </div>
   `;
   
   const itemsContainer = document.createElement("div");
   itemsContainer.className = "hm-category-items" + (isExpanded ? "" : " collapsed");
   
   header.addEventListener("click", (e) => {
      if (e.target.closest("button")) return;
      const key = "cat_" + cId;
      if (expandedDomains.has(key)) {
         expandedDomains.delete(key);
         itemsContainer.classList.add("collapsed");
         header.classList.remove("expanded");
      } else {
         expandedDomains.add(key);
         itemsContainer.classList.remove("collapsed");
         header.classList.add("expanded");
      }
   });

   header.addEventListener("dragover", (e) => {
      const draggedDomain = getDraggedDomainFromEvent(e);
      if (!draggedDomain) return;
      e.preventDefault();
      if (e.dataTransfer) e.dataTransfer.dropEffect = "move";
      header.classList.add("hm-drop-active");
   });

   header.addEventListener("dragleave", (e) => {
      const related = e.relatedTarget;
      if (related && header.contains(related)) return;
      header.classList.remove("hm-drop-active");
   });

   header.addEventListener("drop", async (e) => {
      const draggedDomain = getDraggedDomainFromEvent(e);
      if (!draggedDomain) return;
      e.preventDefault();
      e.stopPropagation();
      clearCategoryDropHighlights();
      await assignDomainToCategory(draggedDomain, cId);
   });
   
   const deleteNode = header.querySelector("[data-delete-cat]");
   if (deleteNode) {
      deleteNode.addEventListener("click", async (e) => {
         e.stopPropagation();
         await deleteCategoryWithUndo(cId, cat.name);
      });
   }
   
   header.querySelector("[data-share-cat]").addEventListener("click", async (e) => {
      e.stopPropagation();
      const res = await sendMsg({ action: "exportCategory", categoryId: cId });
      if (!res.ok || !res.payload) return;
      exportBlob(res.payload, "category-export-" + cat.name + ".json");
   });
   
   for (const [sId, sCat] of subCats) {
       const el = renderCategoryNode(sId, sCat, level + 1, domainByCat);
       if (el) itemsContainer.appendChild(el);
   }
   for (const d of doms) {
       itemsContainer.appendChild(renderDomainNode(d.domain, d.items, level + 1));
   }
   
   group.appendChild(header);
   group.appendChild(itemsContainer);
   return group;
}

function renderDomainNode(domain, items, level) {
    const group = document.createElement("div");
    group.className = "hm-domain-group" + (level > 0 ? " hm-indent-" + level : "");

    const isExpanded = expandedDomains.has(domain);

    const header = document.createElement("div");
    header.className = "hm-domain-header" + (isExpanded ? " expanded" : "");
    header.setAttribute("draggable", "true");
    header.innerHTML = `
      <div class="hm-domain-left">
        <svg class="hm-domain-icon" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"><circle cx="12" cy="12" r="10"/><line x1="2" y1="12" x2="22" y2="12"/><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/></svg>
        <span class="hm-domain-label">${escapeHtml(domain)}</span>
        <span class="hm-domain-count">${items.length}</span>
      </div>
      <div class="hm-domain-actions">
        <button class="hm-icon-btn" data-move-domain="${escapeHtml(domain)}" title="Move to Category">
          <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M20 14.66V20a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h5.34"/><polygon points="18 2 22 6 12 16 8 16 8 12 18 2"/></svg>
        </button>
        <button class="hm-icon-btn danger" data-clear-domain="${escapeHtml(domain)}" title="Clear all">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/><path d="M10 11v6"/><path d="M14 11v6"/><path d="M9 6V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/></svg>
        </button>
      </div>
    `;

    const itemsContainer = document.createElement("div");
    itemsContainer.className = "hm-domain-items" + (isExpanded ? "" : " collapsed");

    header.addEventListener("dragstart", (e) => {
      if (e.target && e.target.closest && e.target.closest("button")) {
        e.preventDefault();
        return;
      }
      draggedDomainName = domain;
      group.classList.add("hm-dragging");
      if (e.dataTransfer) {
        e.dataTransfer.effectAllowed = "move";
        try { e.dataTransfer.setData("text/plain", domain); } catch (_err) {}
      }
    });

    header.addEventListener("dragend", () => {
      draggedDomainName = "";
      group.classList.remove("hm-dragging");
      clearCategoryDropHighlights();
    });

    header.addEventListener("click", (e) => {
      if (e.target.closest("button")) return;
      if (expandedDomains.has(domain)) {
        expandedDomains.delete(domain);
        itemsContainer.classList.add("collapsed");
        header.classList.remove("expanded");
      } else {
        expandedDomains.add(domain);
        itemsContainer.classList.remove("collapsed");
        header.classList.add("expanded");
      }
    });

    header.querySelector("[data-move-domain]").addEventListener("click", (e) => {
       e.stopPropagation();
       const allCatOptions = [];
       for (const [id, c] of Object.entries(categories)) {
           if (c.parentId) continue; // H1
           allCatOptions.push({ id, name: c.name });
           // H2
           for (const [sId, sC] of Object.entries(categories)) {
               if (sC.parentId === id) allCatOptions.push({ id: sId, name: "— " + sC.name });
           }
       }
       
        openDialog({
          title: "Move " + domain,
          hideInput: true,
          showSelect: true,
          selectOptions: allCatOptions,
          onConfirm: async (_, catId) => {
              await assignDomainToCategory(domain, catId || null);
          }
        });
     });

    header.querySelector("[data-clear-domain]").addEventListener("click", async (e) => {
      e.stopPropagation();
      const res = await sendMsg({ action: "clearDomain", domain, pageUrl: currentPageUrl || "" });
      expandedDomains.delete(domain);
      await loadData();
      await loadTimeline();
      render();
      showDeleteUndoToast(res.historyId || null, "Cleared highlights for " + domain);
    });

    const orderedItems = sortByPinAndRecent(items);
    for (const h of orderedItems) {
      itemsContainer.appendChild(createHighlightItem(h));
    }

    group.appendChild(header);
    group.appendChild(itemsContainer);
    return group;
}

function renderFavorites() {
  favoritesList.innerHTML = "";
  const favSet = new Set(favorites);
  let favItems = [];

  for (const domain of Object.keys(allHighlights)) {
    for (const h of allHighlights[domain]) {
      if (favSet.has(h.id)) {
        if (!matchesSearch(h, domain)) continue;
        if (!matchesColorFilter(h)) continue;
        if (!matchesCategoryFilter(h, domain)) continue;
        favItems.push(h);
      }
    }
  }

  const ordered = sortByPinAndRecent(favItems);
  for (const item of ordered) {
    favoritesList.appendChild(createHighlightItem(item));
  }

  emptyFav.style.display = ordered.length === 0 ? "flex" : "none";
}

function renderTimeline() {
  timelineList.innerHTML = "";
  const filtered = timelineEntries.filter((entry) => {
    if (!searchQuery) return true;
    return buildTimelineSearchHaystack(entry).includes(searchQuery);
  });

  filtered.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));
  if (timelineResults) {
    timelineResults.textContent = searchQuery
      ? formatTimelineResultsLabel(filtered.length)
      : (timelineEntries.length === 1 ? "1 entry" : timelineEntries.length + " entries");
  }
  for (const entry of filtered) {
    const row = document.createElement("div");
    row.className = "hm-timeline-item";
    const changesCount = Array.isArray(entry.changes) ? entry.changes.length : 0;
    row.innerHTML = `
      <div class="hm-timeline-title">${escapeHtml(entry.label || "Change")}</div>
      <div class="hm-timeline-meta">
        <span>${escapeHtml(formatDateTime(entry.timestamp))}</span>
        <span>${changesCount} item${changesCount === 1 ? "" : "s"}</span>
        <button class="hm-timeline-undo" data-undo-id="${escapeHtml(entry.id || "")}">Undo</button>
      </div>
    `;
    const undoBtn = row.querySelector("[data-undo-id]");
    if (undoBtn) {
      undoBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        const id = undoBtn.getAttribute("data-undo-id");
        if (id) undoTimelineEntry(id);
      });
    }
    timelineList.appendChild(row);
  }

  emptyTimeline.style.display = filtered.length === 0 ? "flex" : "none";
}

function normalizeNoteText(note) {
  return String(note || "").replace(/\r\n/g, "\n").trim();
}

function normalizeNoteForSave(note) {
  return String(note || "").replace(/\r\n/g, "\n");
}

function getNotePreview(note) {
  const clean = String(note || "").replace(/\s+/g, " ").trim();
  if (!clean) return "";
  return clean.length > 120 ? clean.slice(0, 117).trimEnd() + "..." : clean;
}

async function saveNoteForHighlight(record, nextNoteRaw) {
  const id = normalizeHighlightId(record && record.id);
  if (!id) {
    showToast("Could not save note");
    return;
  }
  const domain = resolveRecordDomain(record);
  if (!domain) {
    showToast("Could not save note");
    return;
  }
  const fromStore = getHighlightById(id);
  const base = fromStore ? fromStore : record;
  const nextNote = normalizeNoteForSave(nextNoteRaw);
  const payload = {
    ...base,
    domain,
    note: nextNote
  };
  const res = await sendMsg({
    action: "saveHighlight",
    highlight: payload,
    historyMeta: {
      action: "update",
      label: nextNote ? "Updated highlight note" : "Cleared highlight note"
    }
  });
  if (!res || !res.ok) {
    showToast("Could not save note");
    return;
  }
  await loadData();
  await loadTimeline();
  render();
  showToast(nextNote ? "Note saved" : "Note cleared");
}

function openHighlightNoteDialog(record) {
  const existingNote = normalizeNoteForSave(record && record.note);
  openDialog({
    title: "Edit Note",
    multiline: true,
    initialValue: existingNote,
    placeholder: "Write a personal note for this highlight...",
    confirmLabel: "Save",
    onConfirm: async (noteText) => {
      await saveNoteForHighlight(record, noteText);
    }
  });
}

function createHighlightItem(h) {
  const isFav = favorites.includes(h.id);
  const isSelected = selectedHighlightIds.has(h.id);
  const colorSolid = getColorSolid(h.color);
  const label = getItemLabel(h);
  const notePreview = getNotePreview(h.note || "");
  const hasNote = !!notePreview;
  const isTruncated = wordCount(label) > 15;
  const itemType = h.type === "shape-cover" ? "Cover" : (h.type === "shape-rect" ? "Rectangle" : "Highlight");
  const metaSecondary = h.pageTitle || h.domain || "";

  const item = document.createElement("div");
  item.className = "hm-item" + (isSelected ? " selected" : "");
  item.setAttribute("data-hm-id", h.id);

  item.innerHTML = `
    <div class="hm-item-select-wrap">
      <input type="checkbox" class="hm-item-select" data-select-id="${h.id}" ${isSelected ? "checked" : ""} aria-label="Select item">
    </div>
    <div class="hm-item-color-bar" style="background: ${colorSolid};"></div>
    <div class="hm-item-body">
      <div class="hm-item-type">${itemType}</div>
      <div class="hm-item-text${isTruncated ? ' truncated' : ''}">${isTruncated ? truncateWords(label, 15) : escapeHtml(label)}</div>
      <div class="hm-item-meta">
        <span>${formatDate(h.timestamp)}</span>
        ${metaSecondary ? `<span class="hm-page-title">${escapeHtml(metaSecondary)}</span>` : ""}
      </div>
      ${hasNote ? `<div class="hm-item-note" title="${escapeHtml(normalizeNoteText(h.note || ""))}">${escapeHtml(notePreview)}</div>` : ""}
    </div>
    <div class="hm-item-actions">
      <button class="hm-icon-btn copy" data-copy title="Copy"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg></button>
      <button class="hm-icon-btn hm-note-btn ${hasNote ? "active" : ""}" data-note-id="${h.id}" title="${hasNote ? "Edit note" : "Add note"}"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg></button>
      <button class="hm-icon-btn hm-star-btn ${isFav ? "active" : ""}" data-fav-id="${h.id}" title="${isFav ? "Unpin" : "Pin"}"><svg width="13" height="13" viewBox="0 0 24 24" fill="${isFav ? "currentColor" : "none"}" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 17v5"/><path d="M8 10V3h10.2a1.8 1.8 0 0 1 1.8 1.8v8.8L17 11l-3 3-2-2-2 2z"/></svg></button>
      <button class="hm-icon-btn danger" data-delete-id="${h.id}" data-delete-domain="${h.domain}" title="Delete"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button>
    </div>
    ${isTruncated ? `<div class="hm-tooltip">${escapeHtml(label)}</div>` : ""}
  `;

  const selectBox = item.querySelector("[data-select-id]");
  if (selectBox) {
    selectBox.addEventListener("click", (e) => e.stopPropagation());
    selectBox.addEventListener("change", () => {
      if (selectBox.checked) selectedHighlightIds.add(h.id);
      else selectedHighlightIds.delete(h.id);
      item.classList.toggle("selected", !!selectBox.checked);
      updateBulkBar();
    });
  }

  item.addEventListener("click", (e) => {
    if (e.target.closest("input[type='checkbox']")) return;
    if (e.target.closest("button")) return;
    sendMsg({ action: "jumpToHighlight", highlight: h });
  });

  item.querySelector("[data-copy]").addEventListener("click", (e) => { e.stopPropagation(); copyToClipboard(label); });
  item.querySelector("[data-note-id]").addEventListener("click", (e) => {
    e.stopPropagation();
    openHighlightNoteDialog(h);
  });
  item.querySelector("[data-fav-id]").addEventListener("click", async (e) => {
    e.stopPropagation();
    await sendMsg({ action: "toggleFavorite", id: h.id });
    await loadData();
    await loadTimeline();
    render();
  });
  item.querySelector("[data-delete-id]").addEventListener("click", async (e) => {
    e.stopPropagation();
    const target = buildDeleteTarget(h);
    if (!target) {
      showToast("Could not delete item");
      return;
    }
    const res = await sendMsg({ action: "deleteHighlight", domain: target.domain, id: target.id });
    const removedCount = res && typeof res.removed === "number" ? res.removed : 0;
    if (res && res.ok && removedCount > 0) {
      selectedHighlightIds.delete(h.id);
      await loadData();
      await loadTimeline();
      render();
      showDeleteUndoToast(res.historyId || null, "Deleted 1 item");
    } else {
      await loadData();
      render();
      showToast("Could not delete item");
    }
  });

  return item;
}

// ---- Setup Actions ----

tabs.forEach(tab => {
  tab.addEventListener("click", async () => {
    tabs.forEach(t => t.classList.remove("active"));
    tab.classList.add("active");
    currentTab = tab.dataset.tab;
    panelAll.classList.toggle("active", currentTab === "all");
    panelFavorites.classList.toggle("active", currentTab === "favorites");
    panelTimeline.classList.toggle("active", currentTab === "timeline");
    if (currentTab === "timeline") {
      await loadTimeline();
    }
    render();
  });
});

searchInput.addEventListener("input", () => {
  const nextQuery = searchInput.value.trim().toLowerCase();
  clearTimeout(searchDebounceTimer);
  searchDebounceTimer = setTimeout(() => {
    searchQuery = nextQuery;
    render();
  }, currentTab === "timeline" ? 140 : 80);
});

if (filterColor) {
  filterColor.addEventListener("change", () => {
    activeColorFilter = filterColor.value || "all";
    render();
  });
}

if (sortOrder) {
  sortOrder.addEventListener("change", () => {
    activeSortOrder = sortOrder.value || "date_desc";
    render();
  });
}

if (filterCategory) {
  filterCategory.addEventListener("change", () => {
    activeCategoryFilter = filterCategory.value || "all";
    render();
  });
}

if (btnSelectVisible) {
  btnSelectVisible.addEventListener("click", () => {
    const visibleIds = Array.from(document.querySelectorAll(".hm-tab-panel.active .hm-item[data-hm-id]"))
      .map((el) => el.getAttribute("data-hm-id"))
      .filter(Boolean);
    for (const id of visibleIds) {
      selectedHighlightIds.add(id);
    }
    render();
  });
}

if (btnClearSelection) {
  btnClearSelection.addEventListener("click", () => {
    selectedHighlightIds.clear();
    render();
  });
}

if (bulkColors) {
  bulkColors.addEventListener("click", (e) => {
    const button = e.target.closest("[data-bulk-color]");
    if (!button) return;
    const colorName = button.getAttribute("data-bulk-color");
    if (!colorName) return;
    applyBulkColor(colorName);
  });
}

if (btnBulkDelete) {
  btnBulkDelete.addEventListener("click", async () => {
    if (!selectedHighlightIds.size) return;
    await deleteSelectedHighlights();
  });
}

if (btnBulkMarkdown) {
  btnBulkMarkdown.addEventListener("click", () => {
    exportSelectedToMarkdown();
  });
}

if (btnUndoLast) {
  btnUndoLast.addEventListener("click", () => {
    undoLastForCurrentPage();
  });
}

toggleEnabled.addEventListener("change", () => {
  sendMsg({ action: "setEnabled", enabled: toggleEnabled.checked });
  if (!toggleEnabled.checked) {
    activeDrawMode = null;
    continuousModeEnabled = false;
    updateDrawModeButtons();
    updateContinuousModeUi();
  } else {
    syncDrawModeFromActiveTab();
    syncContinuousModeFromActiveTab();
  }
  showToast(toggleEnabled.checked ? "✓ Enabled" : "⏸ Paused");
});

btnRectMode.addEventListener("click", () => {
  toggleDrawMode("shape-rect");
});

if (btnAutoOcrCopy) {
  btnAutoOcrCopy.addEventListener("click", () => {
    toggleAutoOcrCopySetting();
  });
}

if (btnOpenDashboard) {
  btnOpenDashboard.addEventListener("click", () => {
    openDashboardInTab();
  });
}

if (btnContinuousMode) {
  btnContinuousMode.addEventListener("click", () => {
    toggleContinuousMode();
  });
}

if (continuousColorSelect) {
  continuousColorSelect.addEventListener("change", () => {
    updateContinuousModeUi();
    updateContinuousColorOnActiveTab();
  });
}

if (continuousCustomColor) {
  continuousCustomColor.addEventListener("change", () => {
    updateContinuousColorOnActiveTab();
  });
}

const refreshPopupContext = async () => {
  await refreshActiveTabContext();
  await syncDrawModeFromActiveTab();
  await syncContinuousModeFromActiveTab();
  await syncAutoOcrCopySetting();
  await loadData();
  await loadTimeline();
  render();
};

const onTabActivated = () => { refreshPopupContext(); };
const onTabUpdated = (_tabId, changeInfo) => {
  if (!changeInfo || changeInfo.status !== "complete") return;
  if (currentTabId && _tabId !== currentTabId) return;
  if (!document.hidden) {
    refreshPopupContext();
  }
};
chrome.tabs.onActivated.addListener(onTabActivated);
chrome.tabs.onUpdated.addListener(onTabUpdated);
window.addEventListener("focus", () => { refreshPopupContext(); });
document.addEventListener("visibilitychange", () => {
  if (!document.hidden) refreshPopupContext();
});
window.addEventListener("unload", () => {
  try { chrome.tabs.onActivated.removeListener(onTabActivated); } catch (_e) {}
  try { chrome.tabs.onUpdated.removeListener(onTabUpdated); } catch (_e) {}
});

btnExport.addEventListener("click", async () => {
  const res = await sendMsg({ action: "exportAll" });
  if (res.ok) exportBlob(res.payload, "highlightmaster-complete-backup.json");
});

if (btnDeletedCsv) {
  btnDeletedCsv.addEventListener("click", async () => {
    const res = await sendMsg({ action: "exportDeletedCsv" });
    if (!res || !res.ok || typeof res.csv !== "string") {
      showToast("Deleted log export failed");
      return;
    }
    downloadTextFile(res.csv, "highlightmaster-deleted-log.csv", "text/csv;charset=utf-8");
    const rowCount = Number(res.rows) || 0;
    showToast("Deleted CSV exported (" + rowCount + " row" + (rowCount === 1 ? "" : "s") + ")");
  });
}

if (btnRestoreDeleted) {
  btnRestoreDeleted.addEventListener("click", async () => {
    await openRestoreDeletedDialog();
  });
}

btnImport.addEventListener("click", () => importFile.click());
importFile.addEventListener("change", async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  try {
    const text = await file.text();
    const payload = JSON.parse(text);
    const res = await sendMsg({ action: "importAll", payload });
    if (res.ok) {
      await loadData();
      await loadTimeline();
      render();
      showToast("📥 Imported " + (res.imported || 0) + " items");
    } else showToast("⚠ Import failed");
  } catch (err) { showToast("⚠ Invalid file"); }
  importFile.value = "";
});

init();
