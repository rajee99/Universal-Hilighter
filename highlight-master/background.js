"use strict";

/*
 * HighlightMaster - Background Service Worker
 * Handles storage, message routing, tab management, export/import
 */

// ---- Default State ----

const LEGACY_DEFAULT_CATEGORY_IDS = new Set([
  "default_study",
  "default_work",
  "default_important",
  "default_review",
  "default_questions"
]);
const LEGACY_DEFAULT_CATEGORY_NAMES = new Set([
  "study",
  "work",
  "important",
  "review",
  "questions"
]);
const LEGACY_DEFAULT_PURGE_KEY = "legacyDefaultCategoriesPurgedV1";

const HISTORY_STORAGE_KEY = "timelineHistory";
const HISTORY_LIMIT = 120;
const DELETED_AUDIT_STORAGE_KEY = "deletedAuditLog";
const DELETED_AUDIT_LIMIT = 2000;
const DOMAIN_ASSIGNMENT_CUTOFFS_KEY = "domainAssignmentCutoffs";
const SHAPE_TYPES = new Set(["shape-rect", "shape-cover"]);
const LOCAL_BACKUP_PREFIX = "__hm_backup__";
const CRITICAL_LOCAL_KEYS = [
  "highlights",
  "favorites",
  "categories",
  "domainAssignments",
  DOMAIN_ASSIGNMENT_CUTOFFS_KEY,
  HISTORY_STORAGE_KEY,
  DELETED_AUDIT_STORAGE_KEY
];

function deepClone(value) {
  if (value === null || typeof value === "undefined") return value;
  return JSON.parse(JSON.stringify(value));
}

function uniq(values) {
  return Array.from(new Set(values));
}

function nowIso() {
  return new Date().toISOString();
}

function sanitizeNumber(value, fallback = 0) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function sanitizeRect(rect) {
  if (!rect || typeof rect !== "object") return null;
  const left = sanitizeNumber(rect.left, NaN);
  const top = sanitizeNumber(rect.top, NaN);
  const width = sanitizeNumber(rect.width, NaN);
  const height = sanitizeNumber(rect.height, NaN);
  if (![left, top, width, height].every(Number.isFinite)) return null;
  if (width <= 0 || height <= 0) return null;
  return {
    left,
    top,
    width,
    height
  };
}

function sanitizeHighlightRecord(h) {
  if (!h || typeof h !== "object") return null;
  if (typeof h.id !== "string" || !h.id.trim()) return null;
  if (typeof h.domain !== "string" || !h.domain.trim()) return null;

  const next = deepClone(h);
  const type = typeof next.type === "string" ? next.type.trim() : "";

  if (type && type !== "text" && !SHAPE_TYPES.has(type)) {
    return null;
  }

  if (SHAPE_TYPES.has(type)) {
    const rect = sanitizeRect(next.rect);
    if (!rect) return null;
    next.rect = rect;

    if (type === "shape-rect") {
      if (typeof next.revealed !== "boolean") next.revealed = false;
      if (typeof next.transient !== "boolean") next.transient = false;
    }
  }

  return next;
}

function normalizeHighlightId(id) {
  if (id === null || typeof id === "undefined") return "";
  return String(id).trim();
}

function normalizeHighlightRecordForStorage(record, domainFallback = "") {
  if (!record || typeof record !== "object") return null;
  const candidate = deepClone(record);
  candidate.id = normalizeHighlightId(candidate.id);
  if (!candidate.id) return null;

  const fallbackDomain = typeof domainFallback === "string" ? domainFallback.trim() : "";
  if (typeof candidate.domain !== "string" || !candidate.domain.trim()) {
    candidate.domain = fallbackDomain;
  } else {
    candidate.domain = candidate.domain.trim();
  }

  return sanitizeHighlightRecord(candidate);
}

function normalizeHighlightsMap(rawHighlights) {
  if (!rawHighlights || typeof rawHighlights !== "object") return {};
  const normalized = {};
  const idIndexByDomain = {};

  for (const [domainKey, items] of Object.entries(rawHighlights)) {
    if (!Array.isArray(items)) continue;
    const domainFallback = typeof domainKey === "string" ? domainKey.trim() : "";
    for (const item of items) {
      const next = normalizeHighlightRecordForStorage(item, domainFallback);
      if (!next) continue;
      const domain = next.domain;
      if (!domain) continue;
      if (!normalized[domain]) {
        normalized[domain] = [];
        idIndexByDomain[domain] = new Map();
      }
      const domainIndex = idIndexByDomain[domain];
      const existingIndex = domainIndex.get(next.id);
      if (typeof existingIndex === "number") {
        normalized[domain][existingIndex] = next;
      } else {
        domainIndex.set(next.id, normalized[domain].length);
        normalized[domain].push(next);
      }
    }
  }

  return normalized;
}

function getDeletedAuditRowKey(row) {
  if (!row || typeof row !== "object") return "";
  if (typeof row.rowId === "string" && row.rowId) return row.rowId;
  return [
    "audit",
    row.deletedAtIso || "",
    row.highlightId || "",
    row.historyId || "",
    row.domain || ""
  ].join("|");
}

function normalizeHistoryChanges(changes) {
  if (!Array.isArray(changes)) return [];
  const normalized = [];
  for (const change of changes) {
    if (!change || !change.domain || !change.id) continue;
    normalized.push({
      domain: change.domain,
      id: change.id,
      before: deepClone(change.before || null),
      after: deepClone(change.after || null)
    });
  }
  return normalized;
}

function summarizeAction(action, count) {
  const c = count || 1;
  switch (action) {
    case "create":
      return c === 1 ? "Created 1 highlight" : "Created " + c + " highlights";
    case "update":
      return c === 1 ? "Updated 1 highlight" : "Updated " + c + " highlights";
    case "delete":
      return c === 1 ? "Deleted 1 highlight" : "Deleted " + c + " highlights";
    case "bulk_color":
      return c === 1 ? "Changed color of 1 highlight" : "Changed color of " + c + " highlights";
    case "bulk_delete":
      return c === 1 ? "Deleted 1 selected highlight" : "Deleted " + c + " selected highlights";
    case "clear_domain":
      return c === 1 ? "Cleared 1 highlight from domain" : "Cleared " + c + " highlights from domain";
    case "clear_rectangles":
      return c === 1 ? "Deleted 1 rectangle" : "Deleted " + c + " rectangles";
    case "resize_rect":
      return c === 1 ? "Resized 1 rectangle" : "Resized " + c + " rectangles";
    case "move_rect":
      return c === 1 ? "Moved 1 rectangle" : "Moved " + c + " rectangles";
    case "restore_deleted":
      return c === 1 ? "Restored 1 deleted item" : "Restored " + c + " deleted items";
    default:
      return c === 1 ? "Changed 1 highlight" : "Changed " + c + " highlights";
  }
}

function buildHistoryEntry(meta, changes) {
  const normalizedChanges = normalizeHistoryChanges(changes);
  if (!normalizedChanges.length) return null;

  const changedDomains = uniq(normalizedChanges.map((c) => c.domain));
  const pageUrlFromChange = normalizedChanges.find((c) => c.after && c.after.url) || normalizedChanges.find((c) => c.before && c.before.url);
  const pageUrl = (meta && meta.pageUrl) || (pageUrlFromChange ? (pageUrlFromChange.after ? pageUrlFromChange.after.url : pageUrlFromChange.before.url) : "");
  const domain = (meta && meta.domain) || changedDomains[0] || "";
  const action = (meta && meta.action) || "update";
  const label = (meta && meta.label) || summarizeAction(action, normalizedChanges.length);

  return {
    id: "hist-" + Date.now() + "-" + Math.random().toString(36).slice(2, 8),
    action,
    label,
    timestamp: Date.now(),
    timestampIso: nowIso(),
    domain,
    pageUrl,
    changedDomains,
    changes: normalizedChanges
  };
}

function cleanupLegacyDefaultCategories(categories, domainAssignments, options = {}) {
  const aggressiveNames = !!(options && options.aggressiveNames);
  const nextCategories = categories && typeof categories === "object"
    ? { ...categories }
    : {};
  const nextAssignments = domainAssignments && typeof domainAssignments === "object"
    ? { ...domainAssignments }
    : {};
  let categoriesChanged = false;
  let assignmentsChanged = false;

  for (const [categoryId, category] of Object.entries(nextCategories)) {
    const id = String(categoryId);
    const normalizedName = category && typeof category.name === "string"
      ? category.name.trim().toLowerCase()
      : "";
    const isLegacyId = LEGACY_DEFAULT_CATEGORY_IDS.has(id);
    const isLegacyName = LEGACY_DEFAULT_CATEGORY_NAMES.has(normalizedName);
    if (!isLegacyId && !(aggressiveNames && isLegacyName)) continue;
    delete nextCategories[id];
    categoriesChanged = true;
  }

  // Clean invalid parent links after removals.
  for (const [categoryId, category] of Object.entries(nextCategories)) {
    if (!category || typeof category !== "object") continue;
    if (!category.parentId || nextCategories[category.parentId]) continue;
    nextCategories[categoryId] = { ...category, parentId: null };
    categoriesChanged = true;
  }

  for (const [domain, categoryId] of Object.entries(nextAssignments)) {
    if (!categoryId || !nextCategories[categoryId]) {
      delete nextAssignments[domain];
      assignmentsChanged = true;
    }
  }

  return {
    categories: nextCategories,
    domainAssignments: nextAssignments,
    categoriesChanged,
    assignmentsChanged
  };
}

function getLocalBackupKey(key) {
  return LOCAL_BACKUP_PREFIX + key;
}

function isPlainObject(value) {
  return !!value && typeof value === "object" && !Array.isArray(value);
}

function isValidCriticalLocalValue(key, value) {
  if (key === "favorites" || key === HISTORY_STORAGE_KEY || key === DELETED_AUDIT_STORAGE_KEY) {
    return Array.isArray(value);
  }
  if (key === "highlights" || key === "categories" || key === "domainAssignments" || key === DOMAIN_ASSIGNMENT_CUTOFFS_KEY) {
    return isPlainObject(value);
  }
  return true;
}

async function restoreCriticalLocalStateFromBackup() {
  const wantedKeys = [];
  for (const key of CRITICAL_LOCAL_KEYS) {
    wantedKeys.push(key, getLocalBackupKey(key));
  }
  const snapshot = await storageLocalGetMany(wantedKeys);
  const payload = {};
  let changed = false;

  for (const key of CRITICAL_LOCAL_KEYS) {
    const backupKey = getLocalBackupKey(key);
    const currentValue = Object.prototype.hasOwnProperty.call(snapshot, key) ? snapshot[key] : null;
    const backupValue = Object.prototype.hasOwnProperty.call(snapshot, backupKey) ? snapshot[backupKey] : null;
    const currentValid = isValidCriticalLocalValue(key, currentValue);
    const backupValid = isValidCriticalLocalValue(key, backupValue);

    if (!currentValid && backupValid) {
      payload[key] = deepClone(backupValue);
      payload[backupKey] = deepClone(backupValue);
      changed = true;
      continue;
    }

    if (currentValid && !backupValid) {
      payload[backupKey] = deepClone(currentValue);
      changed = true;
    }
  }

  if (changed) {
    await storageLocalSetRaw(payload);
  }
}

async function ensureDefaultState() {
  await restoreCriticalLocalStateFromBackup();

  const enabled = await storageSyncGet("enabled");
  if (typeof enabled !== "boolean") {
    await storageSyncSet({ enabled: false });
  }
  const autoOcrCopyEnabled = await storageSyncGet("autoOcrCopyEnabled");
  if (typeof autoOcrCopyEnabled !== "boolean") {
    await storageSyncSet({ autoOcrCopyEnabled: false });
  }

  const highlights = await storageLocalGet("highlights");
  if (!highlights || typeof highlights !== "object") {
    await storageLocalSet({ highlights: {} });
  }

  const favorites = await storageLocalGet("favorites");
  if (!Array.isArray(favorites)) {
    await storageLocalSet({ favorites: [] });
  }

  const categoriesRaw = await storageLocalGet("categories");
  let categories = categoriesRaw && typeof categoriesRaw === "object" && !Array.isArray(categoriesRaw)
    ? categoriesRaw
    : {};
  let categoriesChanged = !categoriesRaw || typeof categoriesRaw !== "object" || Array.isArray(categoriesRaw);

  const domainAssignmentsRaw = await storageLocalGet("domainAssignments");
  let domainAssignments = domainAssignmentsRaw && typeof domainAssignmentsRaw === "object" && !Array.isArray(domainAssignmentsRaw)
    ? domainAssignmentsRaw
    : {};
  let assignmentsChanged = !domainAssignmentsRaw || typeof domainAssignmentsRaw !== "object" || Array.isArray(domainAssignmentsRaw);

  const domainAssignmentCutoffsRaw = await storageLocalGet(DOMAIN_ASSIGNMENT_CUTOFFS_KEY);
  let domainAssignmentCutoffs = domainAssignmentCutoffsRaw && typeof domainAssignmentCutoffsRaw === "object" && !Array.isArray(domainAssignmentCutoffsRaw)
    ? domainAssignmentCutoffsRaw
    : {};
  let cutoffsChanged = !domainAssignmentCutoffsRaw || typeof domainAssignmentCutoffsRaw !== "object" || Array.isArray(domainAssignmentCutoffsRaw);

  const cleanupResult = cleanupLegacyDefaultCategories(categories, domainAssignments);
  categories = cleanupResult.categories;
  domainAssignments = cleanupResult.domainAssignments;
  categoriesChanged = categoriesChanged || cleanupResult.categoriesChanged;
  assignmentsChanged = assignmentsChanged || cleanupResult.assignmentsChanged;

  const legacyPurgeDone = await storageLocalGet(LEGACY_DEFAULT_PURGE_KEY);
  let shouldMarkLegacyPurgeDone = false;
  if (legacyPurgeDone !== true) {
    const aggressiveCleanupResult = cleanupLegacyDefaultCategories(categories, domainAssignments, { aggressiveNames: true });
    categories = aggressiveCleanupResult.categories;
    domainAssignments = aggressiveCleanupResult.domainAssignments;
    categoriesChanged = categoriesChanged || aggressiveCleanupResult.categoriesChanged;
    assignmentsChanged = assignmentsChanged || aggressiveCleanupResult.assignmentsChanged;
    shouldMarkLegacyPurgeDone = true;
  }

  const now = Date.now();
  for (const domain of Object.keys(domainAssignments)) {
    if (!Object.prototype.hasOwnProperty.call(domainAssignmentCutoffs, domain)) {
      domainAssignmentCutoffs[domain] = now;
      cutoffsChanged = true;
    } else {
      const cutoffValue = Number(domainAssignmentCutoffs[domain]);
      if (!Number.isFinite(cutoffValue) || cutoffValue <= 0) {
        domainAssignmentCutoffs[domain] = now;
        cutoffsChanged = true;
      }
    }
  }
  for (const domain of Object.keys(domainAssignmentCutoffs)) {
    if (domainAssignments[domain]) continue;
    delete domainAssignmentCutoffs[domain];
    cutoffsChanged = true;
  }

  if (categoriesChanged) {
    await storageLocalSet({ categories });
  }
  if (assignmentsChanged) {
    await storageLocalSet({ domainAssignments });
  }
  if (cutoffsChanged) {
    const payload = {};
    payload[DOMAIN_ASSIGNMENT_CUTOFFS_KEY] = domainAssignmentCutoffs;
    await storageLocalSet(payload);
  }
  if (shouldMarkLegacyPurgeDone) {
    const payload = {};
    payload[LEGACY_DEFAULT_PURGE_KEY] = true;
    await storageLocalSet(payload);
  }

  const timelineHistory = await storageLocalGet(HISTORY_STORAGE_KEY);
  if (!Array.isArray(timelineHistory)) {
    const payload = {};
    payload[HISTORY_STORAGE_KEY] = [];
    await storageLocalSet(payload);
  }

  const deletedAuditLog = await storageLocalGet(DELETED_AUDIT_STORAGE_KEY);
  if (!Array.isArray(deletedAuditLog)) {
    const payload = {};
    payload[DELETED_AUDIT_STORAGE_KEY] = [];
    await storageLocalSet(payload);
  }
}

chrome.runtime.onInstalled.addListener(() => {
  ensureDefaultState().catch(() => {});
});

ensureDefaultState().catch(() => {});

// ---- Programmatic Injection Fallback ----

chrome.action.onClicked.addListener(() => {});

chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
  if (changeInfo.status === "complete" && tab.url && /^https?:/.test(tab.url)) {
    chrome.tabs.sendMessage(tabId, { action: "ping" }).catch(() => {
      chrome.scripting.executeScript({
        target: { tabId },
        files: ["content.js"]
      }).catch(() => {});
    });
  }
});

// ---- Storage Helpers ----

function storageLocalGet(key) {
  return new Promise((resolve) => {
    chrome.storage.local.get(key, (res) => {
      resolve(res[key] ?? null);
    });
  });
}

function storageLocalGetMany(keys) {
  return new Promise((resolve) => {
    chrome.storage.local.get(keys, (res) => {
      resolve(res || {});
    });
  });
}

function storageLocalSetRaw(obj) {
  return new Promise((resolve) => {
    chrome.storage.local.set(obj, resolve);
  });
}

function storageLocalSet(obj) {
  const payload = obj && typeof obj === "object" ? { ...obj } : {};
  for (const key of CRITICAL_LOCAL_KEYS) {
    if (!Object.prototype.hasOwnProperty.call(payload, key)) continue;
    payload[getLocalBackupKey(key)] = deepClone(payload[key]);
  }
  return storageLocalSetRaw(payload);
}

function storageSyncGet(key) {
  return new Promise((resolve) => {
    chrome.storage.sync.get(key, (res) => {
      resolve(res[key] ?? null);
    });
  });
}

function storageSyncSet(obj) {
  return new Promise((resolve) => {
    chrome.storage.sync.set(obj, resolve);
  });
}

async function getAllHighlights() {
  const data = await storageLocalGet("highlights");
  return normalizeHighlightsMap(data && typeof data === "object" ? data : {});
}

async function setAllHighlights(highlights) {
  await storageLocalSet({ highlights: normalizeHighlightsMap(highlights) });
}

async function getFavorites() {
  const data = await storageLocalGet("favorites");
  if (!Array.isArray(data)) return [];
  return uniq(data.map((value) => normalizeHighlightId(value)).filter(Boolean));
}

async function setFavorites(favorites) {
  const normalizedFavorites = Array.isArray(favorites)
    ? uniq(favorites.map((value) => normalizeHighlightId(value)).filter(Boolean))
    : [];
  await storageLocalSet({ favorites: normalizedFavorites });
}

async function getCategories() {
  const data = await storageLocalGet("categories");
  if (!data || typeof data !== "object" || Array.isArray(data)) {
    return {};
  }
  return data;
}

async function setCategories(cats) {
  await storageLocalSet({ categories: cats });
}

async function getDomainAssignments() {
  const data = await storageLocalGet("domainAssignments");
  return data && typeof data === "object" && !Array.isArray(data) ? data : {};
}

async function setDomainAssignments(assignments) {
  await storageLocalSet({ domainAssignments: assignments });
}

async function getDomainAssignmentCutoffs() {
  const data = await storageLocalGet(DOMAIN_ASSIGNMENT_CUTOFFS_KEY);
  return data && typeof data === "object" && !Array.isArray(data) ? data : {};
}

async function setDomainAssignmentCutoffs(cutoffs) {
  const payload = {};
  payload[DOMAIN_ASSIGNMENT_CUTOFFS_KEY] = cutoffs;
  await storageLocalSet(payload);
}

async function getTimelineHistory() {
  const data = await storageLocalGet(HISTORY_STORAGE_KEY);
  return Array.isArray(data) ? data : [];
}

async function setTimelineHistory(history) {
  const payload = {};
  payload[HISTORY_STORAGE_KEY] = Array.isArray(history) ? history : [];
  await storageLocalSet(payload);
}

async function appendTimelineEntry(entry) {
  if (!entry || !Array.isArray(entry.changes) || entry.changes.length === 0) return;
  const history = await getTimelineHistory();
  history.push(entry);
  if (history.length > HISTORY_LIMIT) {
    history.splice(0, history.length - HISTORY_LIMIT);
  }
  await setTimelineHistory(history);
}

async function getDeletedAuditLog() {
  const data = await storageLocalGet(DELETED_AUDIT_STORAGE_KEY);
  return Array.isArray(data) ? data : [];
}

async function setDeletedAuditLog(rows) {
  const payload = {};
  payload[DELETED_AUDIT_STORAGE_KEY] = Array.isArray(rows) ? rows : [];
  await storageLocalSet(payload);
}

function buildDeletedAuditRowsFromEntry(entry) {
  if (!entry || !Array.isArray(entry.changes)) return [];
  const rows = [];
  for (const change of entry.changes) {
    if (!change || !change.before || change.after) continue;
    const before = sanitizeHighlightRecord(change.before) || deepClone(change.before);
    const rect = before && before.rect && typeof before.rect === "object" ? before.rect : null;
    rows.push({
      rowId: "del-" + Date.now() + "-" + Math.random().toString(36).slice(2, 8),
      deletedAt: entry.timestamp || Date.now(),
      deletedAtIso: entry.timestampIso || nowIso(),
      action: entry.action || "delete",
      historyId: entry.id || "",
      domain: change.domain || before.domain || "",
      pageUrl: before.url || entry.pageUrl || "",
      highlightId: change.id || before.id || "",
      type: before.type || "text",
      color: before.color || "",
      bgColor: before.bgColor || "",
      revealed: typeof before.revealed === "boolean" ? String(before.revealed) : "",
      label: before.label || "",
      text: before.text || "",
      ocrText: before.ocrText || "",
      pageTitle: before.pageTitle || "",
      rectLeft: rect ? String(rect.left ?? "") : "",
      rectTop: rect ? String(rect.top ?? "") : "",
      rectWidth: rect ? String(rect.width ?? "") : "",
      rectHeight: rect ? String(rect.height ?? "") : "",
      restoredAtIso: "",
      restoredAt: "",
      restoreHistoryId: "",
      snapshotJson: JSON.stringify(before)
    });
  }
  return rows;
}

async function appendDeletedAuditRows(rows) {
  if (!Array.isArray(rows) || rows.length === 0) return;
  const log = await getDeletedAuditLog();
  log.push(...rows);
  if (log.length > DELETED_AUDIT_LIMIT) {
    log.splice(0, log.length - DELETED_AUDIT_LIMIT);
  }
  await setDeletedAuditLog(log);
}

function csvEscape(value) {
  const str = value === null || typeof value === "undefined" ? "" : String(value);
  if (/[",\n]/.test(str)) return '"' + str.replace(/"/g, '""') + '"';
  return str;
}

function buildDeletedAuditCsv(rows) {
  const header = [
    "rowId",
    "deletedAtIso",
    "deletedAt",
    "action",
    "historyId",
    "domain",
    "pageUrl",
    "highlightId",
    "type",
    "color",
    "bgColor",
    "revealed",
    "label",
    "text",
    "ocrText",
    "pageTitle",
    "rectLeft",
    "rectTop",
    "rectWidth",
    "rectHeight",
    "restoredAtIso",
    "restoredAt",
    "restoreHistoryId",
    "snapshotJson"
  ];
  const lines = [header.join(",")];
  for (const row of rows) {
    lines.push(header.map((col) => csvEscape(row && row[col])).join(","));
  }
  return lines.join("\n");
}

function getHighlightByDomainAndId(allHighlights, domain, id) {
  if (!allHighlights || !domain || id === null || typeof id === "undefined") return null;
  const normalizedId = normalizeHighlightId(id);
  if (!normalizedId) return null;
  const domainItems = allHighlights[domain];
  if (!Array.isArray(domainItems)) return null;
  const found = domainItems.find((h) => h && normalizeHighlightId(h.id) === normalizedId);
  return found ? deepClone(found) : null;
}

function resolveDomainForHighlightId(allHighlights, id, preferredDomain = "") {
  if (!allHighlights || id === null || typeof id === "undefined") return "";
  const normalizedId = normalizeHighlightId(id);
  if (!normalizedId) return "";
  const wantedDomain = typeof preferredDomain === "string" ? preferredDomain : "";

  if (wantedDomain) {
    const preferredItems = allHighlights[wantedDomain];
    if (Array.isArray(preferredItems) && preferredItems.some((h) => h && normalizeHighlightId(h.id) === normalizedId)) {
      return wantedDomain;
    }
  }

  for (const [domain, items] of Object.entries(allHighlights)) {
    if (!Array.isArray(items)) continue;
    if (items.some((h) => h && normalizeHighlightId(h.id) === normalizedId)) return domain;
  }

  return "";
}

function setHighlightByDomainAndId(allHighlights, domain, id, nextHighlight) {
  if (!allHighlights || !domain || id === null || typeof id === "undefined") return;
  const normalizedId = normalizeHighlightId(id);
  if (!normalizedId) return;
  const items = Array.isArray(allHighlights[domain]) ? allHighlights[domain] : [];
  const idx = items.findIndex((h) => h && normalizeHighlightId(h.id) === normalizedId);
  if (!nextHighlight) {
    if (idx >= 0) items.splice(idx, 1);
    if (items.length > 0) allHighlights[domain] = items;
    else delete allHighlights[domain];
    return;
  }
  const sanitized = normalizeHighlightRecordForStorage(nextHighlight, domain);
  if (!sanitized) return;
  if (idx >= 0) items[idx] = deepClone(sanitized);
  else items.push(deepClone(sanitized));
  allHighlights[domain] = items;
}

// ---- Highlight CRUD ----

function isValidHighlightRecord(h) {
  return !!sanitizeHighlightRecord(h);
}

async function saveHighlight(h, historyMeta = null) {
  const sanitizedHighlight = normalizeHighlightRecordForStorage(h, h && h.domain ? h.domain : "");
  if (!sanitizedHighlight) return;
  const all = await getAllHighlights();
  if (!all[sanitizedHighlight.domain]) all[sanitizedHighlight.domain] = [];
  const idx = all[sanitizedHighlight.domain].findIndex((x) => normalizeHighlightId(x && x.id) === sanitizedHighlight.id);
  const before = idx === -1 ? null : deepClone(all[sanitizedHighlight.domain][idx]);
  if (idx === -1) {
    all[sanitizedHighlight.domain].push(sanitizedHighlight);
  } else {
    all[sanitizedHighlight.domain][idx] = { ...all[sanitizedHighlight.domain][idx], ...sanitizedHighlight };
  }
  const after = idx === -1
    ? deepClone(sanitizedHighlight)
    : deepClone(all[sanitizedHighlight.domain][idx]);
  await setAllHighlights(all);

  if (before && JSON.stringify(before) === JSON.stringify(after)) {
    return;
  }

  const defaultAction = before
    ? (sanitizedHighlight.type === "shape-rect" && before && before.rect && sanitizedHighlight.rect ? "resize_rect" : "update")
    : "create";
  const entry = buildHistoryEntry({
    action: (historyMeta && historyMeta.action) || defaultAction,
    label: historyMeta && historyMeta.label ? historyMeta.label : null,
    domain: sanitizedHighlight.domain,
    pageUrl: sanitizedHighlight.url || ""
  }, [{
    domain: sanitizedHighlight.domain,
    id: sanitizedHighlight.id,
    before,
    after
  }]);
  await appendTimelineEntry(entry);
}

function buildHighlightFromDeletedAuditRow(row) {
  if (!row || typeof row !== "object") return null;

  if (typeof row.snapshotJson === "string" && row.snapshotJson) {
    try {
      const snapshot = JSON.parse(row.snapshotJson);
      const sanitizedSnapshot = normalizeHighlightRecordForStorage(snapshot, row.domain || "");
      if (sanitizedSnapshot) return sanitizedSnapshot;
    } catch (_err) {}
  }

  const fallbackType = typeof row.type === "string" && row.type ? row.type : "text";
  const fallback = {
    id: normalizeHighlightId(row.highlightId),
    domain: row.domain || "",
    url: row.pageUrl || "",
    type: fallbackType,
    color: row.color || "yellow",
    bgColor: row.bgColor || "",
    text: row.text || "",
    label: row.label || "",
    ocrText: row.ocrText || "",
    pageTitle: row.pageTitle || "",
    timestamp: sanitizeNumber(row.deletedAt, Date.now())
  };

  if (fallbackType === "shape-rect" || fallbackType === "shape-cover") {
    fallback.rect = {
      left: sanitizeNumber(row.rectLeft, NaN),
      top: sanitizeNumber(row.rectTop, NaN),
      width: sanitizeNumber(row.rectWidth, NaN),
      height: sanitizeNumber(row.rectHeight, NaN)
    };
    if (fallbackType === "shape-rect") {
      const revealedRaw = String(row.revealed || "").toLowerCase();
      fallback.revealed = revealedRaw === "true" || revealedRaw === "1";
      fallback.transient = false;
    }
  }

  return normalizeHighlightRecordForStorage(fallback, row.domain || "");
}

async function markDeletedAuditRowsRestored(rowKeys, restoreHistoryId = "") {
  const keySet = new Set((rowKeys || []).filter(Boolean));
  if (!keySet.size) return;

  const deletedAuditLog = await getDeletedAuditLog();
  if (!deletedAuditLog.length) return;

  let changed = false;
  const restoredAt = Date.now();
  const restoredAtIso = new Date(restoredAt).toISOString();
  const nextRows = deletedAuditLog.map((row) => {
    if (!keySet.has(getDeletedAuditRowKey(row))) return row;
    changed = true;
    return {
      ...row,
      restoredAt,
      restoredAtIso,
      restoreHistoryId: restoreHistoryId || row.restoreHistoryId || ""
    };
  });

  if (changed) {
    await setDeletedAuditLog(nextRows);
  }
}

async function restoreDeletedAuditRows(rowKeys) {
  const requestedKeys = uniq((Array.isArray(rowKeys) ? rowKeys : []).filter(Boolean));
  if (!requestedKeys.length) {
    return { ok: false, reason: "invalid_row_ids" };
  }

  const deletedAuditLog = await getDeletedAuditLog();
  const rowsByKey = new Map();
  for (const row of deletedAuditLog) {
    rowsByKey.set(getDeletedAuditRowKey(row), row);
  }

  const all = await getAllHighlights();
  const restoredKeys = [];
  const restoredDomains = new Set();
  const skipped = [];
  const changes = [];

  for (const key of requestedKeys) {
    const row = rowsByKey.get(key);
    if (!row) {
      skipped.push({ rowKey: key, reason: "not_found" });
      continue;
    }

    const highlight = buildHighlightFromDeletedAuditRow(row);
    if (!highlight) {
      skipped.push({ rowKey: key, reason: "invalid_snapshot" });
      continue;
    }

    if (getHighlightByDomainAndId(all, highlight.domain, highlight.id)) {
      skipped.push({ rowKey: key, reason: "already_exists" });
      continue;
    }

    setHighlightByDomainAndId(all, highlight.domain, highlight.id, highlight);
    restoredDomains.add(highlight.domain);
    restoredKeys.push(key);
    changes.push({
      domain: highlight.domain,
      id: highlight.id,
      before: null,
      after: deepClone(highlight)
    });
  }

  if (!changes.length) {
    return {
      ok: false,
      reason: skipped[0] ? skipped[0].reason : "nothing_to_restore",
      skipped
    };
  }

  await setAllHighlights(all);

  const entry = buildHistoryEntry({
    action: "restore_deleted",
    label: summarizeAction("restore_deleted", changes.length),
    domain: changes[0] ? changes[0].domain : "",
    pageUrl: changes[0] && changes[0].after ? (changes[0].after.url || "") : ""
  }, changes);

  if (entry) {
    await appendTimelineEntry(entry);
  }

  await markDeletedAuditRowsRestored(restoredKeys, entry ? entry.id : "");

  for (const domain of restoredDomains) {
    notifyTabsForDomain(domain, { action: "highlightsUpdated" });
  }

  return {
    ok: true,
    restored: changes.length,
    historyId: entry ? entry.id : null,
    skipped
  };
}

async function deleteHighlight(domain, id, silent = false, excludedTabId = null) {
  const normalizedId = normalizeHighlightId(id);
  if (!normalizedId) {
    return { ok: false, reason: "invalid_target", removed: 0, historyId: null };
  }
  const all = await getAllHighlights();
  const resolvedDomain = resolveDomainForHighlightId(all, normalizedId, domain || "");
  const before = getHighlightByDomainAndId(all, resolvedDomain, normalizedId);
  if (!before) {
    return { ok: false, reason: "not_found", removed: 0, historyId: null };
  }
  if (all[resolvedDomain]) {
    all[resolvedDomain] = all[resolvedDomain].filter((h) => normalizeHighlightId(h && h.id) !== normalizedId);
    if (all[resolvedDomain].length === 0) delete all[resolvedDomain];
    await setAllHighlights(all);
  }
  const favs = await getFavorites();
  const newFavs = favs.filter((fid) => normalizeHighlightId(fid) !== normalizedId);
  if (newFavs.length !== favs.length) {
    await setFavorites(newFavs);
  }

  let historyId = null;
  if (before) {
    const entry = buildHistoryEntry({
      action: "delete",
      domain: resolvedDomain,
      pageUrl: before.url || ""
    }, [{
      domain: resolvedDomain,
      id: normalizedId,
      before,
      after: null
    }]);
    await appendTimelineEntry(entry);
    historyId = entry ? entry.id : null;
    await appendDeletedAuditRows(buildDeletedAuditRowsFromEntry(entry));
  }

  if (!silent) {
    notifyTabsForDomain(resolvedDomain, {
      action: "removeHighlightFromPage",
      id: normalizedId
    });
  } else if (typeof excludedTabId === "number") {
    notifyTabsForDomainExcept(resolvedDomain, {
      action: "removeHighlightFromPage",
      id: normalizedId
    }, excludedTabId);
  }
  return { ok: true, reason: "", removed: 1, historyId };
}

async function updateHighlightColor(domain, id, color, bgColor) {
  if (!domain) return;
  const normalizedId = normalizeHighlightId(id);
  if (!normalizedId) return;
  const all = await getAllHighlights();
  if (all[domain]) {
    const h = all[domain].find((x) => normalizeHighlightId(x && x.id) === normalizedId);
    if (h) {
      const before = deepClone(h);
      h.color = color;
      h.bgColor = bgColor;
      await setAllHighlights(all);
      const entry = buildHistoryEntry({
        action: "update",
        domain: domain,
        pageUrl: h.url || ""
      }, [{
        domain: domain,
        id: normalizedId,
        before,
        after: deepClone(h)
      }]);
      await appendTimelineEntry(entry);
    }
  }
}

async function clearDomain(domain, sourcePageUrl = "") {
  const all = await getAllHighlights();
  const removed = all[domain] || [];
  const changes = removed.map((h) => ({
    domain,
    id: normalizeHighlightId(h && h.id),
    before: deepClone(h),
    after: null
  }));
  delete all[domain];
  await setAllHighlights(all);
  let historyId = null;
  if (removed.length > 0) {
    const removedIds = new Set(removed.map((h) => normalizeHighlightId(h && h.id)).filter(Boolean));
    const favs = await getFavorites();
    const newFavs = favs.filter((fid) => !removedIds.has(normalizeHighlightId(fid)));
    await setFavorites(newFavs);

    const entry = buildHistoryEntry({
      action: "clear_domain",
      domain,
      pageUrl: sourcePageUrl || (removed[0] && removed[0].url ? removed[0].url : "")
    }, changes);
    await appendTimelineEntry(entry);
    historyId = entry ? entry.id : null;
    await appendDeletedAuditRows(buildDeletedAuditRowsFromEntry(entry));
  }
  return { removed: removed.length, historyId };
}

async function clearRectangles(sourcePageUrl = "", sourceDomain = "") {
  const all = await getAllHighlights();
  const removedByDomain = {};
  const removedIds = new Set();
  const changes = [];
  let count = 0;

  for (const [domain, items] of Object.entries(all)) {
    if (!Array.isArray(items) || items.length === 0) continue;
    const keep = [];
    for (const item of items) {
      if (item && item.type === "shape-rect") {
        const itemId = normalizeHighlightId(item.id);
        if (!itemId) continue;
        count++;
        if (!removedByDomain[domain]) removedByDomain[domain] = [];
        removedByDomain[domain].push(itemId);
        removedIds.add(itemId);
        changes.push({
          domain,
          id: itemId,
          before: deepClone(item),
          after: null
        });
      } else {
        keep.push(item);
      }
    }
    if (keep.length > 0) all[domain] = keep;
    else delete all[domain];
  }

  if (count === 0) return { removed: 0 };

  await setAllHighlights(all);

  const favs = await getFavorites();
  const nextFavs = favs.filter((fid) => !removedIds.has(normalizeHighlightId(fid)));
  if (nextFavs.length !== favs.length) {
    await setFavorites(nextFavs);
  }

  const entry = buildHistoryEntry({
    action: "clear_rectangles",
    pageUrl: sourcePageUrl || "",
    domain: sourceDomain || ""
  }, changes);
  await appendTimelineEntry(entry);
  await appendDeletedAuditRows(buildDeletedAuditRowsFromEntry(entry));

  notifyTabsForRemovedIdsByDomain(removedByDomain);
  return { removed: count, historyId: entry ? entry.id : null };
}

async function bulkUpdateHighlights(updates, meta = {}) {
  if (!Array.isArray(updates) || updates.length === 0) return { updated: 0 };

  const all = await getAllHighlights();
  const changes = [];
  const changedDomains = new Set();
  let updated = 0;

  for (const update of updates) {
    if (!update || !update.domain) continue;
    const normalizedUpdateId = normalizeHighlightId(update.id);
    if (!normalizedUpdateId) continue;
    const items = all[update.domain];
    if (!Array.isArray(items)) continue;
    const idx = items.findIndex((h) => h && normalizeHighlightId(h.id) === normalizedUpdateId);
    if (idx < 0) continue;

    const current = items[idx];
    const next = { ...current };
    const before = deepClone(current);

    if (typeof update.color === "string") next.color = update.color;
    if (typeof update.bgColor === "string") next.bgColor = update.bgColor;
    if (typeof update.text === "string") next.text = update.text;
    if (typeof update.ocrText === "string") next.ocrText = update.ocrText;
    if (typeof update.label === "string") next.label = update.label;
    if (typeof update.revealed === "boolean") next.revealed = update.revealed;
    if (typeof update.transient === "boolean") next.transient = update.transient;
    if (update.rect && typeof update.rect === "object") {
      next.rect = {
        left: Number(update.rect.left) || 0,
        top: Number(update.rect.top) || 0,
        width: Math.max(0, Number(update.rect.width) || 0),
        height: Math.max(0, Number(update.rect.height) || 0)
      };
    }
    next.timestamp = Date.now();

    if (JSON.stringify(before) === JSON.stringify(next)) continue;

    items[idx] = next;
    updated++;
    changedDomains.add(update.domain);
    changes.push({
      domain: update.domain,
      id: normalizedUpdateId,
      before,
      after: deepClone(next)
    });
  }

  if (updated === 0) return { updated: 0 };

  await setAllHighlights(all);
  const entry = buildHistoryEntry({
    action: meta.action || "bulk_color",
    label: meta.label || null,
    pageUrl: meta.pageUrl || "",
    domain: meta.domain || ""
  }, changes);
  await appendTimelineEntry(entry);

  for (const domain of changedDomains) {
    notifyTabsForDomain(domain, { action: "highlightsUpdated" });
  }
  return { updated, historyId: entry ? entry.id : null };
}

async function bulkDeleteHighlights(targets, meta = {}) {
  if (!Array.isArray(targets) || targets.length === 0) return { removed: 0 };

  const all = await getAllHighlights();
  const byDomain = {};
  for (const target of targets) {
    if (!target) continue;
    const normalizedTargetId = normalizeHighlightId(target.id);
    if (!normalizedTargetId) continue;
    const resolvedDomain = resolveDomainForHighlightId(all, normalizedTargetId, target.domain || "");
    if (!resolvedDomain) continue;
    if (!byDomain[resolvedDomain]) byDomain[resolvedDomain] = new Set();
    byDomain[resolvedDomain].add(normalizedTargetId);
  }

  const changes = [];
  const changedDomains = new Set();
  const removedIds = new Set();
  let removed = 0;

  for (const [domain, idsSet] of Object.entries(byDomain)) {
    const items = all[domain];
    if (!Array.isArray(items) || items.length === 0) continue;
    const keep = [];
    for (const item of items) {
      const itemId = normalizeHighlightId(item && item.id);
      if (!item || !itemId || !idsSet.has(itemId)) {
        keep.push(item);
        continue;
      }
      removed++;
      removedIds.add(itemId);
      changedDomains.add(domain);
      changes.push({
        domain,
        id: itemId,
        before: deepClone(item),
        after: null
      });
    }
    if (keep.length > 0) all[domain] = keep;
    else delete all[domain];
  }

  if (removed === 0) return { removed: 0 };

  await setAllHighlights(all);

  const favs = await getFavorites();
  const nextFavs = favs.filter((fid) => !removedIds.has(normalizeHighlightId(fid)));
  if (nextFavs.length !== favs.length) {
    await setFavorites(nextFavs);
  }

  const entry = buildHistoryEntry({
    action: meta.action || "bulk_delete",
    label: meta.label || null,
    pageUrl: meta.pageUrl || "",
    domain: meta.domain || ""
  }, changes);
  await appendTimelineEntry(entry);
  await appendDeletedAuditRows(buildDeletedAuditRowsFromEntry(entry));

  for (const domain of changedDomains) {
    notifyTabsForDomain(domain, { action: "highlightsUpdated" });
  }

  return { removed, historyId: entry ? entry.id : null };
}

async function normalizeFavoritesAgainstHighlights() {
  const all = await getAllHighlights();
  const idSet = new Set();
  for (const domainItems of Object.values(all)) {
    if (!Array.isArray(domainItems)) continue;
    for (const item of domainItems) {
      if (item && item.id !== null && typeof item.id !== "undefined") {
        const normalizedId = normalizeHighlightId(item.id);
        if (normalizedId) idSet.add(normalizedId);
      }
    }
  }
  const favorites = await getFavorites();
  const nextFavorites = favorites.filter((id) => idSet.has(normalizeHighlightId(id)));
  if (nextFavorites.length !== favorites.length) {
    await setFavorites(nextFavorites);
  }
}

async function undoTimelineEntryByIndex(history, index) {
  try {
    if (!Array.isArray(history)) return { ok: false, reason: "history_missing" };
    if (index < 0 || index >= history.length) return { ok: false, reason: "not_found" };

    const entry = history[index];
    if (!entry || !Array.isArray(entry.changes) || entry.changes.length === 0) {
      return { ok: false, reason: "invalid_entry" };
    }

    const all = await getAllHighlights();
    const changedDomains = new Set();
    for (const change of entry.changes) {
      if (!change || !change.domain || !change.id) continue;
      setHighlightByDomainAndId(all, change.domain, change.id, change.before || null);
      changedDomains.add(change.domain);
    }

    await setAllHighlights(all);
    await normalizeFavoritesAgainstHighlights();

    history.splice(index, 1);
    await setTimelineHistory(history);

    for (const domain of changedDomains) {
      notifyTabsForDomain(domain, { action: "highlightsUpdated" });
    }

    return { ok: true, entry };
  } catch (err) {
    console.error("[HighlightMaster] undoTimelineEntryByIndex error:", err);
    return { ok: false, reason: "undo_error" };
  }
}

async function getTimelineForPage(pageUrl, domain, limit = 30) {
  const history = await getTimelineHistory();
  const cappedLimit = Math.max(1, Math.min(100, Number(limit) || 30));
  const filtered = history.filter((entry) => {
    if (!entry) return false;
    if (pageUrl) return entry.pageUrl === pageUrl;
    if (domain) {
      if (entry.domain === domain) return true;
      if (Array.isArray(entry.changedDomains) && entry.changedDomains.includes(domain)) return true;
      return false;
    }
    return true;
  });
  filtered.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));
  return filtered.slice(0, cappedLimit);
}

async function toggleFavorite(id) {
  const normalizedId = normalizeHighlightId(id);
  if (!normalizedId) return { ok: false, reason: "invalid_id", isFavorite: false };
  const favs = await getFavorites();
  const idx = favs.findIndex((favId) => normalizeHighlightId(favId) === normalizedId);
  if (idx >= 0) {
    favs.splice(idx, 1);
  } else {
    favs.push(normalizedId);
  }
  await setFavorites(favs);
  return { ok: true, isFavorite: idx < 0 };
}

// ---- Tab Helpers ----

function notifyTabsForDomain(domain, message) {
  chrome.tabs.query({}, (tabs) => {
    for (const tab of tabs) {
      if (!tab.url || !/^https?:/.test(tab.url)) continue;
      try {
        const tabDomain = new URL(tab.url).hostname;
        if (tabDomain === domain) {
          chrome.tabs.sendMessage(tab.id, message).catch(() => {});
        }
      } catch (_e) {}
    }
  });
}

function notifyTabsForDomainExcept(domain, message, excludedTabId) {
  chrome.tabs.query({}, (tabs) => {
    for (const tab of tabs) {
      if (typeof excludedTabId === "number" && tab.id === excludedTabId) continue;
      if (!tab.url || !/^https?:/.test(tab.url)) continue;
      try {
        const tabDomain = new URL(tab.url).hostname;
        if (tabDomain === domain) {
          chrome.tabs.sendMessage(tab.id, message).catch(() => {});
        }
      } catch (_e) {}
    }
  });
}

function notifyAllTabs(message) {
  chrome.tabs.query({}, (tabs) => {
    for (const tab of tabs) {
      if (!tab.url || !/^https?:/.test(tab.url)) continue;
      chrome.tabs.sendMessage(tab.id, message).catch(() => {});
    }
  });
}

function notifyTabsForRemovedIdsByDomain(removedByDomain) {
  chrome.tabs.query({}, (tabs) => {
    for (const tab of tabs) {
      if (!tab.url || !/^https?:/.test(tab.url)) continue;
      try {
        const tabDomain = new URL(tab.url).hostname;
        const ids = removedByDomain[tabDomain];
        if (!ids || !ids.length) continue;
        chrome.tabs.sendMessage(tab.id, {
          action: "removeHighlightsBatch",
          ids
        }).catch(() => {});
      } catch (_e) {}
    }
  });
}

async function jumpToHighlight(h) {
  const tabs = await chrome.tabs.query({});
  let targetTab = null;

  for (const tab of tabs) {
    if (tab.url === h.url) {
      targetTab = tab;
      break;
    }
  }

  if (!targetTab) {
    for (const tab of tabs) {
      try {
        if (tab.url && new URL(tab.url).hostname === h.domain) {
          targetTab = tab;
          break;
        }
      } catch (_e) {}
    }
  }

  if (targetTab) {
    await chrome.tabs.update(targetTab.id, { active: true });
    try {
      await chrome.windows.update(targetTab.windowId, { focused: true });
    } catch (_e) {}

    if (targetTab.url !== h.url) {
      await chrome.tabs.update(targetTab.id, { url: h.url });
      awaitTabLoadThenScroll(targetTab.id, h.id);
    } else {
      scrollToHighlightWithRetry(targetTab.id, h.id);
    }
  } else {
    const newTab = await chrome.tabs.create({ url: h.url });
    awaitTabLoadThenScroll(newTab.id, h.id);
  }
}

function awaitTabLoadThenScroll(tabId, highlightId) {
  const listener = (updatedTabId, info) => {
    if (updatedTabId === tabId && info.status === "complete") {
      chrome.tabs.onUpdated.removeListener(listener);
      scrollToHighlightWithRetry(tabId, highlightId);
    }
  };
  chrome.tabs.onUpdated.addListener(listener);
  setTimeout(() => chrome.tabs.onUpdated.removeListener(listener), 20000);
}

function scrollToHighlightWithRetry(tabId, highlightId, attempts = 0) {
  setTimeout(() => {
    chrome.tabs.sendMessage(tabId, {
      action: "scrollToHighlight",
      id: highlightId
    }).catch(() => {
      if (attempts < 10) {
        scrollToHighlightWithRetry(tabId, highlightId, attempts + 1);
      }
    });
  }, attempts === 0 ? 600 : 500);
}

// Serialize highlight/storage mutations to avoid lost updates caused by
// overlapping async writes (e.g. delete + immediate create).
let mutationQueue = Promise.resolve();

async function runMutationSerial(work) {
  const queued = mutationQueue.then(work, work);
  mutationQueue = queued.catch(() => {});
  return await queued;
}

// ---- Message Router ----

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (!msg || !msg.action) {
    sendResponse({ error: "no action" });
    return;
  }

  const handle = async () => {
    switch (msg.action) {
      case "getEnabled": {
        const enabled = await storageSyncGet("enabled");
        return { enabled: enabled !== false };
      }

      case "setEnabled": {
        await storageSyncSet({ enabled: msg.enabled });
        notifyAllTabs({ action: "enabledChanged", enabled: msg.enabled });
        return { ok: true };
      }

      case "getAutoOcrCopy": {
        const enabled = await storageSyncGet("autoOcrCopyEnabled");
        return { ok: true, enabled: enabled === true };
      }

      case "setAutoOcrCopy": {
        const nextEnabled = msg.enabled === true;
        await storageSyncSet({ autoOcrCopyEnabled: nextEnabled });
        notifyAllTabs({ action: "autoOcrCopyChanged", enabled: nextEnabled });
        return { ok: true, enabled: nextEnabled };
      }

      case "saveHighlight": {
        return await runMutationSerial(async () => {
          if (!isValidHighlightRecord(msg.highlight)) {
            return { ok: false, reason: "invalid_highlight" };
          }
          await saveHighlight(msg.highlight, msg.historyMeta || null);
          return { ok: true };
        });
      }

      case "getHighlightsForDomain": {
        const all = await getAllHighlights();
        return { highlights: all[msg.domain] || [] };
      }

      case "getAllHighlights": {
        const all = await getAllHighlights();
        return { highlights: all };
      }

      case "getFavorites": {
        const favs = await getFavorites();
        return { favorites: favs };
      }

      case "getDeletedAuditLog": {
        const rows = await getDeletedAuditLog();
        const limit = Math.max(1, Math.min(200, Number(msg.limit) || 50));
        const sortedRows = rows
          .slice()
          .sort((a, b) => sanitizeNumber(b.deletedAt, 0) - sanitizeNumber(a.deletedAt, 0))
          .slice(0, limit)
          .map((row) => ({
            ...row,
            rowKey: getDeletedAuditRowKey(row)
          }));
        return { ok: true, rows: sortedRows };
      }

      case "toggleFavorite": {
        const result = await toggleFavorite(msg.id);
        return result;
      }

      case "deleteHighlight": {
        return await runMutationSerial(async () => {
          const senderTabId = sender && sender.tab && typeof sender.tab.id === "number" ? sender.tab.id : null;
          const result = await deleteHighlight(msg.domain, msg.id, !!msg.silent, senderTabId);
          return {
            ok: !!(result && result.ok),
            reason: result && result.reason ? result.reason : "",
            removed: result && typeof result.removed === "number" ? result.removed : 0,
            historyId: result && result.historyId ? result.historyId : null
          };
        });
      }

      case "updateHighlightColor": {
        return await runMutationSerial(async () => {
          await updateHighlightColor(msg.domain, msg.id, msg.color, msg.bgColor);
          return { ok: true };
        });
      }

      case "bulkUpdateHighlights": {
        return await runMutationSerial(async () => {
          const result = await bulkUpdateHighlights(msg.updates || [], {
            action: msg.historyAction || "bulk_color",
            label: msg.historyLabel || null,
            pageUrl: msg.pageUrl || msg.url || "",
            domain: msg.domain || ""
          });
          return { ok: true, updated: result.updated || 0, historyId: result.historyId || null };
        });
      }

      case "clearDomain": {
        return await runMutationSerial(async () => {
          const result = await clearDomain(msg.domain, msg.pageUrl || msg.url || "");
          notifyTabsForDomain(msg.domain, {
            action: "clearHighlightsFromPage"
          });
          return { ok: true, removed: result && typeof result.removed === "number" ? result.removed : 0, historyId: result && result.historyId ? result.historyId : null };
        });
      }

      case "clearRectangles": {
        return await runMutationSerial(async () => {
          const result = await clearRectangles(msg.pageUrl || msg.url || "", msg.domain || "");
          return { ok: true, removed: result.removed || 0, historyId: result.historyId || null };
        });
      }

      case "bulkDeleteHighlights": {
        return await runMutationSerial(async () => {
          const result = await bulkDeleteHighlights(msg.targets || [], {
            action: msg.historyAction || "bulk_delete",
            label: msg.historyLabel || null,
            pageUrl: msg.pageUrl || msg.url || "",
            domain: msg.domain || ""
          });
          return { ok: true, removed: result.removed || 0, historyId: result.historyId || null };
        });
      }

      case "getTimelineForPage": {
        const entries = await getTimelineForPage(msg.pageUrl || msg.url || "", msg.domain || "", msg.limit || 30);
        return { ok: true, entries };
      }

      case "undoTimelineEntry": {
        return await runMutationSerial(async () => {
          const history = await getTimelineHistory();
          const index = history.findIndex((entry) => entry && entry.id === msg.entryId);
          if (index < 0) return { ok: false, reason: "not_found" };
          return await undoTimelineEntryByIndex(history, index);
        });
      }

      case "undoLastTimelineForPage": {
        return await runMutationSerial(async () => {
          try {
            const history = await getTimelineHistory();
            const pageUrl = msg.pageUrl || msg.url || "";
            const domain = msg.domain || "";
            let index = -1;
            for (let i = history.length - 1; i >= 0; i--) {
              const entry = history[i];
              if (!entry) continue;
              if (pageUrl) {
                if (entry.pageUrl === pageUrl) {
                  index = i;
                  break;
                }
                continue;
              }
              if (domain) {
                if (entry.domain === domain || (Array.isArray(entry.changedDomains) && entry.changedDomains.includes(domain))) {
                  index = i;
                  break;
                }
                continue;
              }
              index = i;
              break;
            }
            if (index < 0) return { ok: false, reason: "not_found" };
            return await undoTimelineEntryByIndex(history, index);
          } catch (err) {
            console.error("[HighlightMaster] undoLastTimelineForPage error:", err);
            return { ok: false, reason: "undo_error" };
          }
        });
      }

      case "jumpToHighlight": {
        await jumpToHighlight(msg.highlight);
        return { ok: true };
      }

      case "getCategories": {
        const cats = await getCategories();
        const assignments = await getDomainAssignments();
        const cutoffs = await getDomainAssignmentCutoffs();
        const cleanupResult = cleanupLegacyDefaultCategories(cats, assignments);
        let nextCategories = cleanupResult.categories;
        let nextAssignments = cleanupResult.domainAssignments;
        let nextCutoffs = cutoffs && typeof cutoffs === "object" ? { ...cutoffs } : {};
        let categoriesChanged = cleanupResult.categoriesChanged;
        let assignmentsChanged = cleanupResult.assignmentsChanged;
        let cutoffsChanged = false;
        const legacyPurgeDone = await storageLocalGet(LEGACY_DEFAULT_PURGE_KEY);
        let shouldMarkLegacyPurgeDone = false;
        if (legacyPurgeDone !== true) {
          const aggressiveCleanupResult = cleanupLegacyDefaultCategories(nextCategories, nextAssignments, { aggressiveNames: true });
          nextCategories = aggressiveCleanupResult.categories;
          nextAssignments = aggressiveCleanupResult.domainAssignments;
          categoriesChanged = categoriesChanged || aggressiveCleanupResult.categoriesChanged;
          assignmentsChanged = assignmentsChanged || aggressiveCleanupResult.assignmentsChanged;
          shouldMarkLegacyPurgeDone = true;
        }
        const now = Date.now();
        for (const domain of Object.keys(nextAssignments)) {
          const cutoffValue = Number(nextCutoffs[domain]);
          if (!Number.isFinite(cutoffValue) || cutoffValue <= 0) {
            nextCutoffs[domain] = now;
            cutoffsChanged = true;
          }
        }
        for (const domain of Object.keys(nextCutoffs)) {
          if (nextAssignments[domain]) continue;
          delete nextCutoffs[domain];
          cutoffsChanged = true;
        }
        if (categoriesChanged) {
          await setCategories(nextCategories);
        }
        if (assignmentsChanged) {
          await setDomainAssignments(nextAssignments);
        }
        if (cutoffsChanged) {
          await setDomainAssignmentCutoffs(nextCutoffs);
        }
        if (shouldMarkLegacyPurgeDone) {
          const payload = {};
          payload[LEGACY_DEFAULT_PURGE_KEY] = true;
          await storageLocalSet(payload);
        }
        return {
          categories: nextCategories,
          domainAssignments: nextAssignments,
          domainAssignmentCutoffs: nextCutoffs
        };
      }

      case "createCategory": {
        const cats = await getCategories();
        cats[msg.id] = { name: msg.name, parentId: msg.parentId || null };
        await setCategories(cats);
        return { ok: true };
      }

      case "updateCategory": {
        const cats = await getCategories();
        if (cats[msg.id]) {
          cats[msg.id].name = msg.name;
          await setCategories(cats);
        }
        return { ok: true };
      }

      case "deleteCategory": {
        const cats = await getCategories();
        delete cats[msg.id];
        for (const [k, v] of Object.entries(cats)) {
            if (v.parentId === msg.id) delete cats[k];
        }
        await setCategories(cats);
        const assignments = await getDomainAssignments();
        const cutoffs = await getDomainAssignmentCutoffs();
        let changed = false;
        let cutoffChanged = false;
        for (const [dom, cId] of Object.entries(assignments)) {
            if (!cats[cId]) {
              delete assignments[dom];
              changed = true;
              if (Object.prototype.hasOwnProperty.call(cutoffs, dom)) {
                delete cutoffs[dom];
                cutoffChanged = true;
              }
            }
        }
        if (changed) await setDomainAssignments(assignments);
        if (cutoffChanged) await setDomainAssignmentCutoffs(cutoffs);
        return { ok: true };
      }

      case "assignDomain": {
        const assignments = await getDomainAssignments();
        const cutoffs = await getDomainAssignmentCutoffs();
        if (msg.categoryId) {
          assignments[msg.domain] = msg.categoryId;
          cutoffs[msg.domain] = Date.now();
        } else {
          delete assignments[msg.domain];
          delete cutoffs[msg.domain];
        }
        await setDomainAssignments(assignments);
        await setDomainAssignmentCutoffs(cutoffs);
        return { ok: true };
      }

      case "exportCategory": {
         const cats = await getCategories();
         const targetCat = cats[msg.categoryId];
         if (!targetCat) return { ok: false };
         
         const exportCats = {};
         exportCats[msg.categoryId] = targetCat;
         for (const [k, v] of Object.entries(cats)) {
             if (v.parentId === msg.categoryId) exportCats[k] = v;
         }
         
         const assignments = await getDomainAssignments();
         const cutoffs = await getDomainAssignmentCutoffs();
         const exportAssignments = {};
         const exportCutoffs = {};
         for (const [dom, cId] of Object.entries(assignments)) {
             if (!exportCats[cId]) continue;
             exportAssignments[dom] = cId;
             if (Object.prototype.hasOwnProperty.call(cutoffs, dom)) {
               exportCutoffs[dom] = cutoffs[dom];
             }
         }
         
         const allHL = await getAllHighlights();
         const exportHighlights = {};
         for (const dom of Object.keys(exportAssignments)) {
             if (allHL[dom]) exportHighlights[dom] = allHL[dom];
         }
         
         return {
             ok: true,
             payload: {
                 version: 3,
                  exportedAt: new Date().toISOString(),
                  categories: exportCats,
                  domainAssignments: exportAssignments,
                  domainAssignmentCutoffs: exportCutoffs,
                  highlights: exportHighlights
              }
          };
      }

      case "exportAll": {
        const [highlights, favorites, enabled, autoOcrCopyEnabled, cats, assignments, assignmentCutoffs, timelineHistory, deletedAuditLog] = await Promise.all([
          getAllHighlights(),
          getFavorites(),
          storageSyncGet("enabled"),
          storageSyncGet("autoOcrCopyEnabled"),
          getCategories(),
          getDomainAssignments(),
          getDomainAssignmentCutoffs(),
          getTimelineHistory(),
          getDeletedAuditLog()
        ]);
        return {
          ok: true,
          payload: {
            version: 3,
            exportedAt: new Date().toISOString(),
            highlights,
            favorites,
            categories: cats,
            domainAssignments: assignments,
            domainAssignmentCutoffs: assignmentCutoffs,
            settings: { enabled, autoOcrCopyEnabled: autoOcrCopyEnabled === true },
            timelineHistory,
            deletedAuditLog
          }
        };
      }

      case "exportDeletedCsv": {
        try {
          const rows = await getDeletedAuditLog();
          if (!Array.isArray(rows)) {
            return { ok: false, reason: "invalid_data" };
          }
          const csv = buildDeletedAuditCsv(rows);
          if (typeof csv !== "string") {
            return { ok: false, reason: "csv_build_failed" };
          }
          return {
            ok: true,
            rows: rows.length,
            csv: csv,
            generatedAt: nowIso()
          };
        } catch (err) {
          console.error("[HighlightMaster] exportDeletedCsv error:", err);
          return { ok: false, reason: "export_error" };
        }
      }

      case "restoreDeletedAuditRows": {
        return await runMutationSerial(async () => {
          return await restoreDeletedAuditRows(msg.rowIds || []);
        });
      }

      case "importAll": {
        return await runMutationSerial(async () => {
          const payload = msg.payload;
          if (!payload || typeof payload !== "object") {
            return { ok: false, reason: "invalid" };
          }
          
          if (payload.categories) {
              const existingCats = await getCategories();
              for (const [k, v] of Object.entries(payload.categories)) {
                  let uniqueId = k;
                  while (existingCats[uniqueId] && existingCats[uniqueId].name !== v.name) {
                      uniqueId = uniqueId + "_" + Date.now();
                  }
                  existingCats[uniqueId] = v;
                  // Update parentIds if they were remapped? 
                  // Simplified: we trust the UUIDs are reasonably unique.
              }
              await setCategories(existingCats);
          }
          
          if (payload.domainAssignments || payload.domainAssignmentCutoffs) {
              const existingAssignments = await getDomainAssignments();
              const existingCutoffs = await getDomainAssignmentCutoffs();
              const now = Date.now();

              if (payload.domainAssignments) {
                for (const [dom, cId] of Object.entries(payload.domainAssignments)) {
                  existingAssignments[dom] = cId;
                  if (!Object.prototype.hasOwnProperty.call(existingCutoffs, dom)) {
                    existingCutoffs[dom] = now;
                  }
                }
              }

              if (payload.domainAssignmentCutoffs && typeof payload.domainAssignmentCutoffs === "object") {
                for (const [dom, cutoff] of Object.entries(payload.domainAssignmentCutoffs)) {
                  if (!existingAssignments[dom]) continue;
                  const nextCutoff = Number(cutoff);
                  if (Number.isFinite(nextCutoff) && nextCutoff > 0) {
                    existingCutoffs[dom] = nextCutoff;
                  }
                }
              }

              for (const dom of Object.keys(existingAssignments)) {
                const cutoffValue = Number(existingCutoffs[dom]);
                if (!Number.isFinite(cutoffValue) || cutoffValue <= 0) {
                  existingCutoffs[dom] = now;
                }
              }
              for (const dom of Object.keys(existingCutoffs)) {
                if (existingAssignments[dom]) continue;
                delete existingCutoffs[dom];
              }

              await setDomainAssignments(existingAssignments);
              await setDomainAssignmentCutoffs(existingCutoffs);
          }

          const existing = await getAllHighlights();
          const imported = payload.highlights || {};
          let count = 0;

          for (const [domain, items] of Object.entries(imported)) {
            if (!Array.isArray(items)) continue;
            if (!existing[domain]) existing[domain] = [];
            const existingIds = new Set(existing[domain].map((h) => h.id));
            for (const item of items) {
              const sanitizedItem = sanitizeHighlightRecord(item);
              if (sanitizedItem && sanitizedItem.id && !existingIds.has(sanitizedItem.id)) {
                existing[domain].push(sanitizedItem);
                existingIds.add(sanitizedItem.id);
                count++;
              }
            }
          }

          await setAllHighlights(existing);

          if (Array.isArray(payload.favorites)) {
            const existingFavs = await getFavorites();
            const merged = [...new Set([...existingFavs, ...payload.favorites])];
            await setFavorites(merged);
          }

          if (payload.settings && typeof payload.settings === "object") {
            const syncPayload = {};
            if (typeof payload.settings.enabled === "boolean") {
              syncPayload.enabled = payload.settings.enabled;
            }
            if (typeof payload.settings.autoOcrCopyEnabled === "boolean") {
              syncPayload.autoOcrCopyEnabled = payload.settings.autoOcrCopyEnabled;
            }
            if (Object.keys(syncPayload).length > 0) {
              await storageSyncSet(syncPayload);
            }
          }

          if (Array.isArray(payload.timelineHistory)) {
            const existingHistory = await getTimelineHistory();
            const mergedHistoryMap = new Map();
            for (const entry of [...existingHistory, ...payload.timelineHistory]) {
              if (!entry || !entry.id || !Array.isArray(entry.changes)) continue;
              mergedHistoryMap.set(entry.id, entry);
            }
            const mergedHistory = Array.from(mergedHistoryMap.values())
              .sort((a, b) => (a.timestamp || 0) - (b.timestamp || 0));
            if (mergedHistory.length > HISTORY_LIMIT) {
              mergedHistory.splice(0, mergedHistory.length - HISTORY_LIMIT);
            }
            await setTimelineHistory(mergedHistory);
          }

          if (Array.isArray(payload.deletedAuditLog)) {
            const existingDeleted = await getDeletedAuditLog();
            const mergedDeletedMap = new Map();
            for (const row of [...existingDeleted, ...payload.deletedAuditLog]) {
              if (!row || !row.highlightId || !row.deletedAtIso) continue;
              mergedDeletedMap.set(getDeletedAuditRowKey(row), row);
            }
            const mergedDeleted = Array.from(mergedDeletedMap.values())
              .sort((a, b) => sanitizeNumber(a.deletedAt, 0) - sanitizeNumber(b.deletedAt, 0));
            if (mergedDeleted.length > DELETED_AUDIT_LIMIT) {
              mergedDeleted.splice(0, mergedDeleted.length - DELETED_AUDIT_LIMIT);
            }
            await setDeletedAuditLog(mergedDeleted);
          }

          notifyAllTabs({ action: "highlightsUpdated" });
          if (payload.settings && typeof payload.settings === "object") {
            if (typeof payload.settings.enabled === "boolean") {
              notifyAllTabs({ action: "enabledChanged", enabled: payload.settings.enabled });
            }
            if (typeof payload.settings.autoOcrCopyEnabled === "boolean") {
              notifyAllTabs({ action: "autoOcrCopyChanged", enabled: payload.settings.autoOcrCopyEnabled });
            }
          }
          return { ok: true, imported: count };
        });
      }

      default:
        return { error: "unknown action" };
    }
  };

  handle()
    .then((result) => sendResponse(result))
    .catch((err) => sendResponse({ ok: false, error: err.message }));

  return true;
});
