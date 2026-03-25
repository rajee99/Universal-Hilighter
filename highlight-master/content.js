/*
 * HighlightMaster - Content Script (Universal Compatibility)
 * Features: hover toolbar, improved restore, color change, delete on hover
 * Fixed: re-highlighting bug, selection on existing highlights
 */

(() => {
  "use strict";

  if (window.__highlightMasterInjected) return;
  window.__highlightMasterInjected = true;

  const HIGHLIGHT_ATTR = "data-hm-id";
  const HIGHLIGHT_COLOR_ATTR = "data-hm-color";
  const HIGHLIGHT_KIND_ATTR = "data-hm-kind";
  const CONTEXT_CHARS = 100;
  const RECT_OCR_WORD_LIMIT = 18;
  const RECT_OCR_COPY_WORD_LIMIT = 600;
  const DOMAIN = location.hostname;

  const PRESET_COLORS = {
    yellow: { bg: "rgba(255, 212, 46, 0.38)", solid: "#ffd42e" },
    green:  { bg: "rgba(51, 195, 172, 0.34)", solid: "#33c3ac" },
    blue:   { bg: "rgba(79, 159, 232, 0.32)", solid: "#4f9fe8" },
    pink:   { bg: "rgba(255, 126, 182, 0.32)", solid: "#ff7eb6" },
    orange: { bg: "rgba(255, 159, 67, 0.34)", solid: "#ff9f43" }
  };
  const QUICK_SHORTCUT_COLORS = ["yellow", "green", "blue", "pink", "orange"];

  let extensionEnabled = true;
  let pickerHost = null;
  let pickerShadow = null;
  let pickerEl = null;
  let currentSelection = null;
  let pickerNoteEditorOpen = false;
  let pickerNoteDraft = "";
  let continuousHighlightEnabled = false;
  let continuousColorName = "yellow";
  let continuousBgColor = PRESET_COLORS.yellow.bg;
  let restoredIds = new Set();
  let storedHighlightsForPage = [];
  let drawMode = null;
  let drawStartPoint = null;
  let drawStartSourceNode = null;
  let drawPreviewEl = null;
  let drawMoved = false;
  let activeResize = null;
  let activeMove = null;
  let isDragging = false;
  let selectOverrideStyle = null;
  let colorPickerOpen = false;
  let autoOcrCopyEnabled = false;

  // Hover toolbar state
  let hoverHost = null;
  let hoverShadow = null;
  let hoverEl = null;
  let hoverTimeout = null;
  let activeHoverId = null;
  let activeHoverTarget = null;
  let hoverColorPickerOpen = false;
  let hoverNoteEditorId = "";
  let hoverNoteDraft = "";
  let shapePositionSyncRaf = null;
  let inlineUndoEl = null;
  let inlineUndoTimer = null;
  let textHighlightContextMenuEl = null;
  let textHighlightContextMenuId = "";
  
  let lastUsedColor = "yellow";
  let lastUsedBgColor = PRESET_COLORS["yellow"].bg;
  const RECT_MIN_SIZE = 18;
  const HOVER_TOOLBAR_SHOW_DELAY_MS = 180;
  const HOVER_TOOLBAR_HIDE_DELAY_MS = 480;

  function normalizeHighlightIdValue(id) {
    if (id === null || typeof id === "undefined") return "";
    return String(id);
  }

  function idsMatch(a, b) {
    const left = normalizeHighlightIdValue(a);
    const right = normalizeHighlightIdValue(b);
    return !!left && left === right;
  }

  function isScrollableContainer(node) {
    if (!node || node.nodeType !== 1) return false;
    const el = node;
    const style = window.getComputedStyle(el);
    if (!style) return false;
    const overflowY = style.overflowY || "";
    const overflowX = style.overflowX || "";
    const canScrollY = (overflowY === "auto" || overflowY === "scroll" || overflowY === "overlay") && el.scrollHeight > (el.clientHeight + 4);
    const canScrollX = (overflowX === "auto" || overflowX === "scroll" || overflowX === "overlay") && el.scrollWidth > (el.clientWidth + 4);
    return canScrollY || canScrollX;
  }

  function getNearestScrollContainer(startNode) {
    let node = startNode;
    if (node && node.nodeType === 3) node = node.parentElement;
    while (node && node !== document.body && node !== document.documentElement) {
      if (isScrollableContainer(node)) return node;
      node = node.parentElement;
    }
    return null;
  }

  function buildShapeScrollAnchor(sourceNode) {
    const host = getNearestScrollContainer(sourceNode);
    if (!host) return null;
    return {
      scrollHostXPath: getXPath(host),
      scrollHostTop: Number(host.scrollTop) || 0,
      scrollHostLeft: Number(host.scrollLeft) || 0
    };
  }

  function resolveShapeScrollHost(record) {
    if (!record || typeof record.scrollHostXPath !== "string" || !record.scrollHostXPath) return null;
    const node = resolveXPath(record.scrollHostXPath);
    return (node && node.nodeType === 1) ? node : null;
  }

  function getShapeScrollDelta(record) {
    const host = resolveShapeScrollHost(record);
    if (!host) return { x: 0, y: 0 };
    const baseTop = Number(record.scrollHostTop) || 0;
    const baseLeft = Number(record.scrollHostLeft) || 0;
    const currentTop = Number(host.scrollTop) || 0;
    const currentLeft = Number(host.scrollLeft) || 0;
    return {
      x: baseLeft - currentLeft,
      y: baseTop - currentTop
    };
  }

  function getRenderedRectForRecord(record) {
    if (!record || !record.rect) return null;
    const baseRect = record.rect;
    const delta = getShapeScrollDelta(record);
    return {
      left: (Number(baseRect.left) || 0) + delta.x,
      top: (Number(baseRect.top) || 0) + delta.y,
      width: Math.max(1, Number(baseRect.width) || 0),
      height: Math.max(1, Number(baseRect.height) || 0)
    };
  }

  function inferShapeScrollAnchorIfMissing(record) {
    if (!record || !record.rect) return;
    if (record.scrollHostXPath) return;
    const baseRect = record.rect;
    const centerClientX = (Number(baseRect.left) || 0) - window.scrollX + ((Number(baseRect.width) || 0) / 2);
    const centerClientY = (Number(baseRect.top) || 0) - window.scrollY + ((Number(baseRect.height) || 0) / 2);
    const clampedX = Math.max(0, Math.min(window.innerWidth - 1, centerClientX));
    const clampedY = Math.max(0, Math.min(window.innerHeight - 1, centerClientY));
    const target = document.elementFromPoint(clampedX, clampedY);
    const anchor = buildShapeScrollAnchor(target);
    if (!anchor) return;
    record.scrollHostXPath = anchor.scrollHostXPath;
    record.scrollHostTop = anchor.scrollHostTop;
    record.scrollHostLeft = anchor.scrollHostLeft;
  }

  function syncShapePositionsFromStorage() {
    if (activeResize || activeMove || drawStartPoint) return;
    for (const record of storedHighlightsForPage) {
      if (!record || !record.rect) continue;
      if (typeof record.type !== "string" || record.type.indexOf("shape-") !== 0) continue;
      const renderedRect = getRenderedRectForRecord(record);
      if (!renderedRect) continue;
      const elements = getAnnotationElements(record.id);
      for (const el of elements) {
        if (!el || el.nodeType !== 1) continue;
        applyRectToShapeElement(el, renderedRect);
      }
    }
  }

  function scheduleShapePositionSync() {
    if (shapePositionSyncRaf) return;
    shapePositionSyncRaf = window.requestAnimationFrame(() => {
      shapePositionSyncRaf = null;
      syncShapePositionsFromStorage();
    });
  }

  function sendRuntimeMessageSafe(message, callback) {
    const done = typeof callback === "function" ? callback : null;
    if (!chrome || !chrome.runtime || typeof chrome.runtime.sendMessage !== "function" || !chrome.runtime.id) {
      if (done) done(null, "runtime_unavailable");
      return false;
    }
    try {
      chrome.runtime.sendMessage(message, (response) => {
        let errorMessage = "";
        try {
          if (chrome.runtime && chrome.runtime.lastError && chrome.runtime.lastError.message) {
            errorMessage = chrome.runtime.lastError.message;
          }
        } catch (_err) {}
        if (done) done(response, errorMessage);
      });
      return true;
    } catch (err) {
      if (done) {
        const messageText = err && err.message ? err.message : String(err || "runtime_error");
        done(null, messageText);
      }
      return false;
    }
  }

  function setContinuousHighlightColor(colorName, bgColor) {
    continuousColorName = colorName || "yellow";
    continuousBgColor = bgColor || PRESET_COLORS.yellow.bg;
  }

  function normalizeContinuousColorPayload(colorName, bgColor) {
    let nextColorName = (typeof colorName === "string" && colorName.trim())
      ? colorName.trim()
      : (continuousColorName || "yellow");
    let nextBgColor = (typeof bgColor === "string" && bgColor.trim())
      ? bgColor.trim()
      : "";

    if (nextColorName.indexOf("custom:") === 0) {
      const customHex = nextColorName.slice(7).trim();
      if (!nextBgColor && /^#[0-9a-fA-F]{6}$/.test(customHex)) {
        nextBgColor = customHex;
      }
    } else if (PRESET_COLORS[nextColorName]) {
      if (!nextBgColor) nextBgColor = PRESET_COLORS[nextColorName].bg;
    } else {
      nextColorName = "yellow";
      if (!nextBgColor) nextBgColor = PRESET_COLORS.yellow.bg;
    }

    if (!nextBgColor) nextBgColor = PRESET_COLORS.yellow.bg;
    return { colorName: nextColorName, bgColor: nextBgColor };
  }

  function getContinuousModeState() {
    return {
      enabled: !!continuousHighlightEnabled,
      colorName: continuousColorName || "yellow",
      bgColor: continuousBgColor || PRESET_COLORS.yellow.bg
    };
  }

  function applyNoteTooltip(el, note) {
    if (!el) return;
    const text = String(note || "").trim();
    if (text) el.setAttribute("title", text);
    else el.removeAttribute("title");
  }

  // ---- Selective user-select Override ----

  function enableSelectOverride() {
    if (selectOverrideStyle) return;
    selectOverrideStyle = document.createElement("style");
    selectOverrideStyle.id = "hm-select-override";
    selectOverrideStyle.textContent = "*, *::before, *::after { -webkit-user-select: text !important; -moz-user-select: text !important; user-select: text !important; }";
    (document.head || document.documentElement).appendChild(selectOverrideStyle);
  }

  function disableSelectOverride() {
    if (selectOverrideStyle && selectOverrideStyle.parentNode) {
      selectOverrideStyle.parentNode.removeChild(selectOverrideStyle);
    }
    selectOverrideStyle = null;
  }

  document.addEventListener("mousedown", (e) => {
    if (drawMode) return;
    if (e.target && e.target.closest && e.target.closest(".hm-resize-handle")) return;
    const tag = (e.target.tagName || "").toLowerCase();
    const interactiveSelector = 'button, a, input, select, textarea, [role="button"], [role="slider"], [role="checkbox"], [role="radio"], [draggable="true"], label';
    const isInteractive = e.target.closest(interactiveSelector);
    if (isInteractive || ["button","a","input","select","textarea"].includes(tag)) {
      return;
    }
    isDragging = true;
    enableSelectOverride();
  }, true);

  document.addEventListener("mouseup", () => {
    setTimeout(() => {
      isDragging = false;
      disableSelectOverride();
    }, 200);
  }, true);

  // ---- Inject Highlight Styles ----

  let highlightStyleEl = null;

  function injectHighlightStyles() {
    if (highlightStyleEl && highlightStyleEl.parentNode) return;
    highlightStyleEl = document.createElement("style");
    highlightStyleEl.id = "hm-highlight-styles";
    highlightStyleEl.textContent = [
      "span[" + HIGHLIGHT_ATTR + "] {",
      "  border-radius: 2px !important;",
      "  padding: 0px !important;",
      "  box-decoration-break: clone !important;",
      "  -webkit-box-decoration-break: clone !important;",
      "  display: inline !important;",
      "  visibility: visible !important;",
      "  opacity: 1 !important;",
      "  transition: filter 0.2s ease, box-shadow 0.2s ease !important;",
      "  cursor: pointer !important;",
      "}",
      "span[" + HIGHLIGHT_ATTR + "].hm-hovered {",
      "  filter: brightness(0.88) !important;",
      "  box-shadow: 0 1px 6px rgba(0,0,0,0.2) !important;",
      "}",
      "div[" + HIGHLIGHT_ATTR + "][data-hm-kind^='shape'] {",
      "  position: absolute !important;",
      "  pointer-events: auto !important;",
      "  box-sizing: border-box !important;",
      "  cursor: pointer !important;",
      "  z-index: 2147483645 !important;",
      "  transition: transform 0.18s ease, box-shadow 0.18s ease, filter 0.18s ease !important;",
      "}",
      "div[" + HIGHLIGHT_ATTR + "][data-hm-kind='shape-rect'] {",
      "  border: 2px solid currentColor !important;",
      "  border-radius: 6px !important;",
      "  backdrop-filter: saturate(1.05) !important;",
      "}",
      "div[" + HIGHLIGHT_ATTR + "][data-hm-kind='shape-rect'][data-hm-revealed='1'] {",
      "  border-style: dashed !important;",
      "}",
      "div[" + HIGHLIGHT_ATTR + "][data-hm-kind='shape-cover'] {",
      "  border: 2px solid rgba(0,0,0,0.16) !important;",
      "  border-radius: 6px !important;",
      "}",
      "div[" + HIGHLIGHT_ATTR + "][data-hm-kind^='shape'].hm-hovered {",
      "  transform: scale(1.01) !important;",
      "  filter: brightness(0.94) !important;",
      "  box-shadow: 0 10px 18px rgba(0,0,0,0.22) !important;",
      "}",
      "div[" + HIGHLIGHT_ATTR + "][data-hm-kind='shape-rect'] .hm-resize-handle {",
      "  position: absolute !important;",
      "  width: 12px !important;",
      "  height: 12px !important;",
      "  border-radius: 999px !important;",
      "  background: linear-gradient(180deg, rgba(255,255,255,0.98), rgba(226,237,255,0.96)) !important;",
      "  border: 1.5px solid rgba(14, 49, 87, 0.88) !important;",
      "  box-shadow: 0 2px 10px rgba(0,0,0,0.26) !important;",
      "  opacity: 0 !important;",
      "  transform: scale(0.88) !important;",
      "  transition: opacity 0.16s ease, transform 0.16s ease !important;",
      "  pointer-events: auto !important;",
      "  z-index: 2 !important;",
      "}",
      "div[" + HIGHLIGHT_ATTR + "][data-hm-kind='shape-rect'] .hm-resize-handle[data-hm-corner='nw'] {",
      "  top: -8px !important;",
      "  left: -8px !important;",
      "  cursor: nwse-resize !important;",
      "}",
      "div[" + HIGHLIGHT_ATTR + "][data-hm-kind='shape-rect'] .hm-resize-handle[data-hm-corner='ne'] {",
      "  top: -8px !important;",
      "  right: -8px !important;",
      "  cursor: nesw-resize !important;",
      "}",
      "div[" + HIGHLIGHT_ATTR + "][data-hm-kind='shape-rect'] .hm-resize-handle[data-hm-corner='sw'] {",
      "  bottom: -8px !important;",
      "  left: -8px !important;",
      "  cursor: nesw-resize !important;",
      "}",
      "div[" + HIGHLIGHT_ATTR + "][data-hm-kind='shape-rect'] .hm-resize-handle[data-hm-corner='se'] {",
      "  bottom: -8px !important;",
      "  right: -8px !important;",
      "  cursor: nwse-resize !important;",
      "}",
      "div[" + HIGHLIGHT_ATTR + "][data-hm-kind='shape-rect'].hm-hovered .hm-resize-handle,",
      "div[" + HIGHLIGHT_ATTR + "][data-hm-kind='shape-rect'].hm-resizing .hm-resize-handle {",
      "  opacity: 1 !important;",
      "  transform: scale(1) !important;",
      "}",
      ".hm-ocr-copy-chip {",
      "  position: absolute !important;",
      "  top: -8px !important;",
      "  left: calc(100% + 8px) !important;",
      "  max-width: 220px !important;",
      "  height: 22px !important;",
      "  padding: 0 8px !important;",
      "  border-radius: 999px !important;",
      "  border: 1px solid rgba(79, 159, 232, 0.5) !important;",
      "  background: linear-gradient(180deg, rgba(16, 40, 67, 0.95), rgba(8, 22, 39, 0.95)) !important;",
      "  color: #d9ecff !important;",
      "  font: 700 10px/1 -apple-system, BlinkMacSystemFont, 'SF Pro Display', 'SF Pro Text', system-ui, sans-serif !important;",
      "  letter-spacing: 0.02em !important;",
      "  display: inline-flex !important;",
      "  align-items: center !important;",
      "  gap: 5px !important;",
      "  white-space: nowrap !important;",
      "  overflow: hidden !important;",
      "  text-overflow: ellipsis !important;",
      "  cursor: pointer !important;",
      "  pointer-events: auto !important;",
      "  box-shadow: 0 6px 14px rgba(0,0,0,0.24) !important;",
      "  z-index: 3 !important;",
      "}",
      ".hm-ocr-copy-chip:hover {",
      "  border-color: rgba(79, 159, 232, 0.78) !important;",
      "  color: #f3f9ff !important;",
      "  background: linear-gradient(180deg, rgba(22, 55, 88, 0.96), rgba(10, 27, 47, 0.96)) !important;",
      "}",
      ".hm-ocr-copy-chip .hm-ocr-copy-prefix {",
      "  color: rgba(208, 231, 255, 0.9) !important;",
      "  flex: 0 0 auto !important;",
      "}",
      ".hm-ocr-copy-chip .hm-ocr-copy-snippet {",
      "  color: rgba(185, 214, 246, 0.95) !important;",
      "  overflow: hidden !important;",
      "  text-overflow: ellipsis !important;",
      "}",
      ".hm-highlight-context-menu {",
      "  position: absolute !important;",
      "  z-index: 2147483647 !important;",
      "  min-width: 170px !important;",
      "  max-width: 240px !important;",
      "  padding: 6px !important;",
      "  border-radius: 10px !important;",
      "  border: 1px solid rgba(66, 123, 181, 0.45) !important;",
      "  background: linear-gradient(180deg, rgba(11, 25, 41, 0.98), rgba(7, 18, 31, 0.98)) !important;",
      "  box-shadow: 0 14px 28px rgba(0,0,0,0.32) !important;",
      "}",
      ".hm-highlight-context-menu button {",
      "  appearance: none !important;",
      "  width: 100% !important;",
      "  border: 1px solid rgba(95, 163, 229, 0.35) !important;",
      "  border-radius: 8px !important;",
      "  background: linear-gradient(180deg, rgba(30, 66, 102, 0.94), rgba(18, 44, 71, 0.94)) !important;",
      "  color: #eaf4ff !important;",
      "  font: 700 12px -apple-system, BlinkMacSystemFont, 'SF Pro Display', 'SF Pro Text', system-ui, sans-serif !important;",
      "  line-height: 1.1 !important;",
      "  text-align: left !important;",
      "  padding: 8px 10px !important;",
      "  cursor: pointer !important;",
      "}",
      ".hm-highlight-context-menu button:hover {",
      "  border-color: rgba(117, 190, 255, 0.68) !important;",
      "  background: linear-gradient(180deg, rgba(37, 79, 120, 0.96), rgba(22, 55, 86, 0.96)) !important;",
      "}",
      ".hm-inline-undo {",
      "  position: fixed !important;",
      "  left: 50% !important;",
      "  bottom: 24px !important;",
      "  transform: translateX(-50%) !important;",
      "  z-index: 2147483647 !important;",
      "  display: inline-flex !important;",
      "  align-items: center !important;",
      "  gap: 10px !important;",
      "  padding: 10px 12px !important;",
      "  border-radius: 999px !important;",
      "  border: 1px solid rgba(255,255,255,0.22) !important;",
      "  background: linear-gradient(180deg, rgba(21, 28, 41, 0.96), rgba(11, 16, 26, 0.94)) !important;",
      "  color: #eaf4ff !important;",
      "  font: 600 12px -apple-system, BlinkMacSystemFont, 'SF Pro Display', 'SF Pro Text', system-ui, sans-serif !important;",
      "  box-shadow: 0 10px 28px rgba(0,0,0,0.35) !important;",
      "}",
      ".hm-inline-undo button {",
      "  border: 1px solid rgba(255,255,255,0.28) !important;",
      "  border-radius: 999px !important;",
      "  padding: 4px 10px !important;",
      "  background: rgba(255,255,255,0.12) !important;",
      "  color: #ffffff !important;",
      "  cursor: pointer !important;",
      "  font: 700 11px -apple-system, BlinkMacSystemFont, 'SF Pro Display', 'SF Pro Text', system-ui, sans-serif !important;",
      "}",
      "@keyframes hm-pop-anim {",
      "  0%   { opacity: 0; filter: brightness(1.5) saturate(1.5); outline: 3px solid rgba(255,255,255,0.6); outline-offset: -1px; }",
      "  100% { opacity: 1; filter: brightness(1) saturate(1); outline: 0px solid transparent; outline-offset: 4px; }",
      "}",
      ".hm-pop-anim {",
      "  animation: hm-pop-anim 0.4s cubic-bezier(0.16, 1, 0.3, 1) !important;",
      "}",
      "@keyframes hm-pulse {",
      "  0%   { outline: 3px solid rgba(79,159,232,0.88); outline-offset: 2px; }",
      "  50%  { outline: 3px solid rgba(79,159,232,0.2); outline-offset: 5px; }",
      "  100% { outline: 3px solid rgba(79,159,232,0.88); outline-offset: 2px; }",
      "}",
      ".hm-pulse {",
      "  animation: hm-pulse 0.5s ease-in-out 3 !important;",
      "}",
      "#hm-picker-host, #hm-hover-host {",
      "  position: absolute !important;",
      "  z-index: 2147483647 !important;",
      "  pointer-events: none !important;",
      "  top: 0 !important; left: 0 !important;",
      "  width: 0 !important; height: 0 !important;",
      "}",
      ".hm-draw-preview {",
      "  position: absolute !important;",
      "  z-index: 2147483646 !important;",
      "  pointer-events: none !important;",
      "  box-sizing: border-box !important;",
      "}",
      ".hm-draw-cursor, .hm-draw-cursor * {",
      "  cursor: crosshair !important;",
      "}",
      ".hm-draw-cursor .hm-ocr-copy-chip, .hm-draw-cursor .hm-ocr-copy-chip * {",
      "  cursor: pointer !important;",
      "}",
      ".hm-rect-resize-nwse, .hm-rect-resize-nwse * {",
      "  cursor: nwse-resize !important;",
      "}",
      ".hm-rect-resize-nesw, .hm-rect-resize-nesw * {",
      "  cursor: nesw-resize !important;",
      "}",
      ".hm-rect-move, .hm-rect-move * {",
      "  cursor: move !important;",
      "}",
      "div[" + HIGHLIGHT_ATTR + "][data-hm-kind='shape-rect'].hm-moving {",
      "  cursor: move !important;",
      "  transform: scale(1.01) !important;",
      "}"
    ].join("\n");
    (document.head || document.documentElement).appendChild(highlightStyleEl);
  }

  injectHighlightStyles();

  const styleGuard = new MutationObserver(() => {
    if (!document.getElementById("hm-highlight-styles")) {
      injectHighlightStyles();
    }
  });
  if (document.head) {
    styleGuard.observe(document.head, { childList: true });
  } else {
    document.addEventListener("DOMContentLoaded", () => {
      styleGuard.observe(document.head || document.documentElement, { childList: true });
    });
  }

  // ---- Shadow DOM Picker Styles (iOS-inspired) ----

  function getPickerStyles() {
    return [
      ".hm-picker {",
      "  position: fixed;",
      "  display: flex;",
      "  flex-direction: column;",
      "  align-items: stretch;",
      "  gap: 8px;",
      "  padding: 10px 12px 12px;",
      "  background: linear-gradient(180deg, rgba(22, 27, 37, 0.94), rgba(12, 16, 24, 0.9));",
      "  border: 1px solid rgba(255,255,255,0.1);",
      "  border-radius: 18px;",
      "  box-shadow: 0 18px 42px rgba(0,0,0,0.28), inset 0 1px 0 rgba(255,255,255,0.08);",
      "  backdrop-filter: blur(16px) saturate(1.1);",
      "  opacity: 0;",
      "  transform: translateY(6px) scale(0.92);",
      "  transition: opacity 0.22s cubic-bezier(0.32,0.72,0,1), transform 0.22s cubic-bezier(0.32,0.72,0,1);",
      "  pointer-events: auto;",
      "  user-select: none;",
      "  z-index: 2147483647;",
      "  font-family: -apple-system, BlinkMacSystemFont, 'SF Pro Display', 'SF Pro Text', system-ui, sans-serif;",
      "}",
      ".hm-controls-row {",
      "  width: 100%;",
      "  display: flex;",
      "  flex-wrap: wrap;",
      "  align-items: center;",
      "  justify-content: center;",
      "  gap: 7px;",
      "  min-height: 28px;",
      "}",
      "@keyframes hm-picker-float {",
      "  0%, 100% { transform: translateY(0) scale(1); }",
      "  50% { transform: translateY(-6px) scale(1.01); }",
      "}",
      ".hm-picker.visible {",
      "  opacity: 1;",
      "  transform: translateY(0) scale(1);",
      "  animation: hm-picker-float 3s ease-in-out 0.25s infinite;",
      "}",
      ".hm-color-btn {",
      "  width: 26px;",
      "  height: 26px;",
      "  border-radius: 50%;",
      "  cursor: pointer;",
      "  border: 2px solid rgba(255,255,255,0.22);",
      "  transition: transform 0.2s cubic-bezier(0.32,0.72,0,1), border-color 0.2s ease, box-shadow 0.2s ease, filter 0.2s ease;",
      "  flex-shrink: 0;",
      "  box-sizing: border-box;",
      "  box-shadow: inset 0 1px 0 rgba(255,255,255,0.25);",
      "}",
      ".hm-color-btn:hover {",
      "  transform: translateY(-1px) scale(1.16);",
      "  border-color: rgba(255,255,255,0.68);",
      "  filter: saturate(1.06) brightness(1.04);",
      "}",
      ".hm-color-btn:active {",
      "  transform: scale(0.92);",
      "}",
      ".hm-divider {",
      "  width: 1px;",
      "  height: 22px;",
      "  background: rgba(255,255,255,0.1);",
      "  margin: 0 1px;",
      "  border-radius: 1px;",
      "}",
      ".hm-custom-wrap {",
      "  position: relative;",
      "  width: 26px;",
      "  height: 26px;",
      "  flex-shrink: 0;",
      "}",
      ".hm-custom-btn {",
      "  width: 26px;",
      "  height: 26px;",
      "  border-radius: 50%;",
      "  cursor: pointer;",
      "  border: 2px solid rgba(255,255,255,0.22);",
      "  background: conic-gradient(from 0deg, #ffd42e, #ff9f43, #ff7eb6, #4f9fe8, #33c3ac, #ffd42e);",
      "  transition: transform 0.2s cubic-bezier(0.32,0.72,0,1), border-color 0.2s ease, box-shadow 0.2s ease;",
      "  box-sizing: border-box;",
      "  box-shadow: inset 0 1px 0 rgba(255,255,255,0.3);",
      "}",
      ".hm-custom-btn:hover {",
      "  transform: translateY(-1px) scale(1.16);",
      "  border-color: rgba(255,255,255,0.68);",
      "  box-shadow: 0 0 0 4px rgba(255,255,255,0.08);",
      "}",
      ".hm-color-input {",
      "  position: absolute;",
      "  top: 0; left: 0;",
      "  width: 26px; height: 26px;",
      "  opacity: 0;",
      "  cursor: pointer;",
      "  border: none;",
      "  padding: 0;",
      "}",
      ".hm-delete-btn {",
      "  width: 26px;",
      "  height: 26px;",
      "  border-radius: 50%;",
      "  cursor: pointer;",
      "  border: 2px solid rgba(255,255,255,0.14);",
      "  background: rgba(255, 69, 58, 0.25);",
      "  display: flex;",
      "  align-items: center;",
      "  justify-content: center;",
      "  transition: transform 0.2s cubic-bezier(0.32,0.72,0,1), background 0.2s ease, border-color 0.2s ease;",
      "  flex-shrink: 0;",
      "  box-sizing: border-box;",
      "}",
      ".hm-delete-btn:hover {",
      "  transform: scale(1.25);",
      "  background: rgba(255, 69, 58, 0.5);",
      "  border-color: rgba(255, 69, 58, 0.7);",
      "}",
      ".hm-delete-btn:active {",
      "  transform: scale(0.92);",
      "}",
      ".hm-delete-btn svg {",
      "  width: 12px;",
      "  height: 12px;",
      "  stroke: #fff;",
      "  stroke-width: 2.5;",
      "  fill: none;",
      "}",
      ".hm-preview {",
      "  width: 100%;",
      "  margin-top: 0;",
      "  padding-top: 10px;",
      "  border-top: 1px solid rgba(255,255,255,0.08);",
      "  display: flex;",
      "  justify-content: center;",
      "}",
      ".hm-preview-text {",
      "  padding: 4px 10px;",
      "  border-radius: 9px;",
      "  font-size: 12px;",
      "  line-height: 1.35;",
      "  font-weight: 600;",
      "  letter-spacing: 0.01em;",
      "  box-shadow: inset 0 -1px 0 rgba(0,0,0,0.08), 0 1px 8px rgba(0,0,0,0.12);",
      "}",
      ".hm-mode-btn {",
      "  min-width: 38px;",
      "  height: 28px;",
      "  border-radius: 999px;",
      "  padding: 0 10px;",
      "  display: inline-flex;",
      "  align-items: center;",
      "  justify-content: center;",
      "  border: 1px solid rgba(255,255,255,0.14);",
      "  color: rgba(255,255,255,0.82);",
      "  background: rgba(255,255,255,0.08);",
      "  cursor: pointer;",
      "  font-size: 11px;",
      "  font-weight: 700;",
      "  letter-spacing: 0.03em;",
      "}",
      ".hm-mode-btn.active {",
      "  background: rgba(255,255,255,0.2);",
      "  border-color: rgba(255,255,255,0.32);",
      "  color: #ffffff;",
      "}",
      ".hm-mode-btn.hm-continuous-btn {",
      "  min-width: 44px;",
      "  height: 32px;",
      "  border-radius: 12px;",
      "  font-size: 12px;",
      "  font-weight: 800;",
      "  letter-spacing: 0.02em;",
      "}",
      ".hm-mode-btn.hm-continuous-btn.active {",
      "  border-color: rgba(255, 212, 46, 0.62);",
      "  background: rgba(255, 212, 46, 0.24);",
      "  color: #fff5cc;",
      "}",
      ".hm-state-btn {",
      "  min-width: 30px;",
      "  height: 26px;",
      "  border-radius: 999px;",
      "  padding: 0 8px;",
      "  display: inline-flex;",
      "  align-items: center;",
      "  justify-content: center;",
      "  border: 1px solid rgba(255,255,255,0.14);",
      "  color: rgba(255,255,255,0.82);",
      "  background: rgba(255,255,255,0.08);",
      "  cursor: pointer;",
      "  font-size: 10px;",
      "  font-weight: 700;",
      "  letter-spacing: 0.03em;",
      "  line-height: 1;",
      "  transition: transform 0.18s cubic-bezier(0.32,0.72,0,1), border-color 0.2s ease, background 0.2s ease, color 0.2s ease;",
      "  box-sizing: border-box;",
      "}",
      ".hm-state-btn:hover {",
      "  transform: translateY(-1px) scale(1.06);",
      "  border-color: rgba(255,255,255,0.3);",
      "  background: rgba(255,255,255,0.16);",
      "}",
      ".hm-state-btn:active {",
      "  transform: scale(0.94);",
      "}",
      ".hm-state-btn.active {",
      "  background: rgba(255,255,255,0.24);",
      "  border-color: rgba(255,255,255,0.4);",
      "  color: #ffffff;",
      "}",
      ".hm-note-preview {",
      "  margin-top: 8px;",
      "  max-width: 420px;",
      "  padding: 6px 10px;",
      "  border-radius: 10px;",
      "  border: 1px solid rgba(255,255,255,0.12);",
      "  background: rgba(255,255,255,0.06);",
      "  color: rgba(234, 244, 255, 0.9);",
      "  font-size: 11px;",
      "  line-height: 1.35;",
      "  white-space: normal;",
      "  word-break: break-word;",
      "}",
      ".hm-note-editor {",
      "  margin-top: 8px;",
      "  display: flex;",
      "  flex-direction: column;",
      "  gap: 6px;",
      "  max-width: 420px;",
      "}",
      ".hm-note-editor-label {",
      "  font-size: 10px;",
      "  font-weight: 700;",
      "  letter-spacing: 0.04em;",
      "  text-transform: uppercase;",
      "  color: rgba(220, 235, 255, 0.78);",
      "}",
      ".hm-note-editor textarea {",
      "  width: 100%;",
      "  min-height: 82px;",
      "  max-height: 240px;",
      "  resize: vertical;",
      "  border-radius: 10px;",
      "  border: 1px solid rgba(255,255,255,0.18);",
      "  background: rgba(8, 16, 30, 0.55);",
      "  color: #eaf4ff;",
      "  font: 500 12px/1.4 -apple-system, BlinkMacSystemFont, 'SF Pro Display', 'SF Pro Text', system-ui, sans-serif;",
      "  padding: 8px 10px;",
      "  outline: none;",
      "  box-sizing: border-box;",
      "}",
      ".hm-note-editor textarea:focus {",
      "  border-color: rgba(255, 212, 46, 0.6);",
      "  box-shadow: 0 0 0 2px rgba(255, 212, 46, 0.2);",
      "}",
      ".hm-note-editor-actions {",
      "  display: flex;",
      "  justify-content: flex-end;",
      "  gap: 6px;",
      "}",
      ".hm-note-editor-btn {",
      "  border: 1px solid rgba(255,255,255,0.18);",
      "  border-radius: 999px;",
      "  background: rgba(255,255,255,0.08);",
      "  color: rgba(255,255,255,0.92);",
      "  font: 700 10px/1 -apple-system, BlinkMacSystemFont, 'SF Pro Display', 'SF Pro Text', system-ui, sans-serif;",
      "  padding: 6px 10px;",
      "  cursor: pointer;",
      "}",
      ".hm-note-editor-btn:hover {",
      "  background: rgba(255,255,255,0.14);",
      "}",
      ".hm-note-editor-btn.primary {",
      "  border-color: rgba(255, 212, 46, 0.5);",
      "  background: rgba(255, 212, 46, 0.22);",
      "  color: #fff8de;",
      "}",
      ".hm-draw-preview {",
      "  position: absolute;",
      "  z-index: 2147483646;",
      "  pointer-events: none;",
      "  box-sizing: border-box;",
      "}"
    ].join("\n");
  }

  function normalizeColorValue(color) {
    return color && color.startsWith("custom:") ? color.replace("custom:", "") : color;
  }

  function relativeLuminance(rgb) {
    if (!rgb) return 1;
    const transform = (channel) => {
      const v = Math.max(0, Math.min(255, Number(channel) || 0)) / 255;
      return v <= 0.03928 ? (v / 12.92) : Math.pow((v + 0.055) / 1.055, 2.4);
    };
    const r = transform(rgb.r);
    const g = transform(rgb.g);
    const b = transform(rgb.b);
    return (0.2126 * r) + (0.7152 * g) + (0.0722 * b);
  }

  function blendRgbOverWhite(rgb, alpha) {
    const a = Math.max(0, Math.min(1, Number(alpha)));
    const inv = 1 - a;
    return {
      r: Math.round((rgb.r * a) + (255 * inv)),
      g: Math.round((rgb.g * a) + (255 * inv)),
      b: Math.round((rgb.b * a) + (255 * inv))
    };
  }

  function getReadableTextColor(bgColor) {
    const rgb = parseRgbColor(bgColor);
    if (!rgb) return "#111111";

    const alpha = colorAlphaValue(bgColor);
    const effectiveRgb = alpha < 0.999 ? blendRgbOverWhite(rgb, alpha) : rgb;
    const luminance = relativeLuminance(effectiveRgb);
    const contrastWithWhite = 1.05 / (luminance + 0.05);
    const contrastWithDark = (luminance + 0.05) / 0.05;

    return contrastWithDark >= contrastWithWhite ? "#111111" : "#ffffff";
  }

  function createColorPreview(initialColor) {
    const wrap = document.createElement("div");
    wrap.className = "hm-preview";
    const sample = document.createElement("span");
    sample.className = "hm-preview-text";
    sample.textContent = "Preview highlighted text";
    wrap.appendChild(sample);

    const update = (bgColor) => {
      sample.style.background = bgColor;
      sample.style.color = getReadableTextColor(bgColor);
    };

    update(initialColor);
    return { wrap, update };
  }

  function getAnnotationElements(id) {
    const normalizedId = normalizeHighlightIdValue(id);
    if (!normalizedId) return [];
    return Array.from(document.querySelectorAll("[" + HIGHLIGHT_ATTR + '="' + normalizedId + '"]'));
  }

  function getPrimaryAnnotationElement(id) {
    return getAnnotationElements(id)[0] || null;
  }

  function getEventTargetElement(event) {
    if (!event || !event.target) return null;
    const target = event.target;
    if (target.nodeType === 1) return target;
    return target.parentElement || null;
  }

  function isEventFromOcrCopyChip(event) {
    const targetEl = getEventTargetElement(event);
    return !!(targetEl && targetEl.closest && targetEl.closest(".hm-ocr-copy-chip"));
  }

  function getTextHighlightOwnerFromEvent(event) {
    const targetEl = getEventTargetElement(event);
    if (!targetEl || !targetEl.closest) return null;
    const owner = targetEl.closest("span[" + HIGHLIGHT_ATTR + "]");
    if (!owner) return null;
    const kind = owner.getAttribute(HIGHLIGHT_KIND_ATTR) || "";
    if (kind.indexOf("shape-") === 0) return null;
    return owner;
  }

  function isShapeRecord(record) {
    return record && typeof record.type === "string" && record.type.indexOf("shape-") === 0;
  }

  function parseRgbColor(colorValue) {
    const color = normalizeColorValue(colorValue || "");
    const hexMatch = color.match(/^#([0-9a-f]{6})$/i);
    if (hexMatch) {
      const hex = hexMatch[1];
      return {
        r: parseInt(hex.slice(0, 2), 16),
        g: parseInt(hex.slice(2, 4), 16),
        b: parseInt(hex.slice(4, 6), 16)
      };
    }
    const rgbMatch = color.match(/^rgb\(\s*([0-9]{1,3})\s*,\s*([0-9]{1,3})\s*,\s*([0-9]{1,3})\s*\)$/i);
    if (rgbMatch) {
      return {
        r: parseInt(rgbMatch[1], 10),
        g: parseInt(rgbMatch[2], 10),
        b: parseInt(rgbMatch[3], 10)
      };
    }
    const rgbaMatch = color.match(/^rgba\(\s*([0-9]{1,3})\s*,\s*([0-9]{1,3})\s*,\s*([0-9]{1,3})\s*,\s*([0-9.]+)\s*\)$/i);
    if (rgbaMatch) {
      return {
        r: parseInt(rgbaMatch[1], 10),
        g: parseInt(rgbaMatch[2], 10),
        b: parseInt(rgbaMatch[3], 10)
      };
    }
    return null;
  }

  function colorAlphaValue(colorValue) {
    const color = normalizeColorValue(colorValue || "");
    const rgbaMatch = color.match(/^rgba\(\s*[0-9]{1,3}\s*,\s*[0-9]{1,3}\s*,\s*[0-9]{1,3}\s*,\s*([0-9.]+)\s*\)$/i);
    if (!rgbaMatch) return 1;
    const parsed = parseFloat(rgbaMatch[1]);
    return Number.isFinite(parsed) ? parsed : 1;
  }

  function withAlphaColor(colorValue, alpha) {
    const rgb = parseRgbColor(colorValue);
    if (!rgb) return normalizeColorValue(colorValue || "#ffd42e");
    return "rgba(" + rgb.r + ", " + rgb.g + ", " + rgb.b + ", " + alpha + ")";
  }

  function getSolidColorFromName(colorName, fallbackColor) {
    if (colorName && colorName.startsWith("custom:")) return colorName.replace("custom:", "");
    if (PRESET_COLORS[colorName]) return PRESET_COLORS[colorName].solid;
    return normalizeColorValue(fallbackColor || PRESET_COLORS.yellow.solid);
  }

  function getRectBgForState(colorName, fallbackColor, revealed) {
    const solid = getSolidColorFromName(colorName, fallbackColor);
    if (revealed) return withAlphaColor(solid, 0.28);
    return solid;
  }

  function normalizeComparableUrl(urlValue) {
    if (!urlValue) return "";
    try {
      const url = new URL(urlValue, location.href);
      url.hash = "";
      return url.toString();
    } catch (_err) {
      return String(urlValue).split("#")[0];
    }
  }

  function isCurrentPageUrl(urlValue) {
    return normalizeComparableUrl(urlValue) === normalizeComparableUrl(location.href);
  }

  function buildTextSearchIndex(root) {
    if (!root) return null;
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
      acceptNode(node) {
        if (!node || !node.textContent || !normalizeWS(node.textContent)) {
          return NodeFilter.FILTER_REJECT;
        }
        const parent = node.parentElement;
        if (!parent) return NodeFilter.FILTER_REJECT;
        const tag = (parent.tagName || "").toLowerCase();
        if (tag === "script" || tag === "style" || tag === "noscript" || tag === "textarea") {
          return NodeFilter.FILTER_REJECT;
        }
        return NodeFilter.FILTER_ACCEPT;
      }
    });

    let node;
    let fullText = "";
    const nodeMap = [];
    while ((node = walker.nextNode())) {
      const start = fullText.length;
      fullText += node.textContent;
      nodeMap.push({ node, start, end: fullText.length });
    }

    return {
      fullText,
      nodeMap,
      normalizedFull: normalizeWS(fullText)
    };
  }

  function normalizeShapeRecord(record, fromStorage) {
    if (!isShapeRecord(record)) return record;
    if (record.type === "shape-rect") {
      if (typeof record.revealed !== "boolean") {
        const inferredAlpha = colorAlphaValue(record.bgColor || "");
        record.revealed = inferredAlpha < 0.95;
      }
      if (typeof record.transient !== "boolean") {
        record.transient = false;
      }
      record.color = record.color || "yellow";
      const storedBgColor = normalizeColorValue(record.bgColor || "");
      const expectedBgColor = getRectBgForState(record.color, storedBgColor || PRESET_COLORS.yellow.solid, record.revealed);
      if (!storedBgColor) {
        record.bgColor = expectedBgColor;
      } else {
        const currentAlpha = colorAlphaValue(storedBgColor);
        const alphaMatchesState = record.revealed ? currentAlpha < 0.95 : currentAlpha >= 0.95;
        record.bgColor = alphaMatchesState ? storedBgColor : expectedBgColor;
      }
      if (!record.text || record.text === "Rectangle Highlight") {
        record.text = record.ocrText || "Rectangle Highlight";
      }
      if (typeof record.ocrText !== "string") {
        record.ocrText = record.text === "Rectangle Highlight" ? "" : (record.text || "");
      }
      if (typeof record.ocrFullText !== "string") {
        record.ocrFullText = record.ocrText || "";
      }
      if (typeof record.label !== "string" || !record.label.trim()) {
        record.label = (record.text && record.text !== "Rectangle Highlight")
          ? record.text
          : (record.ocrText || "Rectangle Highlight");
      }
    } else if (record.type === "shape-cover") {
      if (typeof record.transient !== "boolean") record.transient = !fromStorage;
      record.bgColor = normalizeColorValue(record.bgColor || getSolidColorFromName(record.color, PRESET_COLORS.yellow.solid));
    }
    return record;
  }

  function getShapeFillColor(record, bgColor) {
    if (!record) return bgColor;
    if (record.type === "shape-cover") {
      return normalizeColorValue(bgColor);
    }
    if (record.type === "shape-rect") {
      return getRectBgForState(record.color, bgColor, !!record.revealed);
    }
    return normalizeColorValue(bgColor);
  }

  function getColorInputValue(colorName, fallbackColor) {
    if (colorName && colorName.startsWith("custom:")) return colorName.replace("custom:", "");
    if (PRESET_COLORS[colorName]) return PRESET_COLORS[colorName].solid;
    return normalizeColorValue(fallbackColor || "#ffd42e");
  }

  function getRectangleCopyableOcrText(record) {
    if (!record || record.type !== "shape-rect" || !record.revealed) return "";
    const full = normalizeSnippetText(record.ocrFullText || "");
    if (full) return full;
    const primary = normalizeSnippetText(record.ocrText || "");
    if (primary) return primary;
    const fallback = normalizeSnippetText(record.text || "");
    return fallback === "Rectangle Highlight" ? "" : fallback;
  }

  function maybeAutoCopyRectangleOcrText(record) {
    if (!autoOcrCopyEnabled) return;
    const text = getRectangleCopyableOcrText(record);
    if (!text) return;
    copyTextToClipboardInPage(text, "OCR text copied automatically");
  }

  function ensureRectangleFullOcrText(record) {
    if (!record || record.type !== "shape-rect" || !record.revealed) return "";
    const current = normalizeSnippetText(record.ocrFullText || "");
    if (current) return current;
    const visualRect = getRenderedRectForRecord(record) || record.rect;
    const fullText = extractRectangleCopyText(visualRect);
    if (fullText) {
      record.ocrFullText = fullText;
      return fullText;
    }
    return normalizeSnippetText(record.ocrText || "");
  }

  function getOcrCopySnippet(text, maxChars) {
    const clean = normalizeSnippetText(text || "");
    if (!clean) return "";
    const limit = Math.max(12, Number(maxChars) || 36);
    if (clean.length <= limit) return clean;
    return clean.slice(0, limit - 3).trimEnd() + "...";
  }

  function fallbackExecCopyText(payload) {
    let copied = false;
    const copyListener = (event) => {
      if (!event || !event.clipboardData) return;
      event.preventDefault();
      event.clipboardData.setData("text/plain", payload);
      copied = true;
    };
    document.addEventListener("copy", copyListener, true);
    try {
      copied = document.execCommand("copy") || copied;
    } catch (_err) {}
    document.removeEventListener("copy", copyListener, true);
    if (copied) return true;

    const ta = document.createElement("textarea");
    ta.value = payload;
    ta.style.position = "fixed";
    ta.style.opacity = "0";
    ta.style.pointerEvents = "none";
    ta.style.left = "-9999px";
    ta.style.top = "0";
    (document.body || document.documentElement).appendChild(ta);
    try { ta.focus({ preventScroll: true }); } catch (_err) {}
    ta.select();
    try { ta.setSelectionRange(0, ta.value.length); } catch (_err) {}
    try {
      copied = document.execCommand("copy") || copied;
    } catch (_err) {}
    if (ta.parentNode) ta.parentNode.removeChild(ta);
    return copied;
  }

  function copyTextToClipboardInPage(text, successMessage) {
    const payload = String(text || "");
    if (!payload) return;
    const done = () => showInlineUndoToast(successMessage || "Copied", null, 2200);
    const fail = () => showInlineUndoToast("Copy failed. Please try again.", null, 2600);
    const fallback = () => {
      if (fallbackExecCopyText(payload)) done();
      else fail();
    };

    if (navigator.clipboard && typeof navigator.clipboard.writeText === "function") {
      navigator.clipboard.writeText(payload).then(done).catch(fallback);
      return;
    }
    fallback();
  }

  function getHighlightTextCopyPayload(highlightId) {
    const normalizedId = normalizeHighlightIdValue(highlightId);
    if (!normalizedId) return "";
    const record = storedHighlightsForPage.find((item) => idsMatch(item && item.id, normalizedId));
    if (record && !isShapeRecord(record)) {
      const recordText = normalizeSnippetText(record.text || "");
      if (recordText) return recordText;
    }
    const spans = Array.from(document.querySelectorAll("span[" + HIGHLIGHT_ATTR + '="' + normalizedId + '"]'))
      .filter((node) => {
        const kind = node.getAttribute(HIGHLIGHT_KIND_ATTR) || "";
        return kind.indexOf("shape-") !== 0;
      });
    if (spans.length) {
      const joined = normalizeSnippetText(spans.map((node) => node.textContent || "").join(" "));
      if (joined) return joined;
    }
    return "";
  }

  function hideTextHighlightContextMenu() {
    if (textHighlightContextMenuEl && textHighlightContextMenuEl.parentNode) {
      textHighlightContextMenuEl.parentNode.removeChild(textHighlightContextMenuEl);
    }
    textHighlightContextMenuEl = null;
    textHighlightContextMenuId = "";
  }

  function showTextHighlightContextMenu(highlightId, pageX, pageY) {
    const copyText = getHighlightTextCopyPayload(highlightId);
    if (!copyText) return;
    hideTextHighlightContextMenu();

    const menu = document.createElement("div");
    menu.className = "hm-highlight-context-menu";
    menu.setAttribute("role", "menu");
    menu.style.left = Math.max(4, Number(pageX) || 0) + "px";
    menu.style.top = Math.max(4, Number(pageY) || 0) + "px";

    const copyBtn = document.createElement("button");
    copyBtn.type = "button";
    copyBtn.textContent = "Copy same-color text";
    copyBtn.setAttribute("role", "menuitem");
    copyBtn.addEventListener("mousedown", (event) => {
      event.preventDefault();
      event.stopPropagation();
      event.stopImmediatePropagation();
    });
    copyBtn.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      event.stopImmediatePropagation();
      const targetText = getHighlightTextCopyPayload(textHighlightContextMenuId || highlightId);
      hideTextHighlightContextMenu();
      copyTextToClipboardInPage(targetText, "Highlighted text copied");
    });

    menu.appendChild(copyBtn);
    (document.body || document.documentElement).appendChild(menu);
    textHighlightContextMenuEl = menu;
    textHighlightContextMenuId = normalizeHighlightIdValue(highlightId);

    const menuRect = menu.getBoundingClientRect();
    const maxLeft = Math.max(4, window.scrollX + window.innerWidth - menuRect.width - 6);
    const maxTop = Math.max(4, window.scrollY + window.innerHeight - menuRect.height - 6);
    const clampedLeft = Math.max(window.scrollX + 4, Math.min(pageX, maxLeft));
    const clampedTop = Math.max(window.scrollY + 4, Math.min(pageY, maxTop));
    menu.style.left = clampedLeft + "px";
    menu.style.top = clampedTop + "px";
  }

  function ensureRectangleOcrCopyChip(el, record) {
    if (!el || !record || record.type !== "shape-rect") return;
    const existing = el.querySelector(".hm-ocr-copy-chip");
    const ocrText = getRectangleCopyableOcrText(record);
    if (!ocrText) {
      if (existing && existing.parentNode) existing.parentNode.removeChild(existing);
      return;
    }

    const chip = existing || document.createElement("button");
    if (!existing) {
      chip.type = "button";
      chip.className = "hm-ocr-copy-chip";
      chip.addEventListener("mousedown", (event) => {
        event.stopPropagation();
        event.stopImmediatePropagation();
      });
      chip.addEventListener("click", (event) => {
        event.stopPropagation();
        event.stopImmediatePropagation();
        setDrawMode(null);
        const id = el.getAttribute ? (el.getAttribute(HIGHLIGHT_ATTR) || "") : "";
        const liveRecord = storedHighlightsForPage.find((item) => idsMatch(item && item.id, id) && item.type === "shape-rect");
        if (liveRecord) {
          normalizeShapeRecord(liveRecord, false);
          ensureRectangleFullOcrText(liveRecord);
          ensureRectangleOcrCopyChip(el, liveRecord);
        }
        const textToCopy = chip.getAttribute("data-hm-ocr-text") || "";
        copyTextToClipboardInPage(textToCopy, "OCR text copied");
      });
      el.appendChild(chip);
    }
    const snippet = getOcrCopySnippet(ocrText, 36);
    chip.setAttribute("data-hm-ocr-text", ocrText);
    chip.setAttribute("title", ocrText);
    let prefixEl = chip.querySelector(".hm-ocr-copy-prefix");
    if (!prefixEl) {
      prefixEl = document.createElement("span");
      prefixEl.className = "hm-ocr-copy-prefix";
      chip.appendChild(prefixEl);
    }
    prefixEl.textContent = "Copy OCR";

    let snippetEl = chip.querySelector(".hm-ocr-copy-snippet");
    if (!snippetEl) {
      snippetEl = document.createElement("span");
      snippetEl.className = "hm-ocr-copy-snippet";
      chip.appendChild(snippetEl);
    }
    snippetEl.textContent = snippet;
  }

  function applyShapeVisual(el, record, bgColor) {
    normalizeShapeRecord(record, false);
    const fillColor = getShapeFillColor(record, bgColor);
    const solidColor = normalizeColorValue(
      record.color && record.color.startsWith("custom:")
        ? record.color.replace("custom:", "")
        : (PRESET_COLORS[record.color] ? PRESET_COLORS[record.color].solid : fillColor)
    );
    el.style.setProperty("background-color", fillColor, "important");
    el.style.setProperty("color", solidColor, "important");
    el.style.setProperty("opacity", "1", "important");
    if (record.type === "shape-rect") {
      el.setAttribute("data-hm-revealed", record.revealed ? "1" : "0");
      el.style.setProperty("border-color", solidColor, "important");
      el.style.setProperty("border-style", record.revealed ? "dashed" : "solid", "important");
      el.style.setProperty("border-width", "2px", "important");
      el.style.setProperty(
        "box-shadow",
        record.revealed
          ? "0 0 0 1px rgba(0,0,0,0.16), inset 0 0 0 1px rgba(255,255,255,0.42)"
          : "0 0 0 1px rgba(0,0,0,0.14)",
        "important"
      );
      ensureRectangleOcrCopyChip(el, record);
    } else {
      el.removeAttribute("data-hm-revealed");
      const chip = el.querySelector(".hm-ocr-copy-chip");
      if (chip && chip.parentNode) chip.parentNode.removeChild(chip);
    }
  }

  // ---- Shadow DOM Picker (for new highlights) ----

  function ensurePickerHost() {
    if (pickerHost && pickerHost.parentNode) return;
    pickerHost = document.createElement("div");
    pickerHost.id = "hm-picker-host";
    pickerShadow = pickerHost.attachShadow({ mode: "closed" });

    const style = document.createElement("style");
    style.textContent = getPickerStyles();
    pickerShadow.appendChild(style);

    (document.body || document.documentElement).appendChild(pickerHost);
  }

  function showPicker(rect) {
    if (!extensionEnabled) return;
    ensurePickerHost();

    if (pickerEl && pickerEl.parentNode) pickerEl.remove();
    if (currentSelection && typeof currentSelection.note === "string" && !pickerNoteDraft) {
      pickerNoteDraft = currentSelection.note;
    }

    const picker = document.createElement("div");
    picker.className = "hm-picker";
    const controlsRow = document.createElement("div");
    controlsRow.className = "hm-controls-row";
    picker.appendChild(controlsRow);
    const preview = createColorPreview(lastUsedBgColor);

    for (const [name, colors] of Object.entries(PRESET_COLORS)) {
      const btn = document.createElement("div");
      btn.className = "hm-color-btn";
      btn.style.background = colors.solid;
      const shortcutIndex = QUICK_SHORTCUT_COLORS.indexOf(name);
      const shortcutHint = shortcutIndex >= 0 ? " (Alt+" + (shortcutIndex + 1) + ")" : "";
      btn.title = name.charAt(0).toUpperCase() + name.slice(1) + shortcutHint;
      btn.setAttribute("aria-label", btn.title);
      btn.addEventListener("mousedown", (e) => { e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation(); });
      btn.addEventListener("click", (e) => {
        e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
        handleColorClick(name, colors.bg, pickerNoteDraft);
      });
      btn.addEventListener("mouseenter", () => { preview.update(colors.bg); });
      btn.addEventListener("mouseenter", () => { btn.style.boxShadow = "0 0 12px " + colors.solid + "88"; });
      btn.addEventListener("mouseleave", () => { preview.update(lastUsedBgColor); });
      btn.addEventListener("mouseleave", () => { btn.style.boxShadow = "none"; });
      controlsRow.appendChild(btn);
    }

    const divider = document.createElement("div");
    divider.className = "hm-divider";
    controlsRow.appendChild(divider);

    const customWrap = document.createElement("div");
    customWrap.className = "hm-custom-wrap";
    const customBtn = document.createElement("div");
    customBtn.className = "hm-custom-btn";
    customBtn.title = "Custom color";
    customBtn.style.background = "linear-gradient(135deg, #ffffff 0%, #f4f7fb 18%, #ffd42e 18%, #ff9f43 38%, #ff7eb6 58%, #4f9fe8 78%, #33c3ac 100%)";
    const colorInput = document.createElement("input");
    colorInput.type = "color";
    colorInput.className = "hm-color-input";
    colorInput.value = "#ffd42e";
    colorInput.title = "Pick any color";
    colorInput.addEventListener("mousedown", (e) => { e.stopPropagation(); e.stopImmediatePropagation(); });
    colorInput.addEventListener("click", (e) => { e.stopPropagation(); e.stopImmediatePropagation(); colorPickerOpen = true; });
    colorInput.addEventListener("focus", () => { colorPickerOpen = true; });
    colorInput.addEventListener("blur", () => { setTimeout(() => { colorPickerOpen = false; }, 300); });
    colorInput.addEventListener("input", (e) => {
      const hex = colorInput.value;
      customBtn.style.background = hex;
      preview.update(hex);
    });
    colorInput.addEventListener("change", (e) => {
      e.stopPropagation(); e.stopImmediatePropagation();
      const hex = colorInput.value;
      colorPickerOpen = false;
      handleColorClick("custom:" + hex, hex, pickerNoteDraft);
    });
    customWrap.appendChild(customBtn);
    customWrap.appendChild(colorInput);
    controlsRow.appendChild(customWrap);

    const divider2 = document.createElement("div");
    divider2.className = "hm-divider";
    controlsRow.appendChild(divider2);

    const noteBtn = document.createElement("button");
    const hasDraftNote = !!String(pickerNoteDraft || "").trim();
    noteBtn.className = "hm-state-btn" + ((pickerNoteEditorOpen || hasDraftNote) ? " active" : "");
    noteBtn.textContent = "N";
    noteBtn.title = "Add comment (applies default yellow first)";
    noteBtn.setAttribute("aria-label", noteBtn.title);
    noteBtn.addEventListener("mousedown", (e) => { e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation(); });
    noteBtn.addEventListener("click", (e) => {
      e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
      applyDefaultHighlightForNoteFlow();
    });
    controlsRow.appendChild(noteBtn);

    picker.appendChild(preview.wrap);

    if (pickerNoteEditorOpen) {
      const editorWrap = document.createElement("div");
      editorWrap.className = "hm-note-editor";
      const label = document.createElement("div");
      label.className = "hm-note-editor-label";
      label.textContent = "Comment";
      editorWrap.appendChild(label);
      const textarea = document.createElement("textarea");
      textarea.placeholder = "Write your comment (no limit)...";
      textarea.value = String(pickerNoteDraft || "");
      textarea.addEventListener("mousedown", (e) => { e.stopPropagation(); });
      textarea.addEventListener("click", (e) => { e.stopPropagation(); });
      textarea.addEventListener("input", () => {
        pickerNoteDraft = textarea.value;
        if (currentSelection) currentSelection.note = pickerNoteDraft;
      });
      editorWrap.appendChild(textarea);

      const actionRow = document.createElement("div");
      actionRow.className = "hm-note-editor-actions";

      const clearBtn = document.createElement("button");
      clearBtn.className = "hm-note-editor-btn";
      clearBtn.textContent = "Clear";
      clearBtn.addEventListener("mousedown", (e) => { e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation(); });
      clearBtn.addEventListener("click", (e) => {
        e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
        textarea.value = "";
        pickerNoteDraft = "";
        if (currentSelection) currentSelection.note = "";
      });
      actionRow.appendChild(clearBtn);

      const closeBtn = document.createElement("button");
      closeBtn.className = "hm-note-editor-btn";
      closeBtn.textContent = "Done";
      closeBtn.addEventListener("mousedown", (e) => { e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation(); });
      closeBtn.addEventListener("click", (e) => {
        e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
        pickerNoteEditorOpen = false;
        showPicker(rect);
      });
      actionRow.appendChild(closeBtn);

      editorWrap.appendChild(actionRow);
      picker.appendChild(editorWrap);
      setTimeout(() => {
        try { textarea.focus(); } catch (_e) {}
      }, 0);
    }

    pickerShadow.appendChild(picker);
    pickerEl = picker;

    positionElement(picker, rect, 360, pickerNoteEditorOpen ? 232 : 84);

    requestAnimationFrame(() => {
      requestAnimationFrame(() => { picker.classList.add("visible"); });
    });
  }

  function hidePicker() {
    if (colorPickerOpen) return;
    if (pickerEl) {
      pickerEl.classList.remove("visible");
      const el = pickerEl;
      setTimeout(() => { if (el.parentNode) el.remove(); }, 220);
      pickerEl = null;
    }
    pickerNoteEditorOpen = false;
    pickerNoteDraft = "";
  }

  // ---- Hover Toolbar (for existing highlights) ----

  function ensureHoverHost() {
    if (hoverHost && hoverHost.parentNode) return;
    hoverHost = document.createElement("div");
    hoverHost.id = "hm-hover-host";
    hoverShadow = hoverHost.attachShadow({ mode: "closed" });

    const style = document.createElement("style");
    style.textContent = getPickerStyles();
    hoverShadow.appendChild(style);

    (document.body || document.documentElement).appendChild(hoverHost);
  }

  function showHoverToolbar(highlightId, rect) {
    if (!extensionEnabled) return;
    ensureHoverHost();

    if (hoverEl && hoverEl.parentNode) hoverEl.remove();
    activeHoverId = highlightId;

    const toolbar = document.createElement("div");
    toolbar.className = "hm-picker";
    const controlsRow = document.createElement("div");
    controlsRow.className = "hm-controls-row";
    toolbar.appendChild(controlsRow);

    const currentRecord = storedHighlightsForPage.find((x) => idsMatch(x && x.id, highlightId)) || null;
    if (currentRecord && isShapeRecord(currentRecord)) {
      normalizeShapeRecord(currentRecord, false);
    }
    const isRectHighlight = !!(currentRecord && currentRecord.type === "shape-rect");
    const isCoverShape = !!(currentRecord && currentRecord.type === "shape-cover");
    const currentEl = getPrimaryAnnotationElement(highlightId);
    const currentColor = currentRecord ? (currentRecord.color || "yellow") : (currentEl ? (currentEl.getAttribute(HIGHLIGHT_COLOR_ATTR) || "yellow") : "yellow");
    const fallbackRecordBg = currentRecord
      ? (currentRecord.type === "shape-rect"
        ? getRectBgForState(currentColor, PRESET_COLORS.yellow.solid, !!currentRecord.revealed)
        : getSolidColorFromName(currentColor, PRESET_COLORS.yellow.solid))
      : PRESET_COLORS.yellow.bg;
    const currentBg = currentRecord
      ? (currentRecord.bgColor || fallbackRecordBg)
      : (currentEl
        ? (currentEl.style.backgroundColor || ((currentColor && currentColor.startsWith("custom:")) ? currentColor.replace("custom:", "") : (PRESET_COLORS[currentColor] ? PRESET_COLORS[currentColor].bg : PRESET_COLORS.yellow.bg)))
        : PRESET_COLORS.yellow.bg);
    const preview = createColorPreview(currentBg);

    for (const [name, colors] of Object.entries(PRESET_COLORS)) {
      const btn = document.createElement("div");
      btn.className = "hm-color-btn";
      btn.style.background = colors.solid;
      if (name === currentColor) {
        btn.style.borderColor = "rgba(255,255,255,0.8)";
        btn.style.boxShadow = "0 0 8px " + colors.solid + "66";
      }
      const shortcutIndex = QUICK_SHORTCUT_COLORS.indexOf(name);
      const shortcutHint = shortcutIndex >= 0 ? " (Alt+" + (shortcutIndex + 1) + ")" : "";
      btn.title = name.charAt(0).toUpperCase() + name.slice(1) + shortcutHint;
      btn.setAttribute("aria-label", btn.title);
      btn.addEventListener("mousedown", (e) => { e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation(); });
      btn.addEventListener("click", (e) => {
        e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
        const actionBg = (isRectHighlight || isCoverShape) ? colors.solid : colors.bg;
        changeHighlightColor(highlightId, name, actionBg);
        refreshHoverToolbar(highlightId);
      });
      btn.addEventListener("mouseenter", () => {
        const previewBg = isRectHighlight
          ? getRectBgForState(name, colors.solid, !!(currentRecord && currentRecord.revealed))
          : ((isCoverShape ? colors.solid : colors.bg));
        preview.update(previewBg);
      });
      btn.addEventListener("mouseenter", () => { btn.style.boxShadow = "0 0 12px " + colors.solid + "88"; });
      btn.addEventListener("mouseleave", () => {
        preview.update(currentBg);
        if (name === currentColor) {
          btn.style.boxShadow = "0 0 8px " + colors.solid + "66";
        } else {
          btn.style.boxShadow = "none";
        }
      });
      controlsRow.appendChild(btn);
    }

    const divider = document.createElement("div");
    divider.className = "hm-divider";
    controlsRow.appendChild(divider);

    const customWrap = document.createElement("div");
    customWrap.className = "hm-custom-wrap";
    const customBtn = document.createElement("div");
    customBtn.className = "hm-custom-btn";
    customBtn.title = "Custom color";
    let customInputPreviewing = false;
    const resetCustomButtonSwatch = () => {
      customBtn.style.background = "linear-gradient(135deg, #ffffff 0%, #f4f7fb 18%, #ffd42e 18%, #ff9f43 38%, #ff7eb6 58%, #4f9fe8 78%, #33c3ac 100%)";
    };
    resetCustomButtonSwatch();
    const colorInput = document.createElement("input");
    colorInput.type = "color";
    colorInput.className = "hm-color-input";
    colorInput.value = getColorInputValue(currentColor, PRESET_COLORS.yellow.solid);
    colorInput.addEventListener("mousedown", (e) => { e.stopPropagation(); e.stopImmediatePropagation(); });
    colorInput.addEventListener("click", (e) => { e.stopPropagation(); e.stopImmediatePropagation(); hoverColorPickerOpen = true; });
    colorInput.addEventListener("focus", () => { hoverColorPickerOpen = true; });
    colorInput.addEventListener("blur", () => {
      setTimeout(() => {
        hoverColorPickerOpen = false;
        if (!customInputPreviewing) return;
        customInputPreviewing = false;
        resetCustomButtonSwatch();
        changeHighlightColor(highlightId, currentColor, currentBg, { previewOnly: true });
        preview.update(currentBg);
      }, 300);
    });
    colorInput.addEventListener("input", (e) => {
      const hex = colorInput.value;
      customBtn.style.background = hex;
      const previewBg = isRectHighlight
        ? getRectBgForState("custom:" + hex, hex, !!(currentRecord && currentRecord.revealed))
        : hex;
      preview.update(previewBg);
      customInputPreviewing = true;
      changeHighlightColor(highlightId, "custom:" + hex, hex, { previewOnly: true });
    });
    colorInput.addEventListener("change", (e) => {
      e.stopPropagation(); e.stopImmediatePropagation();
      const hex = colorInput.value;
      hoverColorPickerOpen = false;
       customInputPreviewing = false;
      changeHighlightColor(highlightId, "custom:" + hex, hex);
      refreshHoverToolbar(highlightId);
    });
    customWrap.appendChild(customBtn);
    customWrap.appendChild(colorInput);
    controlsRow.appendChild(customWrap);

    const dividerNote = document.createElement("div");
    dividerNote.className = "hm-divider";
    controlsRow.appendChild(dividerNote);

    const currentNote = currentRecord && typeof currentRecord.note === "string" ? currentRecord.note : "";
    const hasNote = !!currentNote.trim();
    const noteEditorOpen = hoverNoteEditorId === highlightId;
    const noteBtn = document.createElement("button");
    noteBtn.className = "hm-state-btn" + ((hasNote || noteEditorOpen) ? " active" : "");
    noteBtn.textContent = "N";
    noteBtn.title = noteEditorOpen ? "Close comment editor" : (hasNote ? "View or edit comment" : "Add comment");
    noteBtn.setAttribute("aria-label", noteBtn.title);
    noteBtn.addEventListener("mousedown", (e) => { e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation(); });
    noteBtn.addEventListener("click", (e) => {
      e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
      if (hoverNoteEditorId === highlightId) {
        hoverNoteEditorId = "";
        hoverNoteDraft = "";
      } else {
        hoverNoteEditorId = highlightId;
        hoverNoteDraft = currentNote;
      }
      refreshHoverToolbar(highlightId);
    });
    controlsRow.appendChild(noteBtn);

    if (isRectHighlight) {
      const dividerState = document.createElement("div");
      dividerState.className = "hm-divider";
      controlsRow.appendChild(dividerState);

      const transparencyBtn = document.createElement("button");
      transparencyBtn.className = "hm-state-btn" + (currentRecord.revealed ? " active" : "");
      transparencyBtn.textContent = "T";
      transparencyBtn.title = currentRecord.revealed ? "Set solid (hide text)" : "Set transparent (save)";
      transparencyBtn.addEventListener("mousedown", (e) => { e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation(); });
      transparencyBtn.addEventListener("click", (e) => {
        e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
        setRectangleRevealState(highlightId, !currentRecord.revealed);
        refreshHoverToolbar(highlightId);
      });
      controlsRow.appendChild(transparencyBtn);
    }

    const divider2 = document.createElement("div");
    divider2.className = "hm-divider";
    controlsRow.appendChild(divider2);

    const deleteBtn = document.createElement("div");
    deleteBtn.className = "hm-delete-btn";
    deleteBtn.title = "Delete all stacked highlights here";
    deleteBtn.innerHTML = '<svg viewBox="0 0 24 24"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>';
    deleteBtn.addEventListener("mousedown", (e) => { e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation(); });
    deleteBtn.addEventListener("click", (e) => {
      e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
      
      const idsToDelete = new Set([highlightId]);
      let added = true;
      while (added) {
          added = false;
          const currentIds = Array.from(idsToDelete);
          for (const id of currentIds) {
              const spans = document.querySelectorAll("[" + HIGHLIGHT_ATTR + '="' + id + '"]');
              for (const span of spans) {
                  // Ancestors
                  let curr = span.parentElement;
                  while (curr && curr.hasAttribute && curr.hasAttribute(HIGHLIGHT_ATTR)) {
                      const pId = curr.getAttribute(HIGHLIGHT_ATTR);
                      if (!idsToDelete.has(pId)) { idsToDelete.add(pId); added = true; }
                      curr = curr.parentElement;
                  }
                  // Descendants
                  const desc = span.querySelectorAll("[" + HIGHLIGHT_ATTR + "]");
                  for (const d of desc) {
                      const dId = d.getAttribute(HIGHLIGHT_ATTR);
                      if (!idsToDelete.has(dId)) { idsToDelete.add(dId); added = true; }
                  }
              }
          }
      }
      
      deleteHighlightsWithUndo(Array.from(idsToDelete));
      hideHoverToolbar();
    });
    controlsRow.appendChild(deleteBtn);
    toolbar.appendChild(preview.wrap);
    if (noteEditorOpen) {
      const editorWrap = document.createElement("div");
      editorWrap.className = "hm-note-editor";
      const label = document.createElement("div");
      label.className = "hm-note-editor-label";
      label.textContent = "Comment (auto-save)";
      editorWrap.appendChild(label);
      const textarea = document.createElement("textarea");
      textarea.placeholder = "Write your comment...";
      textarea.value = hoverNoteDraft;
      let autoSaveTimer = null;
      const queueAutoSave = (flushNow) => {
        if (autoSaveTimer) {
          clearTimeout(autoSaveTimer);
          autoSaveTimer = null;
        }
        const commit = () => {
          saveHighlightComment(highlightId, textarea.value, { silentToast: true });
        };
        if (flushNow) {
          commit();
        } else {
          autoSaveTimer = setTimeout(commit, 360);
        }
      };
      textarea.addEventListener("mousedown", (e) => { e.stopPropagation(); });
      textarea.addEventListener("click", (e) => { e.stopPropagation(); });
      textarea.addEventListener("input", () => {
        hoverNoteDraft = textarea.value;
        queueAutoSave(false);
      });
      textarea.addEventListener("blur", () => {
        queueAutoSave(true);
      });
      textarea.addEventListener("keydown", (e) => {
        if (e.key === "Escape") {
          e.preventDefault();
          queueAutoSave(true);
          hoverNoteEditorId = "";
          hoverNoteDraft = "";
          refreshHoverToolbar(highlightId);
          return;
        }
        if ((e.ctrlKey || e.metaKey) && (e.key || "").toLowerCase() === "enter") {
          e.preventDefault();
          queueAutoSave(true);
          hoverNoteEditorId = "";
          hoverNoteDraft = "";
          refreshHoverToolbar(highlightId);
        }
      });
      editorWrap.appendChild(textarea);

      const actionRow = document.createElement("div");
      actionRow.className = "hm-note-editor-actions";

      const clearBtn = document.createElement("button");
      clearBtn.className = "hm-note-editor-btn";
      clearBtn.textContent = "Clear";
      clearBtn.addEventListener("mousedown", (e) => { e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation(); });
      clearBtn.addEventListener("click", (e) => {
        e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
        textarea.value = "";
        hoverNoteDraft = "";
        queueAutoSave(false);
      });
      actionRow.appendChild(clearBtn);

      const doneBtn = document.createElement("button");
      doneBtn.className = "hm-note-editor-btn primary";
      doneBtn.textContent = "Done";
      doneBtn.addEventListener("mousedown", (e) => { e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation(); });
      doneBtn.addEventListener("click", (e) => {
        e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
        queueAutoSave(true);
        hoverNoteEditorId = "";
        hoverNoteDraft = "";
        refreshHoverToolbar(highlightId);
      });
      actionRow.appendChild(doneBtn);

      editorWrap.appendChild(actionRow);
      toolbar.appendChild(editorWrap);
      setTimeout(() => {
        try { textarea.focus(); } catch (_e) {}
      }, 0);
    } else if (hasNote) {
      const notePreview = document.createElement("div");
      notePreview.className = "hm-note-preview";
      notePreview.textContent = formatNotePreview(currentNote);
      notePreview.title = currentNote;
      toolbar.appendChild(notePreview);
    }

    hoverShadow.appendChild(toolbar);
    hoverEl = toolbar;

    const estimatedHeight = noteEditorOpen ? 248 : (hasNote ? 116 : 84);
    positionElement(toolbar, rect, isRectHighlight ? 460 : 360, estimatedHeight);

    requestAnimationFrame(() => {
      requestAnimationFrame(() => { toolbar.classList.add("visible"); });
    });
  }

  function hideHoverToolbar() {
    if (hoverColorPickerOpen) return;
    if (hoverEl) {
      hoverEl.classList.remove("visible");
      const el = hoverEl;
      setTimeout(() => { if (el.parentNode) el.remove(); }, 220);
      hoverEl = null;
    }
    activeHoverId = null;
    activeHoverTarget = null;
    hoverNoteEditorId = "";
    hoverNoteDraft = "";
  }

  function refreshHoverToolbar(highlightId) {
    if (!highlightId) return;
    const target = getPrimaryAnnotationElement(highlightId);
    if (!target) {
      hideHoverToolbar();
      return;
    }
    const rect = target.getBoundingClientRect();
    showHoverToolbar(highlightId, rect);
  }

  function setRectangleRevealState(id, revealed) {
    const record = storedHighlightsForPage.find((x) => idsMatch(x && x.id, id) && x.type === "shape-rect");
    if (!record) return;
    normalizeShapeRecord(record, false);
    const nextRevealed = !!revealed;
    if (record.revealed === nextRevealed) return;

    record.revealed = nextRevealed;
    record.bgColor = getRectBgForState(record.color, record.bgColor, record.revealed);
    const visualRect = getRenderedRectForRecord(record) || record.rect;

    if (record.revealed) {
      const snippet = extractRectangleSnippet(visualRect);
      const fullText = extractRectangleCopyText(visualRect);
      if (snippet) {
        record.ocrText = snippet;
        record.text = snippet;
      }
      if (fullText) {
        record.ocrFullText = fullText;
      } else if (snippet && !record.ocrFullText) {
        record.ocrFullText = snippet;
      }
    }
    record.label = getSmartRectangleLabel(visualRect, record.ocrText || record.ocrFullText || record.text || "");

    const spans = getAnnotationElements(id);
    for (const span of spans) {
      span.setAttribute(HIGHLIGHT_COLOR_ATTR, record.color);
      applyShapeVisual(span, record, record.bgColor);
    }

    if (record.revealed) {
      maybeAutoCopyRectangleOcrText(record);
    }

    if (record.transient) {
      record.transient = false;
      sendRuntimeMessageSafe({ action: "saveHighlight", highlight: record });
      return;
    }

    sendRuntimeMessageSafe({
      action: "saveHighlight",
      highlight: record,
      historyMeta: {
        action: "update",
        label: "Updated rectangle visibility"
      }
    });
  }

  function changeHighlightColor(id, newColorName, newBgColor, options) {
    const opts = options && typeof options === "object" ? options : {};
    const previewOnly = !!opts.previewOnly;
    const record = storedHighlightsForPage.find((x) => idsMatch(x && x.id, id));
    if (record && isShapeRecord(record) && !previewOnly) {
      normalizeShapeRecord(record, false);
    }
    const visualRecord = record && isShapeRecord(record)
      ? (previewOnly ? normalizeShapeRecord({ ...record }, false) : record)
      : record;

    let effectiveBgColor = newBgColor;
    if (visualRecord && visualRecord.type === "shape-rect") {
      effectiveBgColor = getRectBgForState(newColorName, newBgColor, !!visualRecord.revealed);
    } else if (visualRecord && visualRecord.type === "shape-cover") {
      effectiveBgColor = getSolidColorFromName(newColorName, newBgColor);
    }

    const spans = getAnnotationElements(id);
    for (const span of spans) {
      span.setAttribute(HIGHLIGHT_COLOR_ATTR, newColorName);
      if ((span.getAttribute(HIGHLIGHT_KIND_ATTR) || "").indexOf("shape-") === 0) {
        if (visualRecord) {
          visualRecord.color = newColorName;
          visualRecord.bgColor = effectiveBgColor;
          applyShapeVisual(span, visualRecord, effectiveBgColor);
        }
      } else {
        span.style.setProperty("background-color", effectiveBgColor, "important");
        span.style.removeProperty("color");
      }
    }
    if (previewOnly) return;

    if (record && !isShapeRecord(record)) {
      record.color = newColorName;
      record.bgColor = effectiveBgColor;
    } else if (record && visualRecord && isShapeRecord(record)) {
      record.color = visualRecord.color;
      record.bgColor = visualRecord.bgColor;
      if (record.type === "shape-rect") {
        record.revealed = !!visualRecord.revealed;
      }
    }

    if (record && record.type === "shape-rect" && record.transient) {
      record.transient = false;
      sendRuntimeMessageSafe({ action: "saveHighlight", highlight: record });
      return;
    }

    if (!record || !record.transient) {
      sendRuntimeMessageSafe({
        action: "updateHighlightColor",
        id: id,
        domain: DOMAIN,
        color: newColorName,
        bgColor: effectiveBgColor
      });
    }
  }

  function formatNotePreview(note) {
    const text = String(note || "").replace(/\s+/g, " ").trim();
    if (!text) return "";
    if (text.length <= 80) return text;
    return text.slice(0, 77).trimEnd() + "...";
  }

  function saveHighlightComment(highlightId, nextNoteRaw, options) {
    const opts = options && typeof options === "object" ? options : {};
    const record = storedHighlightsForPage.find((x) => idsMatch(x && x.id, highlightId));
    if (!record || record.transient) return false;
    const existingNote = typeof record.note === "string" ? record.note : "";
    const nextNote = String(nextNoteRaw || "");
    if (nextNote === existingNote) return true;
    record.note = nextNote;
    const noteEls = getAnnotationElements(highlightId);
    for (const el of noteEls) {
      applyNoteTooltip(el, nextNote);
    }
    sendRuntimeMessageSafe({
      action: "saveHighlight",
      highlight: record,
      historyMeta: {
        action: "update",
        label: nextNote ? "Updated highlight comment" : "Cleared highlight comment"
      }
    });
    if (!opts.silentToast) {
      showInlineUndoToast(nextNote ? "Comment saved" : "Comment cleared", null, 2400);
    }
    return true;
  }

  function applyDefaultHighlightForNoteFlow() {
    if (!currentSelection || !currentSelection.range) return false;
    const defaultColorName = "yellow";
    const defaultColor = PRESET_COLORS[defaultColorName] || PRESET_COLORS.yellow;
    lastUsedColor = defaultColorName;
    lastUsedBgColor = defaultColor.bg;
    const createdId = createHighlightFromSelection(defaultColorName, defaultColor.bg, pickerNoteDraft);
    if (!createdId) return false;
    const target = getPrimaryAnnotationElement(createdId);
    if (!target) return true;
    hoverNoteEditorId = createdId;
    hoverNoteDraft = "";
    const targetRect = target.getBoundingClientRect();
    showHoverToolbar(createdId, targetRect);
    return true;
  }

  function createHighlightFromSelection(colorName, bgColor, noteOverride) {
    let range = null;
    let text = "";
    let noteText = "";
    const sel = window.getSelection();

    if (currentSelection && currentSelection.range) {
      range = currentSelection.range.cloneRange();
      text = String(currentSelection.text || "").trim();
      noteText = typeof currentSelection.note === "string" ? currentSelection.note : "";
      if (!text) {
        text = String(range.toString() || "").trim();
      }
    } else {
      if (!sel || sel.isCollapsed || !sel.toString().trim() || !sel.rangeCount) return "";
      range = sel.getRangeAt(0).cloneRange();
      text = String(sel.toString() || "").trim();
    }

    if (!range || !text) return "";
    const id = "hm-" + Date.now() + "-" + Math.random().toString(36).slice(2, 8);

    try {
      removeExistingTextHighlightsForRange(range);
      applyHighlight(range, id, colorName, bgColor);
      const record = buildHighlightRecord(id, colorName, bgColor, text, range);
      if (typeof noteOverride === "string") {
        noteText = noteOverride;
      }
      record.note = String(noteText || "");
      storedHighlightsForPage.push(record);
      const createdEls = getAnnotationElements(id);
      for (const el of createdEls) {
        applyNoteTooltip(el, record.note || "");
      }
      sendRuntimeMessageSafe({ action: "saveHighlight", highlight: record });
      if (sel && sel.removeAllRanges) {
        try { sel.removeAllRanges(); } catch (_e) {}
      }
      colorPickerOpen = false;
      hidePicker();
      hideHoverToolbar();
      currentSelection = null;
      selectionDetected = false;
      setTimeout(() => pulseHighlight(id), 50);
      return id;
    } catch (err) {
      console.warn("[HighlightMaster] Could not create highlight:", err);
      return "";
    }
  }

  function clearInlineUndoToast() {
    if (inlineUndoTimer) {
      clearTimeout(inlineUndoTimer);
      inlineUndoTimer = null;
    }
    if (inlineUndoEl && inlineUndoEl.parentNode) {
      inlineUndoEl.parentNode.removeChild(inlineUndoEl);
    }
    inlineUndoEl = null;
  }

  function showInlineUndoToast(message, onUndo, durationMs) {
    clearInlineUndoToast();
    const toast = document.createElement("div");
    toast.className = "hm-inline-undo";
    const label = document.createElement("span");
    label.textContent = message || "Deleted";
    toast.appendChild(label);
    if (typeof onUndo === "function") {
      const btn = document.createElement("button");
      btn.textContent = "Undo";
      btn.addEventListener("click", (e) => {
        e.preventDefault();
        e.stopPropagation();
        clearInlineUndoToast();
        try { onUndo(); } catch (_e) {}
      });
      toast.appendChild(btn);
    }
    (document.body || document.documentElement).appendChild(toast);
    inlineUndoEl = toast;
    inlineUndoTimer = setTimeout(clearInlineUndoToast, Math.max(1200, durationMs || 10000));
  }

  function deleteHighlightsWithUndo(ids) {
    const uniqueIds = Array.from(new Set((ids || []).map((id) => normalizeHighlightIdValue(id)).filter(Boolean)));
    if (!uniqueIds.length) return;

    const targets = [];
    for (const id of uniqueIds) {
      const record = storedHighlightsForPage.find((x) => idsMatch(x && x.id, id));
      removeHighlightFromPage(id);
      if (record && !record.transient) {
        targets.push({ domain: record.domain || DOMAIN, id: normalizeHighlightIdValue(record.id) });
      } else {
        // Fallback by id+current domain so storage can still resolve deletion.
        targets.push({ domain: DOMAIN, id });
      }
    }

    if (!targets.length) {
      showInlineUndoToast("Deleted", null, 4000);
      return;
    }

    const message = targets.length === 1 ? "Deleted 1 highlight" : ("Deleted " + targets.length + " highlights");
    sendRuntimeMessageSafe({
      action: "bulkDeleteHighlights",
      targets,
      historyAction: "bulk_delete",
      historyLabel: message,
      pageUrl: location.href,
      domain: DOMAIN
    }, (res, errorMessage) => {
      const removedCount = !errorMessage && res && typeof res.removed === "number" ? res.removed : 0;
      const historyId = !errorMessage && res && res.ok && removedCount > 0 ? res.historyId : null;
      if (!historyId) {
        let fallbackRemoved = 0;
        let remaining = targets.length;
        if (!remaining) {
          showInlineUndoToast(message, null, 5000);
          return;
        }
        const finalizeFallback = () => {
          remaining--;
          if (remaining > 0) return;
          if (fallbackRemoved > 0) {
            showInlineUndoToast(message, null, 5000);
            return;
          }
          setTimeout(() => {
            restoredIds.clear();
            storedHighlightsForPage = [];
            restoreAttempts = 0;
            restoreHighlights();
          }, 120);
          showInlineUndoToast(message, null, 5000);
        };
        for (const target of targets) {
          sendRuntimeMessageSafe({
            action: "deleteHighlight",
            domain: target.domain,
            id: target.id
          }, (singleRes, singleErrorMessage) => {
            const singleRemoved = !singleErrorMessage && singleRes && typeof singleRes.removed === "number" ? singleRes.removed : 0;
            if (!singleErrorMessage && singleRes && singleRes.ok && singleRemoved > 0) {
              fallbackRemoved += singleRemoved;
            }
            finalizeFallback();
          });
        }
        return;
      }
      showInlineUndoToast(message, () => {
        sendRuntimeMessageSafe({ action: "undoTimelineEntry", entryId: historyId }, (undoRes, undoErrorMessage) => {
          if (!undoErrorMessage && undoRes && undoRes.ok) {
            activeResize = null;
            setRectangleResizeCursor("se", false);
            activeMove = null;
            setRectangleMoveCursor(false);
            restoredIds.clear();
            storedHighlightsForPage = [];
            restoreAttempts = 0;
            restoreHighlights();
          }
        });
      }, 10000);
    });
  }

  function setHoverStyle(id, apply) {
    if (!id) return;
    const spans = getAnnotationElements(id);
    for (const span of spans) {
      if (apply) span.classList.add("hm-hovered");
      else span.classList.remove("hm-hovered");
    }
  }

  function getActiveRectangleTargetId() {
    if (activeResize && activeResize.id) return activeResize.id;
    if (activeMove && activeMove.id) return activeMove.id;
    if (activeHoverId) {
      const hoveredRect = getPrimaryAnnotationElement(activeHoverId);
      if (hoveredRect && hoveredRect.getAttribute(HIGHLIGHT_KIND_ATTR) === "shape-rect") {
        return activeHoverId;
      }
    }
    const hoveredRect = document.querySelector("div[" + HIGHLIGHT_ATTR + "][data-hm-kind='shape-rect'].hm-hovered");
    return hoveredRect ? (hoveredRect.getAttribute(HIGHLIGHT_ATTR) || "") : "";
  }

  function isEditableShortcutTarget(event) {
    const path = event && event.composedPath ? event.composedPath() : [];
    for (const node of path) {
      if (!node || !node.closest) continue;
      if (node.closest('textarea, input, select, [contenteditable="true"], [contenteditable=""], [contenteditable="plaintext-only"]')) {
        return true;
      }
    }
    const activeEl = document.activeElement;
    return !!(activeEl && activeEl.closest && activeEl.closest('textarea, input, select, [contenteditable="true"], [contenteditable=""], [contenteditable="plaintext-only"]'));
  }

  // Hover event listeners on highlights
  document.addEventListener("mouseover", (e) => {
    if (!extensionEnabled) return;
    if (activeResize) return;
    if (activeMove) return;
    const target = e.target;
    if (!target) return;
    const owner = target.closest ? target.closest("[" + HIGHLIGHT_ATTR + "]") : null;
    const hmId = owner && owner.getAttribute ? owner.getAttribute(HIGHLIGHT_ATTR) : (target.getAttribute ? target.getAttribute(HIGHLIGHT_ATTR) : "");
    if (!hmId) return;

    if (activeHoverId && activeHoverId !== hmId) {
      setHoverStyle(activeHoverId, false);
    }
    
    activeHoverId = hmId;
    setHoverStyle(hmId, true);
  }, true);

  document.addEventListener("mouseout", (e) => {
    if (activeResize) return;
    if (activeMove) return;
    const target = e.target;
    if (!target) return;
    const owner = target.closest ? target.closest("[" + HIGHLIGHT_ATTR + "]") : null;
    const hmId = owner && owner.getAttribute ? owner.getAttribute(HIGHLIGHT_ATTR) : (target.getAttribute ? target.getAttribute(HIGHLIGHT_ATTR) : "");
    if (!hmId) return;

    setHoverStyle(hmId, false);
    if (activeHoverId === hmId) {
      activeHoverId = null;
    }
  }, true);

  document.addEventListener("contextmenu", (e) => {
    if (!extensionEnabled) return;
    if (activeResize || activeMove) return;
    if (isEventFromOcrCopyChip(e)) {
      hideTextHighlightContextMenu();
      return;
    }
    const owner = getTextHighlightOwnerFromEvent(e);
    if (!owner) {
      hideTextHighlightContextMenu();
      return;
    }
    const hmId = owner.getAttribute(HIGHLIGHT_ATTR) || "";
    if (!hmId) return;
    e.preventDefault();
    e.stopPropagation();
    e.stopImmediatePropagation();
    hidePicker();
    hideHoverToolbar();
    showTextHighlightContextMenu(hmId, e.pageX, e.pageY);
  }, true);

  // Floating toolbar opens on click only, not hover.
  document.addEventListener("click", (e) => {
    if (!extensionEnabled) return;
    if (activeResize || activeMove) return;
    if (isEventFromOcrCopyChip(e)) return;
    if (textHighlightContextMenuEl) {
      const targetEl = getEventTargetElement(e);
      if (targetEl && textHighlightContextMenuEl.contains(targetEl)) return;
      hideTextHighlightContextMenu();
    }
    const path = e.composedPath ? e.composedPath() : [];
    if (path.includes(hoverHost)) return;
    const target = e.target;
    if (!target) return;
    const owner = target.closest ? target.closest("[" + HIGHLIGHT_ATTR + "]") : null;
    const hmId = owner && owner.getAttribute ? owner.getAttribute(HIGHLIGHT_ATTR) : "";
    if (!hmId) return;
    e.preventDefault();
    e.stopPropagation();
    hidePicker();
    activeHoverTarget = owner;
    const rect = owner.getBoundingClientRect();
    showHoverToolbar(hmId, rect);
  }, true);

  function clearDrawPreview() {
    if (drawPreviewEl && drawPreviewEl.parentNode) {
      drawPreviewEl.parentNode.removeChild(drawPreviewEl);
    }
    drawPreviewEl = null;
    drawStartPoint = null;
    drawStartSourceNode = null;
    drawMoved = false;
  }

  function applyRectToShapeElement(el, rect) {
    if (!el || !rect) return;
    el.style.left = rect.left + "px";
    el.style.top = rect.top + "px";
    el.style.width = rect.width + "px";
    el.style.height = rect.height + "px";
  }

  function revertRectangleInteraction(state, activeClassName) {
    if (!state) return;
    if (state.record && state.startRect) {
      state.record.rect = {
        left: state.startRect.left,
        top: state.startRect.top,
        width: state.startRect.width,
        height: state.startRect.height
      };
    }
    const renderedRect = getRenderedRectForRecord(state.record) || state.startRect;
    applyRectToShapeElement(state.el, renderedRect);
    if (state.el) {
      state.el.classList.remove(activeClassName, "hm-hovered");
    }
  }

  function cancelInteractionState(clearDrawMode) {
    clearDrawPreview();
    if (activeResize) {
      revertRectangleInteraction(activeResize, "hm-resizing");
    }
    activeResize = null;
    setRectangleResizeCursor("se", false);
    if (activeMove) {
      revertRectangleInteraction(activeMove, "hm-moving");
    }
    activeMove = null;
    setRectangleMoveCursor(false);
    if (clearDrawMode) {
      drawMode = null;
    }
    updateDrawDockState();
  }

  function setDrawMode(nextMode) {
    const normalizedMode = nextMode === "shape-rect" ? nextMode : null;
    if (normalizedMode !== drawMode) {
      clearDrawPreview();
    }
    drawMode = extensionEnabled ? normalizedMode : null;
    updateDrawDockState();
    return drawMode;
  }

  function updateDrawDockState() {
    const drawActive = extensionEnabled && !!drawMode;
    document.documentElement.classList.toggle("hm-draw-cursor", drawActive);
    if (document.body) {
      document.body.classList.toggle("hm-draw-cursor", drawActive);
    }
  }

  function isEventInsideHmUi(event) {
    const path = event.composedPath ? event.composedPath() : [];
    return path.includes(pickerHost) || path.includes(hoverHost);
  }

  function rectsIntersect(a, b) {
    if (!a || !b) return false;
    return !(b.left >= a.right || b.right <= a.left || b.top >= a.bottom || b.bottom <= a.top);
  }

  function isRectContained(inner, outer) {
    if (!inner || !outer) return false;
    const innerRight = inner.left + inner.width;
    const innerBottom = inner.top + inner.height;
    const outerRight = outer.left + outer.width;
    const outerBottom = outer.top + outer.height;
    return inner.left >= outer.left
      && inner.top >= outer.top
      && innerRight <= outerRight
      && innerBottom <= outerBottom;
  }

  function isPointInsideRect(pageX, pageY, rect) {
    if (!rect) return false;
    return pageX >= rect.left
      && pageX <= rect.left + rect.width
      && pageY >= rect.top
      && pageY <= rect.top + rect.height;
  }

  function isInsideSolidRectangleByPoint(pageX, pageY) {
    for (const record of storedHighlightsForPage) {
      if (!record || record.type !== "shape-rect" || !record.rect) continue;
      normalizeShapeRecord(record, false);
      if (record.revealed) continue;
      const renderedRect = getRenderedRectForRecord(record) || record.rect;
      if (isPointInsideRect(pageX, pageY, renderedRect)) return true;
    }
    return false;
  }

  function isNestedInsideSolidRectangle(rect) {
    for (const record of storedHighlightsForPage) {
      if (!record || record.type !== "shape-rect" || !record.rect) continue;
      normalizeShapeRecord(record, false);
      if (record.revealed) continue;
      const renderedRect = getRenderedRectForRecord(record) || record.rect;
      if (isRectContained(rect, renderedRect)) return true;
    }
    return false;
  }

  function normalizeSnippetText(text) {
    return (text || "").replace(/[\u200B-\u200D\uFEFF]/g, "").replace(/\s+/g, " ").trim();
  }

  function pushWords(targetWords, text, maxWords) {
    const clean = normalizeSnippetText(text);
    if (!clean) return;
    const words = clean.split(" ");
    for (const word of words) {
      targetWords.push(word);
      if (targetWords.length >= maxWords) break;
    }
  }

  function isVisibleElementForCapture(el) {
    if (!el || el.nodeType !== 1) return false;
    const tag = (el.tagName || "").toLowerCase();
    if (tag === "script" || tag === "style" || tag === "noscript") return false;
    const style = window.getComputedStyle(el);
    if (!style) return true;
    if (style.display === "none" || style.visibility === "hidden") return false;
    if (parseFloat(style.opacity || "1") <= 0) return false;
    return true;
  }

  function extractWordsFromTextNodes(viewRect, maxWords) {
    if (!document.body) return [];
    const words = [];
    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, {
      acceptNode(node) {
        if (!node || !node.nodeValue || !node.nodeValue.trim()) return NodeFilter.FILTER_REJECT;
        const parent = node.parentElement;
        if (!parent || !isVisibleElementForCapture(parent)) return NodeFilter.FILTER_REJECT;
        return NodeFilter.FILTER_ACCEPT;
      }
    });

    let node;
    while ((node = walker.nextNode())) {
      if (words.length >= maxWords) break;
      try {
        const range = document.createRange();
        range.selectNodeContents(node);
        const rect = range.getBoundingClientRect();
        if (!rect || (!rect.width && !rect.height)) continue;
        if (!rectsIntersect(viewRect, rect)) continue;
        pushWords(words, node.nodeValue, maxWords);
      } catch (_e) {}
    }

    return words;
  }

  function extractWordsFromElementsFallback(viewRect, maxWords) {
    const words = [];
    const seenTexts = new Set();
    const stepX = Math.max(18, Math.floor(viewRect.width / 6));
    const stepY = Math.max(18, Math.floor(viewRect.height / 6));

    for (let y = viewRect.top + 2; y < viewRect.bottom && words.length < maxWords; y += stepY) {
      for (let x = viewRect.left + 2; x < viewRect.right && words.length < maxWords; x += stepX) {
        let elements = [];
        try {
          elements = document.elementsFromPoint(x, y) || [];
        } catch (_e) {
          elements = [];
        }
        for (const el of elements) {
          if (!isVisibleElementForCapture(el)) continue;
          const text = normalizeSnippetText(el.innerText || el.textContent || "");
          if (!text || seenTexts.has(text)) continue;
          seenTexts.add(text);
          pushWords(words, text, maxWords);
          break;
        }
      }
    }

    return words;
  }

  function extractRectangleWords(rect, maxWords) {
    if (!rect) return [];
    const wordLimit = Math.max(12, Number(maxWords) || RECT_OCR_WORD_LIMIT);
    const viewRect = {
      left: rect.left - window.scrollX,
      top: rect.top - window.scrollY,
      right: rect.left - window.scrollX + rect.width,
      bottom: rect.top - window.scrollY + rect.height,
      width: rect.width,
      height: rect.height
    };
    if (viewRect.width < 4 || viewRect.height < 4) return [];

    const words = extractWordsFromTextNodes(viewRect, wordLimit);
    if (words.length < 6) {
      const fallbackWords = extractWordsFromElementsFallback(viewRect, wordLimit);
      for (const word of fallbackWords) {
        words.push(word);
        if (words.length >= wordLimit) break;
      }
    }
    return words;
  }

  function extractRectangleSnippet(rect) {
    const words = extractRectangleWords(rect, RECT_OCR_WORD_LIMIT);
    return normalizeSnippetText(words.slice(0, RECT_OCR_WORD_LIMIT).join(" "));
  }

  function extractRectangleCopyText(rect) {
    const words = extractRectangleWords(rect, RECT_OCR_COPY_WORD_LIMIT);
    return normalizeSnippetText(words.join(" "));
  }

  function truncateLabelWords(text, maxWords) {
    const clean = normalizeSnippetText(text || "");
    if (!clean) return "";
    const words = clean.split(" ");
    return words.slice(0, Math.max(1, maxWords || 10)).join(" ");
  }

  function findNearbyHeadingLabel(rect) {
    if (!rect) return "";
    const viewRect = {
      left: rect.left - window.scrollX,
      top: rect.top - window.scrollY,
      right: rect.left - window.scrollX + rect.width,
      bottom: rect.top - window.scrollY + rect.height,
      width: rect.width,
      height: rect.height
    };
    const candidates = document.querySelectorAll("h1, h2, h3, h4, h5, h6, [role='heading']");
    let bestText = "";
    let bestScore = -Infinity;

    for (const el of candidates) {
      if (!isVisibleElementForCapture(el)) continue;
      const text = truncateLabelWords(el.innerText || el.textContent || "", 12);
      if (!text) continue;

      const headingRect = el.getBoundingClientRect();
      if (!headingRect || (!headingRect.width && !headingRect.height)) continue;

      const horizontalOverlap = Math.max(0, Math.min(headingRect.right, viewRect.right) - Math.max(headingRect.left, viewRect.left));
      const overlapRatio = horizontalOverlap / Math.max(1, Math.min(Math.max(headingRect.width, 1), Math.max(viewRect.width, 1)));
      const centerDelta = Math.abs((headingRect.left + headingRect.width / 2) - (viewRect.left + viewRect.width / 2));
      const aboveGap = viewRect.top - headingRect.bottom;
      const belowGap = headingRect.top - viewRect.bottom;

      if (aboveGap > 280) continue;
      if (belowGap > 90) continue;

      let score = 220;
      score -= Math.abs(aboveGap >= 0 ? aboveGap : aboveGap * 1.8);
      if (belowGap > 0) score -= belowGap * 1.2;
      score += overlapRatio * 180;
      score -= centerDelta * 0.22;

      const tag = (el.tagName || "").toLowerCase();
      if (tag === "h1") score += 45;
      else if (tag === "h2") score += 35;
      else if (tag === "h3") score += 24;

      if (score > bestScore) {
        bestScore = score;
        bestText = text;
      }
    }

    if (bestScore < 40) return "";
    return bestText;
  }

  function getSmartRectangleLabel(rect, snippetText) {
    const headingLabel = findNearbyHeadingLabel(rect);
    if (headingLabel) return headingLabel;
    const snippetLabel = truncateLabelWords(snippetText, 10);
    if (snippetLabel) return snippetLabel;
    const titleLabel = truncateLabelWords(document.title || "", 8);
    if (titleLabel) return "Area in " + titleLabel;
    return "Rectangle Highlight";
  }

  function ensureRectangleResizeHandles(el) {
    if (!el) return;
    if (el.querySelector(".hm-resize-handle")) return;
    const corners = ["nw", "ne", "sw", "se"];
    for (const corner of corners) {
      const handle = document.createElement("div");
      handle.className = "hm-resize-handle";
      handle.setAttribute("data-hm-corner", corner);
      handle.setAttribute("title", "Drag to resize");
      el.appendChild(handle);
    }
  }

  function setRectangleResizeCursor(corner, enabled) {
    const roots = [document.documentElement, document.body].filter(Boolean);
    const classNwse = "hm-rect-resize-nwse";
    const classNesw = "hm-rect-resize-nesw";
    for (const root of roots) {
      root.classList.remove(classNwse, classNesw);
    }
    if (!enabled) return;
    const className = corner === "ne" || corner === "sw" ? classNesw : classNwse;
    for (const root of roots) {
      root.classList.add(className);
    }
  }

  function setRectangleMoveCursor(enabled) {
    const roots = [document.documentElement, document.body].filter(Boolean);
    const className = "hm-rect-move";
    for (const root of roots) {
      if (enabled) root.classList.add(className);
      else root.classList.remove(className);
    }
  }

  function startRectangleResize(event) {
    if (!extensionEnabled) return false;
    if (!event || !event.target || !event.target.closest) return false;

    const handle = event.target.closest(".hm-resize-handle");
    if (!handle) return false;

    const shapeEl = handle.closest("div[" + HIGHLIGHT_ATTR + "][data-hm-kind='shape-rect']");
    if (!shapeEl) return false;

    const id = shapeEl.getAttribute(HIGHLIGHT_ATTR);
    const record = storedHighlightsForPage.find((x) => idsMatch(x && x.id, id) && x.type === "shape-rect");
    if (!record) return false;

    const currentRect = record.rect ? {
      left: Number(record.rect.left) || 0,
      top: Number(record.rect.top) || 0,
      width: Math.max(RECT_MIN_SIZE, Number(record.rect.width) || 0),
      height: Math.max(RECT_MIN_SIZE, Number(record.rect.height) || 0)
    } : {
      left: parseFloat(shapeEl.style.left) || 0,
      top: parseFloat(shapeEl.style.top) || 0,
      width: Math.max(RECT_MIN_SIZE, parseFloat(shapeEl.style.width) || 0),
      height: Math.max(RECT_MIN_SIZE, parseFloat(shapeEl.style.height) || 0)
    };

    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation();
    hidePicker();
    hideHoverToolbar();

    const corner = handle.getAttribute("data-hm-corner") || "se";
    activeResize = {
      id,
      corner,
      el: shapeEl,
      record,
      startX: event.pageX,
      startY: event.pageY,
      startRect: currentRect,
      didMove: false
    };
    shapeEl.classList.add("hm-resizing", "hm-hovered");
    setRectangleResizeCursor(corner, true);
    return true;
  }

  function startRectangleMove(event) {
    if (!extensionEnabled) return false;
    if (!event) return false;
    if (event.button !== 0) return false;

    const targetEl = getEventTargetElement(event);
    if (!targetEl || !targetEl.closest) return false;

    if (targetEl.closest(".hm-resize-handle")) return false;
    if (targetEl.closest(".hm-ocr-copy-chip")) return false;

    const shapeEl = targetEl.closest("div[" + HIGHLIGHT_ATTR + "][data-hm-kind='shape-rect']");
    if (!shapeEl) return false;

    const id = shapeEl.getAttribute(HIGHLIGHT_ATTR);
    const record = storedHighlightsForPage.find((x) => idsMatch(x && x.id, id) && x.type === "shape-rect");
    if (!record) return false;

    const currentRect = record.rect ? {
      left: Number(record.rect.left) || 0,
      top: Number(record.rect.top) || 0,
      width: Math.max(RECT_MIN_SIZE, Number(record.rect.width) || 0),
      height: Math.max(RECT_MIN_SIZE, Number(record.rect.height) || 0)
    } : {
      left: parseFloat(shapeEl.style.left) || 0,
      top: parseFloat(shapeEl.style.top) || 0,
      width: Math.max(RECT_MIN_SIZE, parseFloat(shapeEl.style.width) || 0),
      height: Math.max(RECT_MIN_SIZE, parseFloat(shapeEl.style.height) || 0)
    };

    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation();
    hidePicker();
    hideHoverToolbar();

    activeMove = {
      id,
      el: shapeEl,
      record,
      startX: event.pageX,
      startY: event.pageY,
      startRect: currentRect,
      didMove: false
    };
    shapeEl.classList.add("hm-moving", "hm-hovered");
    setRectangleMoveCursor(true);
    return true;
  }

  function updateRectangleResize(event) {
    if (!activeResize) return;
    const resizeState = activeResize;
    const corner = resizeState.corner;
    const dx = event.pageX - resizeState.startX;
    const dy = event.pageY - resizeState.startY;

    let left = resizeState.startRect.left;
    let top = resizeState.startRect.top;
    let right = resizeState.startRect.left + resizeState.startRect.width;
    let bottom = resizeState.startRect.top + resizeState.startRect.height;

    if (corner.indexOf("w") !== -1) {
      left = Math.min(resizeState.startRect.left + dx, right - RECT_MIN_SIZE);
    }
    if (corner.indexOf("e") !== -1) {
      right = Math.max(resizeState.startRect.left + resizeState.startRect.width + dx, left + RECT_MIN_SIZE);
    }
    if (corner.indexOf("n") !== -1) {
      top = Math.min(resizeState.startRect.top + dy, bottom - RECT_MIN_SIZE);
    }
    if (corner.indexOf("s") !== -1) {
      bottom = Math.max(resizeState.startRect.top + resizeState.startRect.height + dy, top + RECT_MIN_SIZE);
    }

    const nextRect = {
      left,
      top,
      width: Math.max(RECT_MIN_SIZE, right - left),
      height: Math.max(RECT_MIN_SIZE, bottom - top)
    };

    resizeState.didMove = resizeState.didMove
      || Math.abs(nextRect.left - resizeState.startRect.left) > 0.8
      || Math.abs(nextRect.top - resizeState.startRect.top) > 0.8
      || Math.abs(nextRect.width - resizeState.startRect.width) > 0.8
      || Math.abs(nextRect.height - resizeState.startRect.height) > 0.8;

    resizeState.record.rect = nextRect;
    const renderedRect = getRenderedRectForRecord(resizeState.record) || nextRect;
    resizeState.el.style.left = renderedRect.left + "px";
    resizeState.el.style.top = renderedRect.top + "px";
    resizeState.el.style.width = renderedRect.width + "px";
    resizeState.el.style.height = renderedRect.height + "px";
  }

  function updateRectangleMove(event) {
    if (!activeMove) return;
    const moveState = activeMove;
    const dx = event.pageX - moveState.startX;
    const dy = event.pageY - moveState.startY;

    const maxPageX = Math.max(
      document.documentElement ? document.documentElement.scrollWidth : 0,
      document.body ? document.body.scrollWidth : 0,
      window.scrollX + window.innerWidth
    );
    const maxPageY = Math.max(
      document.documentElement ? document.documentElement.scrollHeight : 0,
      document.body ? document.body.scrollHeight : 0,
      window.scrollY + window.innerHeight
    );

    const unclampedLeft = moveState.startRect.left + dx;
    const unclampedTop = moveState.startRect.top + dy;
    const left = Math.max(0, Math.min(unclampedLeft, Math.max(0, maxPageX - moveState.startRect.width)));
    const top = Math.max(0, Math.min(unclampedTop, Math.max(0, maxPageY - moveState.startRect.height)));

    moveState.didMove = moveState.didMove
      || Math.abs(left - moveState.startRect.left) > 0.8
      || Math.abs(top - moveState.startRect.top) > 0.8;

    const nextRect = {
      left,
      top,
      width: moveState.startRect.width,
      height: moveState.startRect.height
    };
    moveState.record.rect = nextRect;
    const renderedRect = getRenderedRectForRecord(moveState.record) || nextRect;
    moveState.el.style.left = renderedRect.left + "px";
    moveState.el.style.top = renderedRect.top + "px";
  }

  function finishRectangleResize() {
    if (!activeResize) return;
    const resizeState = activeResize;
    activeResize = null;
    resizeState.el.classList.remove("hm-resizing", "hm-hovered");
    setRectangleResizeCursor(resizeState.corner, false);

    if (!resizeState.didMove) return;

    const nextRect = resizeState.record.rect;
    if (!nextRect || nextRect.width < 10 || nextRect.height < 10) return;

    const visualRect = getRenderedRectForRecord(resizeState.record) || nextRect;
    const snippet = extractRectangleSnippet(visualRect);
    if (snippet) {
      resizeState.record.ocrText = snippet;
      resizeState.record.text = snippet;
    }
    if (resizeState.record.revealed) {
      const fullText = extractRectangleCopyText(visualRect);
      if (fullText) resizeState.record.ocrFullText = fullText;
    }
    resizeState.record.label = getSmartRectangleLabel(visualRect, snippet || resizeState.record.ocrText || "");
    const resizeEls = getAnnotationElements(resizeState.id);
    for (const el of resizeEls) {
      ensureRectangleOcrCopyChip(el, resizeState.record);
    }

    sendRuntimeMessageSafe({
      action: "saveHighlight",
      highlight: resizeState.record,
      historyMeta: {
        action: "resize_rect",
        label: "Resized rectangle"
      }
    });

    if (activeHoverId === resizeState.id) {
      refreshHoverToolbar(resizeState.id);
    }
  }

  function finishRectangleMove() {
    if (!activeMove) return;
    const moveState = activeMove;
    activeMove = null;
    moveState.el.classList.remove("hm-moving", "hm-hovered");
    setRectangleMoveCursor(false);

    if (!moveState.didMove) {
      const targetEl = getPrimaryAnnotationElement(moveState.id) || moveState.el;
      if (targetEl) {
        const rect = targetEl.getBoundingClientRect();
        showHoverToolbar(moveState.id, rect);
      }
      return;
    }

    const nextRect = moveState.record.rect;
    if (!nextRect || nextRect.width < 10 || nextRect.height < 10) return;

    const visualRect = getRenderedRectForRecord(moveState.record) || nextRect;
    const snippet = extractRectangleSnippet(visualRect);
    if (snippet) {
      moveState.record.ocrText = snippet;
      moveState.record.text = snippet;
    }
    if (moveState.record.revealed) {
      const fullText = extractRectangleCopyText(visualRect);
      if (fullText) moveState.record.ocrFullText = fullText;
    }
    moveState.record.label = getSmartRectangleLabel(visualRect, snippet || moveState.record.ocrText || "");
    const moveEls = getAnnotationElements(moveState.id);
    for (const el of moveEls) {
      ensureRectangleOcrCopyChip(el, moveState.record);
    }

    sendRuntimeMessageSafe({
      action: "saveHighlight",
      highlight: moveState.record,
      historyMeta: {
        action: "move_rect",
        label: "Moved rectangle"
      }
    });

    if (activeHoverId === moveState.id) {
      refreshHoverToolbar(moveState.id);
    }
  }

  function createShapeElement(record) {
    normalizeShapeRecord(record, false);
    inferShapeScrollAnchorIfMissing(record);
    const rect = record.rect;
    if (!rect) return null;
    const el = document.createElement("div");
    el.setAttribute(HIGHLIGHT_ATTR, record.id);
    el.setAttribute(HIGHLIGHT_COLOR_ATTR, record.color);
    el.setAttribute(HIGHLIGHT_KIND_ATTR, record.type);
    el.classList.add("hm-pop-anim");
    const renderRect = getRenderedRectForRecord(record) || rect;
    applyRectToShapeElement(el, renderRect);
    if (record.type === "shape-rect") {
      ensureRectangleResizeHandles(el);
    }
    applyShapeVisual(el, record, record.bgColor);
    applyNoteTooltip(el, record.note || "");
    (document.body || document.documentElement).appendChild(el);
    return el;
  }

  function buildShapeRecord(id, type, colorName, bgColor, rect, ocrText, scrollAnchor) {
    const cleanOcr = normalizeSnippetText(ocrText || "");
    const isRect = type === "shape-rect";
    const isCover = type === "shape-cover";
    const smartLabel = isRect ? getSmartRectangleLabel(rect, cleanOcr) : "";
    const normalizedColor = isRect ? "yellow" : (colorName || "yellow");
    const startRevealed = false;
    const normalizedBg = isRect
      ? getRectBgForState(normalizedColor, PRESET_COLORS.yellow.solid, startRevealed)
      : getSolidColorFromName(normalizedColor, bgColor || PRESET_COLORS.yellow.solid);
    return {
      id,
      type,
      domain: DOMAIN,
      url: location.href,
      text: isCover ? "Cover Rectangle" : (cleanOcr || smartLabel || "Rectangle Highlight"),
      note: "",
      label: isRect ? (smartLabel || cleanOcr || "Rectangle Highlight") : "",
      ocrText: isCover ? "" : cleanOcr,
      ocrFullText: isCover ? "" : cleanOcr,
      color: normalizedColor,
      bgColor: normalizedBg,
      timestamp: Date.now(),
      pageTitle: document.title || "",
      revealed: isRect ? startRevealed : false,
      transient: isRect ? false : isCover,
      scrollHostXPath: scrollAnchor && scrollAnchor.scrollHostXPath ? scrollAnchor.scrollHostXPath : "",
      scrollHostTop: scrollAnchor && typeof scrollAnchor.scrollHostTop === "number" ? scrollAnchor.scrollHostTop : 0,
      scrollHostLeft: scrollAnchor && typeof scrollAnchor.scrollHostLeft === "number" ? scrollAnchor.scrollHostLeft : 0,
      rect
    };
  }

  function startShapeDraw(event) {
    if (!drawMode || !extensionEnabled || isEventInsideHmUi(event)) return;
    if (isEventFromOcrCopyChip(event)) return;
    if (event.button !== 0) return;
    if (drawMode === "shape-rect" && isInsideSolidRectangleByPoint(event.pageX, event.pageY)) return;
    const target = event.target;
    if (target && target.closest && target.closest('input, textarea, select, button, a, [contenteditable="true"], [contenteditable=""], [contenteditable="plaintext-only"]')) {
      return;
    }
    event.preventDefault();
    hidePicker();
    hideHoverToolbar();
    window.getSelection().removeAllRanges();
    const pageX = event.pageX;
    const pageY = event.pageY;
    drawStartPoint = { x: pageX, y: pageY };
    drawStartSourceNode = event.target || null;
    drawMoved = false;
    drawPreviewEl = document.createElement("div");
    drawPreviewEl.className = "hm-draw-preview";
    drawPreviewEl.setAttribute(HIGHLIGHT_KIND_ATTR, drawMode);
    drawPreviewEl.style.position = "absolute";
    drawPreviewEl.style.zIndex = "2147483646";
    drawPreviewEl.style.pointerEvents = "none";
    drawPreviewEl.style.boxSizing = "border-box";
    const previewColorName = drawMode === "shape-rect" ? "yellow" : lastUsedColor;
    const previewBgColor = drawMode === "shape-rect"
      ? PRESET_COLORS.yellow.solid
      : lastUsedBgColor;
    const previewRecord = normalizeShapeRecord({
      type: drawMode,
      color: previewColorName,
      bgColor: previewBgColor,
      revealed: false,
      transient: true
    }, false);
    applyShapeVisual(drawPreviewEl, previewRecord, previewBgColor);
    if (drawMode === "shape-rect") {
      drawPreviewEl.style.background = "rgba(255, 212, 46, 0.14)";
      drawPreviewEl.style.border = "2px dashed " + PRESET_COLORS.yellow.solid;
      drawPreviewEl.style.borderRadius = "6px";
      drawPreviewEl.style.outline = "1px solid rgba(0,0,0,0.24)";
      drawPreviewEl.style.boxShadow = "0 0 0 1px rgba(255,255,255,0.5), 0 8px 20px rgba(0,0,0,0.18)";
    } else {
      drawPreviewEl.style.borderRadius = "6px";
    }
    (document.body || document.documentElement).appendChild(drawPreviewEl);
    updateShapePreview(event);
  }

  function updateShapePreview(event) {
    if (!drawStartPoint || !drawPreviewEl) return;
    const left = Math.min(drawStartPoint.x, event.pageX);
    const top = Math.min(drawStartPoint.y, event.pageY);
    const width = Math.abs(event.pageX - drawStartPoint.x);
    const height = Math.abs(event.pageY - drawStartPoint.y);
    drawMoved = drawMoved || width > 3 || height > 3;
    drawPreviewEl.style.left = left + "px";
    drawPreviewEl.style.top = top + "px";
    drawPreviewEl.style.width = Math.max(1, width) + "px";
    drawPreviewEl.style.height = Math.max(1, height) + "px";
  }

  function finishShapeDraw() {
    if (!drawStartPoint) return;
    const modeAtDraw = drawMode;
    const hadMoved = drawMoved;
    const drawSourceNode = drawStartSourceNode;
    const drawScrollAnchor = buildShapeScrollAnchor(drawSourceNode);
    const rect = drawPreviewEl ? {
      left: parseFloat(drawPreviewEl.style.left) || 0,
      top: parseFloat(drawPreviewEl.style.top) || 0,
      width: parseFloat(drawPreviewEl.style.width) || 0,
      height: parseFloat(drawPreviewEl.style.height) || 0
    } : null;

    clearDrawPreview();

    if (modeAtDraw && rect && hadMoved && rect.width >= 10 && rect.height >= 10) {
      if (!(modeAtDraw === "shape-rect" && isNestedInsideSolidRectangle(rect))) {
        const id = "hm-" + Date.now() + "-" + Math.random().toString(36).slice(2, 8);
        const snippet = modeAtDraw === "shape-rect" ? extractRectangleSnippet(rect) : "";
        const record = buildShapeRecord(id, modeAtDraw, lastUsedColor, lastUsedBgColor, rect, snippet, drawScrollAnchor);
        createShapeElement(record);
        storedHighlightsForPage.push(record);

        if (!record.transient) {
          sendRuntimeMessageSafe({ action: "saveHighlight", highlight: record });
        }
        setTimeout(() => pulseHighlight(id), 50);
      }
    }

    if (modeAtDraw) {
      setDrawMode(null);
    }

  }

  // ---- Shared positioning ----

  function positionElement(el, rect, elWidth, elHeight) {
    const GAP = 10;
    const vpW = window.innerWidth;
    const vpH = window.innerHeight;

    let left = rect.left + (rect.width / 2) - (elWidth / 2);
    let top = rect.top - elHeight - GAP;

    if (top < 4) top = rect.bottom + GAP;
    if (top + elHeight > vpH - 4) {
      top = rect.top + (rect.height / 2) - (elHeight / 2);
      left = rect.right + GAP;
    }

    left = Math.max(4, Math.min(left, vpW - elWidth - 4));
    top = Math.max(4, Math.min(top, vpH - elHeight - 4));

    el.style.left = left + "px";
    el.style.top = top + "px";
  }

  // ---- Selection Detection ----
  // FIX: Allow selection inside existing highlights so users can re-color

  let selectionDetected = false;

  document.addEventListener("mouseup", onMouseUp, true);
  window.addEventListener("mouseup", onMouseUp, true);
  document.addEventListener("mousedown", onMouseDown, true);

  function onMouseUp(e) {
    if (activeResize) {
      finishRectangleResize();
      return;
    }
    if (activeMove) {
      finishRectangleMove();
      return;
    }
    if (drawStartPoint) {
      finishShapeDraw();
      return;
    }
    if (pickerHost) {
      const path = e.composedPath ? e.composedPath() : [];
      for (const node of path) {
        if (node === pickerHost) return;
      }
    }
    if (hoverHost) {
      const path = e.composedPath ? e.composedPath() : [];
      for (const node of path) {
        if (node === hoverHost) return;
      }
    }
    if (colorPickerOpen) return;

    selectionDetected = false;
    tryDetectSelection(10);
    tryDetectSelection(60);
    tryDetectSelection(180);
  }

  function onMouseDown(e) {
    if (textHighlightContextMenuEl) {
      const targetEl = getEventTargetElement(e);
      if (targetEl && textHighlightContextMenuEl.contains(targetEl)) return;
      hideTextHighlightContextMenu();
    }
    if (isEventFromOcrCopyChip(e)) return;
    if (startRectangleResize(e)) return;
    if (startRectangleMove(e)) return;
    if (drawMode) {
      startShapeDraw(e);
      return;
    }
    if (colorPickerOpen) return;
    if (hoverColorPickerOpen) return;

    if (pickerHost) {
      const path = e.composedPath ? e.composedPath() : [];
      for (const node of path) {
        if (node === pickerHost) return;
      }
    }
    if (hoverHost) {
      const path = e.composedPath ? e.composedPath() : [];
      for (const node of path) {
        if (node === hoverHost) return;
      }
    }
    hidePicker();
    hideHoverToolbar();
    selectionDetected = false;
  }

  function tryDetectSelection(delay) {
    setTimeout(() => {
      if (selectionDetected) return;
      if (!extensionEnabled) return;
      if (drawMode) return;
      if (colorPickerOpen) return;

      const sel = window.getSelection();
      if (!sel || sel.isCollapsed || !sel.toString().trim()) return;

      const anchor = sel.anchorNode;
      if (anchor) {
        const el = anchor.nodeType === 3 ? anchor.parentElement : anchor;
        if (el && el.closest('textarea, input, [contenteditable="true"], [contenteditable=""], [contenteditable="plaintext-only"]')) {
          return;
        }
      }

      try {
        const range = sel.getRangeAt(0);
        const rect = range.getBoundingClientRect();
        if (rect.width === 0 && rect.height === 0) return;

        // FIX: REMOVED the guard that blocked selection inside existing highlights
        // This allows users to select highlighted text and re-color it

        currentSelection = {
          range: range.cloneRange(),
          text: sel.toString(),
          note: ""
        };
        pickerNoteDraft = "";
        pickerNoteEditorOpen = false;

        selectionDetected = true;
        hideHoverToolbar();
        if (continuousHighlightEnabled) {
          createHighlightFromSelection(continuousColorName, continuousBgColor);
        } else {
          showPicker(rect);
        }
        setTimeout(() => { selectionDetected = false; }, 500);
      } catch (_err) {}
    }, delay);
  }

  document.addEventListener("keyup", (e) => {
    if (!extensionEnabled) return;
    if (colorPickerOpen) return;
    if (e.shiftKey || (e.ctrlKey && e.key === "a") || (e.metaKey && e.key === "a")) {
      selectionDetected = false;
      tryDetectSelection(80);
      tryDetectSelection(250);
    }
  }, true);

  let scrollTimer;
  const onAnyScroll = () => {
    hideTextHighlightContextMenu();
    if (pickerEl && !colorPickerOpen) {
      clearTimeout(scrollTimer);
      scrollTimer = setTimeout(() => hidePicker(), 200);
    }
    if (hoverEl && !hoverColorPickerOpen) {
      clearTimeout(hoverTimeout);
      hoverTimeout = setTimeout(() => hideHoverToolbar(), HOVER_TOOLBAR_HIDE_DELAY_MS);
    }
    scheduleShapePositionSync();
  };
  window.addEventListener("scroll", onAnyScroll, { passive: true, capture: true });
  document.addEventListener("scroll", onAnyScroll, { passive: true, capture: true });
  window.addEventListener("resize", scheduleShapePositionSync, { passive: true });

  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      e.preventDefault();
      e.stopPropagation();
      hideTextHighlightContextMenu();
      if (colorPickerOpen) colorPickerOpen = false;
      if (hoverColorPickerOpen) hoverColorPickerOpen = false;
      hidePicker();
      hideHoverToolbar();
      cancelInteractionState(true);
      if (window.getSelection) {
        try { window.getSelection().removeAllRanges(); } catch (_err) {}
      }
      return;
    }

    const keyLower = (e.key || "").toLowerCase();
    const activeRectangleId = getActiveRectangleTargetId();

    if (!e.altKey && !e.ctrlKey && !e.metaKey && e.key === "Delete" && continuousHighlightEnabled) {
      if (isEditableShortcutTarget(e)) return;
      e.preventDefault();
      e.stopPropagation();
      continuousHighlightEnabled = false;
      showInlineUndoToast("Continuous highlight off", null, 2200);
      return;
    }

    if (pickerEl && !e.altKey && !e.ctrlKey && !e.metaKey && !e.shiftKey && keyLower === "n") {
      if (pickerNoteEditorOpen) return;
      if (isEditableShortcutTarget(e)) return;
      e.preventDefault();
      e.stopPropagation();
      applyDefaultHighlightForNoteFlow();
      return;
    }

    if (!e.altKey && !e.ctrlKey && !e.metaKey && e.key === "Delete") {
      if (isEditableShortcutTarget(e)) return;
      if (activeRectangleId) {
        e.preventDefault();
        e.stopPropagation();
        hideHoverToolbar();
        deleteHighlightsWithUndo([activeRectangleId]);
        return;
      }
    }

    if (e.altKey && !e.ctrlKey && !e.metaKey && keyLower === "t" && activeRectangleId) {
      if (isEditableShortcutTarget(e)) return;
      const rectRecord = storedHighlightsForPage.find((x) => idsMatch(x && x.id, activeRectangleId) && x.type === "shape-rect");
      if (rectRecord) {
        e.preventDefault();
        e.stopPropagation();
        normalizeShapeRecord(rectRecord, false);
        setRectangleRevealState(activeRectangleId, !rectRecord.revealed);
        refreshHoverToolbar(activeRectangleId);
        return;
      }
    }
    
    if (!e.altKey && e.ctrlKey && e.shiftKey && !e.metaKey && keyLower === "h") {
      if (isEditableShortcutTarget(e)) return;
      e.preventDefault();
      e.stopPropagation();
      const nextEnabled = !extensionEnabled;
      extensionEnabled = nextEnabled;
      if (!nextEnabled) {
        hidePicker();
        hideHoverToolbar();
        cancelInteractionState(true);
      } else {
        updateDrawDockState();
      }
      sendRuntimeMessageSafe({ action: "setEnabled", enabled: nextEnabled });
      showInlineUndoToast(nextEnabled ? "HighlightMaster enabled" : "HighlightMaster paused", null, 2200);
      return;
    }

    if (e.altKey && !e.ctrlKey && !e.metaKey && !e.shiftKey) {
      if (!extensionEnabled) return;
      if (isEditableShortcutTarget(e)) return;
      const shortcutNum = Number(keyLower);
      if (Number.isFinite(shortcutNum) && shortcutNum >= 1 && shortcutNum <= 5) {
        e.preventDefault();
        e.stopPropagation();
        const colorName = QUICK_SHORTCUT_COLORS[shortcutNum - 1];
        const color = PRESET_COLORS[colorName];
        if (!color) return;
        lastUsedColor = colorName;
        lastUsedBgColor = color.bg;
        createHighlightFromSelection(colorName, color.bg);
        return;
      }
    }

    // Alt+H keeps legacy quick-highlight behavior with the last used color.
    if (e.altKey && keyLower === "h") {
      if (!extensionEnabled) return;
      if (isEditableShortcutTarget(e)) return;
      e.preventDefault();
      createHighlightFromSelection(lastUsedColor, lastUsedBgColor);
    }
  }, true);

  document.addEventListener("mousemove", (e) => {
    if (activeResize) {
      e.preventDefault();
      updateRectangleResize(e);
      return;
    }
    if (activeMove) {
      e.preventDefault();
      updateRectangleMove(e);
      return;
    }
    if (drawStartPoint) {
      e.preventDefault();
      updateShapePreview(e);
    }
  }, true);

  window.addEventListener("blur", () => {
    if (activeResize) finishRectangleResize();
    if (activeMove) finishRectangleMove();
    hideTextHighlightContextMenu();
    clearInlineUndoToast();
    setDrawMode(null);
  });

  // ---- Highlight Creation ----

  function handleColorClick(colorName, bgColor, noteText) {
    lastUsedColor = colorName;
    lastUsedBgColor = bgColor;
    if (continuousHighlightEnabled) {
      setContinuousHighlightColor(colorName, bgColor);
    }
    createHighlightFromSelection(colorName, bgColor, noteText);
  }

  function collectTextHighlightIdsInRange(range) {
    const ids = new Set();
    if (!range) return ids;
    const textNodes = getTextNodesInRange(range);
    if (!textNodes.length && range.startContainer && range.startContainer.nodeType === 3) {
      textNodes.push(range.startContainer);
    }
    if (!textNodes.length && range.endContainer && range.endContainer.nodeType === 3) {
      textNodes.push(range.endContainer);
    }

    for (const node of textNodes) {
      let current = node && node.parentNode ? node.parentNode : null;
      while (current && current !== document.body && current !== document.documentElement) {
        if (current.nodeType === 1 && current.getAttribute) {
          const id = current.getAttribute(HIGHLIGHT_ATTR);
          if (id) {
            const kind = current.getAttribute(HIGHLIGHT_KIND_ATTR) || "";
            if (kind.indexOf("shape-") !== 0) {
              ids.add(id);
            }
          }
        }
        current = current.parentNode;
      }
    }

    return ids;
  }

  function removeExistingTextHighlightsForRange(range) {
    const highlightIds = collectTextHighlightIdsInRange(range);
    if (!highlightIds.size) return;

    const targets = [];
    for (const id of highlightIds) {
      const record = storedHighlightsForPage.find((x) => idsMatch(x && x.id, id));
      if (record && isShapeRecord(record)) continue;
      removeHighlightFromPage(id);
      if (record && !record.transient) {
        targets.push({ domain: record.domain || DOMAIN, id: normalizeHighlightIdValue(record.id) });
      } else {
        targets.push({ domain: DOMAIN, id: normalizeHighlightIdValue(id) });
      }
    }

    if (!targets.length) return;
    sendRuntimeMessageSafe({
      action: "bulkDeleteHighlights",
      targets: targets,
      historyAction: "bulk_delete",
      historyLabel: targets.length === 1 ? "Replaced 1 highlight color" : ("Replaced " + targets.length + " highlight colors"),
      pageUrl: location.href,
      domain: DOMAIN
    });
  }

  function applyHighlight(range, id, colorName, bgColor) {
    if (range.startContainer === range.endContainer && range.startContainer.nodeType === 3) {
      // Check if parent is already a highlight span - if so, skip the nesting guard
      // since we already unwrapped in handleColorClick
      try {
        const span = makeSpan(id, colorName, bgColor);
        range.surroundContents(span);
        return;
      } catch (_e) {}
    }

    const textNodes = getTextNodesInRange(range);
    if (textNodes.length === 0) return;

    for (let i = textNodes.length - 1; i >= 0; i--) {
      const tn = textNodes[i];
      // Allow nested highlights by NOT skipping text nodes inside existing highlights
      // if (tn.parentNode && tn.parentNode.closest && tn.parentNode.closest("[" + HIGHLIGHT_ATTR + "]")) continue;

      let startOffset = 0;
      let endOffset = tn.length;
      if (tn === range.startContainer) startOffset = range.startOffset;
      if (tn === range.endContainer) endOffset = range.endOffset;

      if (endOffset <= startOffset) continue;

      try {
        const span = makeSpan(id, colorName, bgColor);
        if (startOffset === 0 && endOffset === tn.length) {
          tn.parentNode.insertBefore(span, tn);
          span.appendChild(tn);
        } else {
          const subRange = document.createRange();
          subRange.setStart(tn, startOffset);
          subRange.setEnd(tn, endOffset);
          subRange.surroundContents(span);
        }
      } catch (_e) {}
    }
  }

  function makeSpan(id, colorName, bgColor) {
    const span = document.createElement("span");
    span.setAttribute(HIGHLIGHT_ATTR, id);
    span.setAttribute(HIGHLIGHT_COLOR_ATTR, colorName);
    span.classList.add("hm-pop-anim");
    span.style.setProperty("background-color", bgColor, "important");
    span.style.removeProperty("color");
    span.style.setProperty("display", "inline", "important");
    span.style.setProperty("visibility", "visible", "important");
    span.style.setProperty("opacity", "1", "important");
    return span;
  }

  function getTextNodesInRange(range) {
    const nodes = [];
    const ancestor = range.commonAncestorContainer;
    const root = ancestor.nodeType === 3 ? ancestor.parentNode : ancestor;

    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
      acceptNode(node) {
        try {
          const nodeRange = document.createRange();
          nodeRange.selectNodeContents(node);
          if (range.compareBoundaryPoints(Range.END_TO_START, nodeRange) < 0 &&
              range.compareBoundaryPoints(Range.START_TO_END, nodeRange) > 0) {
            return NodeFilter.FILTER_ACCEPT;
          }
        } catch (_e) {}
        return NodeFilter.FILTER_REJECT;
      }
    });

    let node;
    while ((node = walker.nextNode())) nodes.push(node);
    return nodes;
  }

  // ---- Record Building ----

  function buildHighlightRecord(id, colorName, bgColor, text, range) {
    return {
      id,
      domain: DOMAIN,
      url: location.href,
      text,
      note: "",
      color: colorName,
      bgColor,
      timestamp: Date.now(),
      pageTitle: document.title || "",
      startXPath: getXPath(range.startContainer),
      startOffset: range.startOffset,
      endXPath: getXPath(range.endContainer),
      endOffset: range.endOffset,
      contextBefore: getContextBefore(range, CONTEXT_CHARS),
      contextAfter: getContextAfter(range, CONTEXT_CHARS)
    };
  }

  function getXPath(node) {
    if (!node) return "";
    if (node.nodeType === 3) {
      const parent = node.parentNode;
      if (!parent) return "";
      const parentPath = getXPath(parent);
      let textIndex = 1;
      for (const child of parent.childNodes) {
        if (child === node) break;
        if (child.nodeType === 3) textIndex++;
      }
      return parentPath + "/text()[" + textIndex + "]";
    }
    if (node === document.body) return "/html/body";
    if (node === document.documentElement) return "/html";
    if (node === document) return "/";

    const parent = node.parentNode;
    if (!parent) return "/" + node.nodeName.toLowerCase();

    let index = 1;
    for (const sibling of parent.children) {
      if (sibling === node) break;
      if (sibling.nodeName === node.nodeName) index++;
    }
    const tagName = node.nodeName.toLowerCase();
    let sameTagCount = 0;
    try {
      sameTagCount = parent.querySelectorAll(":scope > " + tagName).length;
    } catch (_e) {
      sameTagCount = 2;
    }
    const suffix = sameTagCount > 1 ? "[" + index + "]" : "";
    return getXPath(parent) + "/" + tagName + suffix;
  }

  function getContextBefore(range, chars) {
    try {
      const node = range.startContainer;
      const text = node.nodeType === 3 ? node.textContent : (node.innerText || "");
      return text.substring(0, range.startOffset).slice(-chars);
    } catch (_e) { return ""; }
  }

  function getContextAfter(range, chars) {
    try {
      const node = range.endContainer;
      const text = node.nodeType === 3 ? node.textContent : (node.innerText || "");
      return text.substring(range.endOffset).slice(0, chars);
    } catch (_e) { return ""; }
  }

  // ---- Restore Logic (Improved) ----

  let restoreInProgress = false;
  let restoreAttempts = 0;
  let restoreMutationTimer = null;
  const MAX_RESTORE_ATTEMPTS = 15;

  function scheduleRestoreRetry(delayMs) {
    clearTimeout(restoreMutationTimer);
    restoreMutationTimer = setTimeout(() => {
      restoreMutationTimer = null;
      restoreHighlights();
    }, Math.max(60, delayMs || 0));
  }

  function collectExistingHighlightIds() {
    const ids = new Set();
    const nodes = document.querySelectorAll("[" + HIGHLIGHT_ATTR + "]");
    for (const node of nodes) {
      const id = node.getAttribute(HIGHLIGHT_ATTR);
      if (id) ids.add(id);
    }
    return ids;
  }

  function restoreHighlights() {
    if (restoreInProgress) return;
    restoreInProgress = true;

    sendRuntimeMessageSafe({ action: "getHighlightsForDomain", domain: DOMAIN }, (res, errorMessage) => {
      restoreInProgress = false;
      if (errorMessage || !res || !res.highlights) return;

      const normalizedRestoredIds = new Set();
      for (const rid of restoredIds) {
        const normalizedRid = normalizeHighlightIdValue(rid);
        if (normalizedRid) normalizedRestoredIds.add(normalizedRid);
      }
      restoredIds = normalizedRestoredIds;

      storedHighlightsForPage = res.highlights
        .filter((h) => h && h.url && isCurrentPageUrl(h.url))
        .map((h) => ({
          ...h,
          id: normalizeHighlightIdValue(h.id),
          domain: typeof h.domain === "string" && h.domain ? h.domain : DOMAIN
        }))
        .filter((h) => !!h.id);
      const existingIds = collectExistingHighlightIds();
      let textSearchIndex = null;
      const restoreContext = {
        getSearchIndex() {
          if (!textSearchIndex) {
            textSearchIndex = buildTextSearchIndex(document.body);
          }
          return textSearchIndex;
        }
      };

      let allRestored = true;
      for (const h of storedHighlightsForPage) {
        const hid = normalizeHighlightIdValue(h.id);
        if (!hid) continue;
        if (existingIds.has(hid)) {
          if (!syncHighlightVisualFromRecord(h)) {
            allRestored = false;
          }
          restoredIds.add(hid);
          continue;
        }
        if (restoredIds.has(hid)) {
          restoredIds.delete(hid);
        }
        if (tryRestoreHighlight(h, restoreContext)) {
          restoredIds.add(hid);
          existingIds.add(hid);
        } else {
          allRestored = false;
        }
      }

      scheduleShapePositionSync();

      if (!allRestored && restoreAttempts < MAX_RESTORE_ATTEMPTS) {
        restoreAttempts++;
        scheduleRestoreRetry(Math.min(500 * restoreAttempts, 5000));
      } else if (allRestored) {
        restoreAttempts = 0;
      }
    });
  }

  function tryRestoreHighlight(h, restoreContext) {
    const colorName = h.color || "yellow";
    const fallbackBgColor = colorName.startsWith("custom:")
      ? colorName.replace("custom:", "")
      : (PRESET_COLORS[colorName] ? PRESET_COLORS[colorName].bg : PRESET_COLORS.yellow.bg);
    const bgColor = h.bgColor || fallbackBgColor;

    if (isShapeRecord(h)) {
      if (!h.rect) return false;
      const restoredShape = normalizeShapeRecord({
        ...h,
        color: colorName,
        bgColor
      }, true);
      createShapeElement(restoredShape);
      return true;
    }

    // Strategy 1: XPath
    try {
      const startNode = resolveXPath(h.startXPath);
      const endNode = resolveXPath(h.endXPath);
      if (startNode && endNode) {
        const range = document.createRange();
        const startMax = startNode.nodeType === 3 ? startNode.length : (startNode.childNodes ? startNode.childNodes.length : 0);
        const endMax = endNode.nodeType === 3 ? endNode.length : (endNode.childNodes ? endNode.childNodes.length : 0);
        range.setStart(startNode, Math.min(h.startOffset, startMax));
        range.setEnd(endNode, Math.min(h.endOffset, endMax));
        if (normalizeWS(range.toString()) === normalizeWS(h.text)) {
          applyHighlight(range, h.id, colorName, bgColor);
          const els = getAnnotationElements(h.id);
          for (const el of els) applyNoteTooltip(el, h.note || "");
          return true;
        }
      }
    } catch (_e) {}

    // Strategy 2: Text search with context
    try {
      const found = findTextInPage(h.text, h.contextBefore || "", h.contextAfter || "", restoreContext ? restoreContext.getSearchIndex() : null);
      if (found) {
        applyHighlight(found, h.id, colorName, bgColor);
        const els = getAnnotationElements(h.id);
        for (const el of els) applyNoteTooltip(el, h.note || "");
        return true;
      }
    } catch (_e) {}

    // Strategy 3: Loose text search
    try {
      const found = findTextInPage(h.text, "", "", restoreContext ? restoreContext.getSearchIndex() : null);
      if (found) {
        applyHighlight(found, h.id, colorName, bgColor);
        const els = getAnnotationElements(h.id);
        for (const el of els) applyNoteTooltip(el, h.note || "");
        return true;
      }
    } catch (_e) {}

    return false;
  }

  function syncHighlightVisualFromRecord(h) {
    if (!h || !h.id) return false;
    const elements = getAnnotationElements(h.id);
    if (!elements.length) return false;

    const colorName = h.color || "yellow";
    const fallbackBgColor = colorName.startsWith("custom:")
      ? colorName.replace("custom:", "")
      : (PRESET_COLORS[colorName] ? PRESET_COLORS[colorName].bg : PRESET_COLORS.yellow.bg);
    const bgColor = h.bgColor || fallbackBgColor;

    if (isShapeRecord(h)) {
      if (!h.rect) return false;
      const normalizedShape = normalizeShapeRecord({
        ...h,
        color: colorName,
        bgColor
      }, true);
      if (!normalizedShape || !normalizedShape.rect) return false;

      for (const el of elements) {
        inferShapeScrollAnchorIfMissing(normalizedShape);
        el.setAttribute(HIGHLIGHT_COLOR_ATTR, normalizedShape.color || colorName);
        el.setAttribute(HIGHLIGHT_KIND_ATTR, normalizedShape.type || "shape-rect");
        const renderedRect = getRenderedRectForRecord(normalizedShape) || normalizedShape.rect;
        el.style.left = renderedRect.left + "px";
        el.style.top = renderedRect.top + "px";
        el.style.width = renderedRect.width + "px";
        el.style.height = renderedRect.height + "px";
        if (normalizedShape.type === "shape-rect") {
          ensureRectangleResizeHandles(el);
        }
        applyShapeVisual(el, normalizedShape, normalizedShape.bgColor || bgColor);
        applyNoteTooltip(el, h.note || "");
      }
      return true;
    }

    for (const el of elements) {
      if ((el.getAttribute(HIGHLIGHT_KIND_ATTR) || "").indexOf("shape-") === 0) continue;
      el.setAttribute(HIGHLIGHT_COLOR_ATTR, colorName);
      el.style.setProperty("background-color", bgColor, "important");
      el.style.removeProperty("color");
      applyNoteTooltip(el, h.note || "");
    }
    return true;
  }

  function resolveXPath(xpath) {
    if (!xpath) return null;
    try {
      const result = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
      return result.singleNodeValue;
    } catch (_e) { return null; }
  }

  function findTextInPage(searchText, contextBefore, contextAfter, searchIndex) {
    const normalized = normalizeWS(searchText);
    if (!normalized) return null;
    const textIndex = searchIndex || buildTextSearchIndex(document.body);
    if (!textIndex) return null;

    const fullText = textIndex.fullText || "";
    const nodeMap = Array.isArray(textIndex.nodeMap) ? textIndex.nodeMap : [];
    const normalizedFull = textIndex.normalizedFull || "";
    if (!normalizedFull) return null;
    const occurrences = [];
    let searchFrom = 0;
    while (true) {
      const idx = normalizedFull.indexOf(normalized, searchFrom);
      if (idx === -1) break;
      occurrences.push(idx);
      searchFrom = idx + 1;
    }

    if (occurrences.length === 0) {
      const lowerFull = normalizedFull.toLowerCase();
      const lowerSearch = normalized.toLowerCase();
      searchFrom = 0;
      while (true) {
        const idx = lowerFull.indexOf(lowerSearch, searchFrom);
        if (idx === -1) break;
        occurrences.push(idx);
        searchFrom = idx + 1;
      }
      if (occurrences.length === 0) return null;
    }

    let bestIdx = occurrences[0];
    if (occurrences.length > 1 && (contextBefore || contextAfter)) {
      let bestScore = -1;
      for (const idx of occurrences) {
        let score = 0;
        if (contextBefore) {
          const before = normalizedFull.substring(Math.max(0, idx - contextBefore.length), idx);
          score += similarity(before, normalizeWS(contextBefore));
        }
        if (contextAfter) {
          const afterStart = idx + normalized.length;
          const after = normalizedFull.substring(afterStart, afterStart + contextAfter.length);
          score += similarity(after, normalizeWS(contextAfter));
        }
        if (score > bestScore) {
          bestScore = score;
          bestIdx = idx;
        }
      }
    }

    const origStart = mapNormIdxToOrig(fullText, bestIdx);
    const origEnd = mapNormIdxToOrig(fullText, bestIdx + normalized.length);

    let startNode = null, startOffset = 0;
    let endNode = null, endOffset = 0;

    for (const nm of nodeMap) {
      if (!startNode && origStart < nm.end) {
        startNode = nm.node;
        startOffset = origStart - nm.start;
      }
      if (origEnd <= nm.end) {
        endNode = nm.node;
        endOffset = origEnd - nm.start;
        break;
      }
    }

    if (startNode && endNode) {
      try {
        const range = document.createRange();
        range.setStart(startNode, Math.min(startOffset, startNode.length));
        range.setEnd(endNode, Math.min(endOffset, endNode.length));
        return range;
      } catch (_e) { return null; }
    }
    return null;
  }

  function normalizeWS(str) {
    return (str || "").replace(/\s+/g, " ").trim();
  }

  function mapNormIdxToOrig(original, normalizedIdx) {
    let ni = 0, oi = 0;
    while (oi < original.length && /\s/.test(original[oi])) oi++;
    while (ni < normalizedIdx && oi < original.length) {
      if (/\s/.test(original[oi])) {
        while (oi < original.length && /\s/.test(original[oi])) oi++;
        ni++;
      } else {
        oi++;
        ni++;
      }
    }
    return oi;
  }

  function similarity(a, b) {
    if (!a || !b) return 0;
    const maxLen = Math.max(a.length, b.length);
    if (maxLen === 0) return 1;
    let matches = 0;
    const minLen = Math.min(a.length, b.length);
    for (let i = 0; i < minLen; i++) {
      if (a[i] === b[i]) matches++;
    }
    return matches / maxLen;
  }

  // ---- MutationObserver: Highlight Guardian ----

  let reapplyTimer = null;
  let reapplyInProgress = false;

  function startHighlightGuardian() {
    if (!document.body) return;
    const observer = new MutationObserver((mutations) => {
      let lost = false;
      let addedContent = false;
      for (const m of mutations) {
        if (m.addedNodes && m.addedNodes.length) {
          addedContent = true;
        }
        for (const node of m.removedNodes) {
          if (node.nodeType === 1) {
            if (node.hasAttribute && node.hasAttribute(HIGHLIGHT_ATTR)) { lost = true; break; }
            if (node.querySelector && node.querySelector("[" + HIGHLIGHT_ATTR + "]")) { lost = true; break; }
          }
        }
        if (lost) break;
      }
      if (lost && !reapplyInProgress) {
        clearTimeout(reapplyTimer);
        reapplyTimer = setTimeout(reapplyLost, 250);
      } else if (addedContent && storedHighlightsForPage.length > restoredIds.size) {
        scheduleRestoreRetry(180);
      }
    });
    observer.observe(document.body, { childList: true, subtree: true });
  }

  function reapplyLost() {
    if (reapplyInProgress) return;
    reapplyInProgress = true;
    const restoreContext = {
      searchIndex: null,
      getSearchIndex() {
        if (!this.searchIndex) {
          this.searchIndex = buildTextSearchIndex(document.body);
        }
        return this.searchIndex;
      }
    };
    for (const h of storedHighlightsForPage) {
      const hid = normalizeHighlightIdValue(h && h.id);
      if (!hid) continue;
      if (!document.querySelector("[" + HIGHLIGHT_ATTR + '="' + hid + '"]')) {
        restoredIds.delete(hid);
        if (tryRestoreHighlight(h, restoreContext)) {
          restoredIds.add(hid);
        }
      }
    }
    injectHighlightStyles();
    setTimeout(() => { reapplyInProgress = false; }, 400);
  }

  // Integrity check - less frequent to save memory
  setInterval(() => {
    if (document.hidden) return;
    if (!extensionEnabled) return;
    if (storedHighlightsForPage.length === 0) return;
    let missing = false;
    for (const h of storedHighlightsForPage) {
      const hid = normalizeHighlightIdValue(h && h.id);
      if (!hid) continue;
      if (!document.querySelector("[" + HIGHLIGHT_ATTR + '="' + hid + '"]')) {
        missing = true;
        break;
      }
    }
    if (missing && !reapplyInProgress) reapplyLost();
  }, 5000);

  // ---- SPA Navigation Detection ----

  let lastUrl = location.href;

  const origPushState = history.pushState;
  const origReplaceState = history.replaceState;

  history.pushState = function () {
    origPushState.apply(this, arguments);
    onUrlChange();
  };
  history.replaceState = function () {
    origReplaceState.apply(this, arguments);
    onUrlChange();
  };
  window.addEventListener("popstate", onUrlChange);

  function onUrlChange() {
    if (location.href !== lastUrl) {
      lastUrl = location.href;
      clearInlineUndoToast();
      restoredIds.clear();
      storedHighlightsForPage = [];
      restoreAttempts = 0;
      hidePicker();
      hideHoverToolbar();
      cancelInteractionState(true);
      setTimeout(restoreHighlights, 300);
      setTimeout(restoreHighlights, 1000);
      setTimeout(restoreHighlights, 3000);
    }
  }

  if (document.body) {
    const navObs = new MutationObserver(() => {
      if (location.href !== lastUrl) onUrlChange();
    });
    navObs.observe(document.body, { childList: true, subtree: true });
  }

  // ---- Scroll to Highlight ----

  function pulseHighlight(id) {
    const el = document.querySelector("[" + HIGHLIGHT_ATTR + '="' + id + '"]');
    if (!el) return;
    el.classList.add("hm-pulse");
    setTimeout(() => el.classList.remove("hm-pulse"), 2000);
  }

  function scrollToHighlight(id) {
    const el = document.querySelector("[" + HIGHLIGHT_ATTR + '="' + id + '"]');
    if (!el) return false;
    el.scrollIntoView({ behavior: "smooth", block: "center" });
    pulseHighlight(id);
    return true;
  }

  // ---- Remove / Clear Highlights ----

  function sortNodesDeepestFirst(nodes) {
    const ordered = Array.from(nodes || []).filter((node) => !!node && node.nodeType === 1);
    ordered.sort((a, b) => {
      if (a === b) return 0;
      const relation = a.compareDocumentPosition ? a.compareDocumentPosition(b) : 0;
      if (relation & Node.DOCUMENT_POSITION_CONTAINED_BY) return 1;
      if (relation & Node.DOCUMENT_POSITION_CONTAINS) return -1;
      if (relation & Node.DOCUMENT_POSITION_FOLLOWING) return -1;
      if (relation & Node.DOCUMENT_POSITION_PRECEDING) return 1;
      return 0;
    });
    return ordered;
  }

  function unwrapAnnotationNode(node, parentsToNormalize) {
    if (!node) return;
    const parent = node.parentNode;
    if (!parent) return;
    const kind = node.getAttribute ? (node.getAttribute(HIGHLIGHT_KIND_ATTR) || "") : "";
    if (kind.indexOf("shape-") === 0) {
      parent.removeChild(node);
      return;
    }
    const fragment = document.createDocumentFragment();
    while (node.firstChild) {
      fragment.appendChild(node.firstChild);
    }
    parent.replaceChild(fragment, node);
    if (parentsToNormalize) parentsToNormalize.add(parent);
  }

  function removeAnnotationNodes(nodes) {
    const parentsToNormalize = new Set();
    const ordered = sortNodesDeepestFirst(nodes);
    for (const node of ordered) {
      unwrapAnnotationNode(node, parentsToNormalize);
    }
    for (const parent of parentsToNormalize) {
      if (!parent || typeof parent.normalize !== "function") continue;
      parent.normalize();
    }
  }

  function removeHighlightFromPage(id) {
    const targetId = normalizeHighlightIdValue(id);
    if (!targetId) return;
    if (idsMatch(textHighlightContextMenuId, targetId)) {
      hideTextHighlightContextMenu();
    }
    if (activeResize && idsMatch(activeResize.id, targetId)) {
      activeResize = null;
      setRectangleResizeCursor("se", false);
    }
    if (activeMove && idsMatch(activeMove.id, targetId)) {
      activeMove = null;
      setRectangleMoveCursor(false);
    }
    const nodes = document.querySelectorAll("[" + HIGHLIGHT_ATTR + '="' + targetId + '"]');
    removeAnnotationNodes(nodes);
    restoredIds.delete(targetId);
    storedHighlightsForPage = storedHighlightsForPage.filter((h) => !idsMatch(h && h.id, targetId));
  }

  function clearAllHighlightsFromPage() {
    activeResize = null;
    setRectangleResizeCursor("se", false);
    activeMove = null;
    setRectangleMoveCursor(false);
    clearInlineUndoToast();
    hideTextHighlightContextMenu();
    const nodes = document.querySelectorAll("[" + HIGHLIGHT_ATTR + "]");
    removeAnnotationNodes(nodes);
    restoredIds.clear();
    storedHighlightsForPage = [];
  }

  // ---- Message Listener ----

  chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
    if (!msg || !msg.action) { sendResponse({}); return; }

    switch (msg.action) {
      case "ping":
        sendResponse({ ok: true });
        break;

      case "enabledChanged":
        extensionEnabled = msg.enabled;
        if (!extensionEnabled) {
          continuousHighlightEnabled = false;
          hideTextHighlightContextMenu();
          hidePicker();
          hideHoverToolbar();
          cancelInteractionState(true);
        } else {
          updateDrawDockState();
        }
        sendResponse({ ok: true });
        break;

      case "autoOcrCopyChanged":
        autoOcrCopyEnabled = msg.enabled === true;
        sendResponse({ ok: true });
        break;

      case "setDrawMode": {
        if (!extensionEnabled) {
          cancelInteractionState(true);
          sendResponse({ ok: false, reason: "disabled", drawMode: null });
          break;
        }
        hidePicker();
        hideHoverToolbar();
        if (window.getSelection) {
          try { window.getSelection().removeAllRanges(); } catch (_e) {}
        }
        const newMode = setDrawMode(msg.mode);
        sendResponse({ ok: true, drawMode: newMode });
        break;
      }

      case "getDrawMode":
        sendResponse({ ok: true, drawMode: extensionEnabled ? drawMode : null });
        break;

      case "getContinuousMode":
        sendResponse({ ok: true, ...getContinuousModeState() });
        break;

      case "setContinuousMode": {
        const payload = normalizeContinuousColorPayload(msg.colorName, msg.bgColor);
        setContinuousHighlightColor(payload.colorName, payload.bgColor);
        const requestedEnabled = typeof msg.enabled === "boolean"
          ? msg.enabled
          : continuousHighlightEnabled;
        continuousHighlightEnabled = !!requestedEnabled && extensionEnabled;
        sendResponse({
          ok: true,
          ...getContinuousModeState(),
          reason: (requestedEnabled && !extensionEnabled) ? "disabled" : undefined
        });
        break;
      }

      case "cancelDrawMode":
        cancelInteractionState(true);
        sendResponse({ ok: true });
        break;

      case "scrollToHighlight": {
        let attempts = 0;
        const tryScroll = () => {
          if (scrollToHighlight(msg.id)) {
            sendResponse({ ok: true });
          } else if (attempts < 10) {
            attempts++;
            setTimeout(tryScroll, 400);
          } else {
            sendResponse({ ok: false });
          }
        };
        tryScroll();
        return true;
      }

      case "removeHighlightFromPage":
        removeHighlightFromPage(msg.id);
        sendResponse({ ok: true });
        break;

      case "removeHighlightsBatch": {
        const ids = Array.isArray(msg.ids) ? msg.ids : [];
        for (const id of ids) {
          if (id) removeHighlightFromPage(id);
        }
        sendResponse({ ok: true, removed: ids.length });
        break;
      }

      case "clearHighlightsFromPage":
        clearAllHighlightsFromPage();
        sendResponse({ ok: true });
        break;

      case "highlightsUpdated":
        cancelInteractionState(false);
        hideTextHighlightContextMenu();
        restoredIds.clear();
        storedHighlightsForPage = [];
        restoreAttempts = 0;
        restoreHighlights();
        sendResponse({ ok: true });
        break;

      default:
        sendResponse({});
    }
  });

  // ---- Initialization ----

  function init() {
    sendRuntimeMessageSafe({ action: "getEnabled" }, (res, errorMessage) => {
      if (errorMessage) return;
      extensionEnabled = res && res.enabled !== false;
      updateDrawDockState();
    });
    sendRuntimeMessageSafe({ action: "getAutoOcrCopy" }, (res, errorMessage) => {
      if (errorMessage) return;
      autoOcrCopyEnabled = !!(res && res.enabled);
    });

    const delays = [100, 300, 600, 1000, 1500, 2500, 4000, 6000];
    for (const d of delays) {
      setTimeout(restoreHighlights, d);
    }

    if (document.body) {
      startHighlightGuardian();
    } else {
      document.addEventListener("DOMContentLoaded", startHighlightGuardian);
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }

  window.addEventListener("load", () => {
    setTimeout(restoreHighlights, 500);
    setTimeout(restoreHighlights, 2000);
  });
})();
