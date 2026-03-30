/**
 * Optional verbose tracing for translation / chat flows (console + footer line).
 * Enable in Settings → Developer, or: localStorage.setItem("legal-word-addin:debug-trace","1")
 */

const STORAGE_KEY = "legal-word-addin:debug-trace";

/** @type {HTMLElement | null} */
let footerLineEl = null;

export function isDebugTraceEnabled() {
  try {
    return localStorage.getItem(STORAGE_KEY) === "1";
  } catch {
    return false;
  }
}

export function setDebugTraceEnabled(on) {
  try {
    if (on) localStorage.setItem(STORAGE_KEY, "1");
    else localStorage.removeItem(STORAGE_KEY);
  } catch {
    /* quota / private mode */
  }
  syncFooterVisibility();
}

/** @param {HTMLElement | null} el */
export function initDebugTraceFooter(el) {
  footerLineEl = el;
  syncFooterVisibility();
}

function syncFooterVisibility() {
  if (!footerLineEl) return;
  footerLineEl.classList.toggle("is-hidden", !isDebugTraceEnabled());
  if (!isDebugTraceEnabled()) footerLineEl.textContent = "";
}

/**
 * @param {string} s
 * @param {number} max
 */
export function truncateForLog(s, max = 120) {
  const t = String(s ?? "");
  if (t.length <= max) return t;
  return `${t.slice(0, max)}…`;
}

/**
 * @param {string} section e.g. "translate" / "llm"
 * @param {string} message
 * @param {unknown} [detail] logged when present; keep small (no full documents by default)
 */
export function debugTrace(section, message, detail) {
  if (!isDebugTraceEnabled()) return;
  const ts = new Date().toISOString();
  const line = `[${section}] ${message}`;
  if (detail !== undefined) {
    console.log(`%c[Legal AI]`, "color:#0b5cab;font-weight:600", ts, line, detail);
  } else {
    console.log(`%c[Legal AI]`, "color:#0b5cab;font-weight:600", ts, line);
  }
  if (footerLineEl) {
    const short = `${section}: ${message}`;
    footerLineEl.textContent = short.length > 240 ? `${short.slice(0, 237)}…` : short;
  }
}

/** @param {string} section @param {unknown} err */
export function debugTraceError(section, err) {
  if (!isDebugTraceEnabled()) return;
  const msg = err instanceof Error ? err.message : String(err);
  debugTrace(section, `ERROR ${msg}`, err instanceof Error ? err.stack : err);
}
