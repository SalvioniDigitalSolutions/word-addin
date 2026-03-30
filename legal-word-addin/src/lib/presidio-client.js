import { loadSettings } from "./ai-settings.js";

/** @param {unknown} err */
function fetchErrorHint(err, url) {
  const e = err instanceof Error ? err : new Error(String(err));
  const msg = e.message || "";
  /** @type {{ code?: string } | null} */
  const cause = e.cause && typeof e.cause === "object" ? /** @type {{ code?: string }} */ (e.cause) : null;
  const code = cause?.code || "";
  const refused = code === "ECONNREFUSED" || /ECONNREFUSED/i.test(msg);
  const genericFetch = /fetch failed|failed to fetch|load failed|networkerror/i.test(msg);
  let hint = `Could not reach ${url}. `;
  if (refused || genericFetch) {
    hint +=
      "Typical causes: (1) npm run dev:all or npm run server is not running; " +
      "(2) Backend base URL in Settings points to the wrong host; " +
      "(3) the Presidio Python sidecar on port 3549 is down — check the terminal where npm run server runs, run npm run presidio if you use PRESIDIO_DISABLED=1, or fix Python/Presidio install errors.";
  } else {
    hint += msg || "Network error.";
  }
  return hint;
}

/** Matches Presidio `_is_generic_number_noise`: skip tiny digit-only spans mistaken for IDs / accounts. */
function isGenericNumericNoise(original, category) {
  const t = original.trim();
  const digits = t.replace(/\D/g, "");
  const c = String(category || "").toLowerCase();
  if (c === "phone") return digits.length < 7;
  if (c === "bank_account") {
    if (/\b[A-Z]{2}\s*\d{2}[\s\d.-]{12,}\b/i.test(t)) return false;
    if (/\b(iban|bic|swift|account|acct|konto|compte|conto|a\/c|routing)\b/i.test(t)) return false;
    if (digits.length >= 8) return false;
    if (digits.length <= 6 && /^[\d\s.,'\-–—]+$/u.test(t)) return true;
    return false;
  }
  if (c === "id_number") {
    if (digits.length >= 8) return false;
    if (digits.length >= 6 && /[A-Za-z.]/.test(t)) return false;
    if (/^\d{1,4}$/.test(digits) && digits.length <= 4) return true;
    if (digits.length <= 4 && t.length <= 6 && /^[\d\s]+$/u.test(t)) return true;
    return false;
  }
  if (c === "person" || c === "company") {
    const compact = t.replace(/[\s-]/g, "");
    if (/^\d{1,4}$/.test(compact)) return true;
  }
  return false;
}

/**
 * @param {unknown} items
 * @returns {{ original: string, category: string }[]}
 */
export function normalizePiiItems(items) {
  if (!Array.isArray(items)) throw new Error("PII scan did not return a JSON array.");
  const out = [];
  for (const i of items) {
    if (!i || typeof i !== "object") continue;
    const o = /** @type {{ original?: unknown }} */ (i).original;
    if (typeof o !== "string" || !o.trim()) continue;
    const c = /** @type {{ category?: unknown }} */ (i).category;
    const category = typeof c === "string" && c.trim() ? String(c).trim().toLowerCase() : "person";
    const trimmed = o.trim();
    if (isGenericNumericNoise(trimmed, category)) continue;
    out.push({ original: trimmed, category });
  }
  return out;
}

/**
 * @param {import("./ai-settings.js").AppAiSettings} settings
 * @returns {string | null}
 */
function presidioRequestUrl(settings) {
  if (settings.requestMode === "direct") {
    const u = (settings.presidioBackendUrl || "").trim().replace(/\/$/, "");
    if (!u) return null;
    return `${u}/api/legal-ai/presidio-pii`;
  }
  const origin = (settings.proxyUrl || "").trim().replace(/\/$/, "");
  return origin ? `${origin}/api/legal-ai/presidio-pii` : "/api/legal-ai/presidio-pii";
}

/**
 * @param {string} text
 * @param {{ signal?: AbortSignal }} [opts]
 * @returns {Promise<{ original: string, category: string }[]>}
 */
export async function detectPiiItems(text, opts = {}) {
  const settings = loadSettings();
  const url = presidioRequestUrl(settings);
  if (!url) {
    throw new Error(
      "Microsoft Presidio is selected but no Presidio backend URL is set. Enter it under Settings, or use Proxy request mode.",
    );
  }
  let res;
  try {
    res = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text }),
      signal: opts.signal,
    });
  } catch (e) {
    throw new Error(fetchErrorHint(e, url));
  }
  let raw;
  try {
    raw = await res.text();
  } catch (e) {
    throw new Error(fetchErrorHint(e, url));
  }
  if (!res.ok) {
    const htmlNoRoute = /<!doctype/i.test(raw) && /cannot post/i.test(raw);
    let msg = htmlNoRoute
      ? ""
      : raw || `HTTP ${res.status}`;
    if (!htmlNoRoute) {
      try {
        const j = JSON.parse(raw);
        if (j?.message) msg = String(j.message);
        else if (j?.error) msg = String(j.error);
      } catch {
        /* keep msg */
      }
    }
    if (htmlNoRoute) {
      msg =
        "Nothing handled POST /api/legal-ai/presidio-pii at this URL. " +
        "Leave Backend base URL empty and use npm run dev:all, " +
        "or point it at the machine running server/proxy.mjs (same server as /api/legal-ai/complete — not a bare OpenAI-compatible URL). " +
        "Restart npm run server after git pull.";
    }
    throw new Error(msg);
  }
  let data;
  try {
    data = JSON.parse(raw);
  } catch {
    throw new Error("Presidio response was not JSON.");
  }
  return normalizePiiItems(data.items);
}
