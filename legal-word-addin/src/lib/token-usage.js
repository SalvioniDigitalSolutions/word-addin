/**
 * Persists LLM token totals per calendar month (local timezone).
 * Uses provider `usage` when present; otherwise a rough estimate from message + reply size.
 */

import { estimateTokens } from "./translate-line-batching.js";

const STORAGE_KEY = "legal-word-addin:llm-token-month-v1";

/** @returns {string} YYYY-MM */
export function currentMonthKey() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

/** @returns {Record<string, number>} */
function loadStore() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return {};
    const o = JSON.parse(raw);
    if (o && typeof o === "object") return /** @type {Record<string, number>} */ (o);
  } catch {
    /* ignore */
  }
  return {};
}

/** @param {Record<string, number>} store */
function saveStore(store) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(store));
  } catch {
    /* quota */
  }
}

/** @param {unknown} v */
function toNonNegInt(v) {
  if (typeof v === "number" && Number.isFinite(v) && v >= 0) return Math.floor(v);
  if (typeof v === "string" && v.trim()) {
    const n = Number(v);
    if (Number.isFinite(n) && n >= 0) return Math.floor(n);
  }
  return null;
}

/** @param {unknown} u */
function billableTokensFromUsageObject(u) {
  if (!u || typeof u !== "object") return 0;
  const usage = /** @type {Record<string, unknown>} */ (u);
  const total = toNonNegInt(usage.total_tokens ?? usage.totalTokens);
  if (total != null && total > 0) return total;
  const input = toNonNegInt(
    usage.prompt_tokens ?? usage.promptTokens ?? usage.input_tokens ?? usage.inputTokens,
  );
  const output = toNonNegInt(
    usage.completion_tokens ?? usage.completionTokens ?? usage.output_tokens ?? usage.outputTokens,
  );
  const pi = input ?? 0;
  const po = output ?? 0;
  if (pi === 0 && po === 0) return 0;
  return pi + po;
}

/**
 * @param {unknown} data Parsed JSON body (OpenAI-style, Anthropic messages, etc.)
 * @returns {number}
 */
export function extractBillableTokensFromResponse(data) {
  if (!data || typeof data !== "object") return 0;
  const d = /** @type {Record<string, unknown>} */ (data);

  let n = billableTokensFromUsageObject(d.usage);
  if (n > 0) return n;

  const inner = d.data;
  if (inner && typeof inner === "object") {
    n = billableTokensFromUsageObject(/** @type {Record<string, unknown>} */ (inner).usage);
    if (n > 0) return n;
  }

  const pt = toNonNegInt(d.prompt_tokens);
  const ct = toNonNegInt(d.completion_tokens);
  if (pt != null || ct != null) {
    return (pt ?? 0) + (ct ?? 0);
  }

  return 0;
}

/**
 * Rough token count for one chat turn when the API omits usage (same estimator as batching).
 * @param {{ role: string, content: string }[]} messages
 * @param {string} assistantText
 */
export function estimateChatTurnTokens(messages, assistantText) {
  let input = 0;
  if (Array.isArray(messages)) {
    for (const m of messages) {
      if (m && typeof m.content === "string") input += estimateTokens(m.content);
    }
  }
  const output = estimateTokens(String(assistantText ?? ""));
  const sum = input + output;
  return sum > 0 ? sum : 0;
}

/** @param {number} n */
function addTokensToMonth(n) {
  if (!Number.isFinite(n) || n <= 0) return;
  const key = currentMonthKey();
  const store = loadStore();
  store[key] = (store[key] || 0) + n;
  saveStore(store);
  try {
    window.dispatchEvent(
      new CustomEvent("legal-ai-tokens-updated", { detail: { month: key, total: store[key] } }),
    );
  } catch {
    /* no window */
  }
}

/**
 * After a successful chat completion: record provider usage if present, else client estimate.
 * @param {unknown} data Parsed JSON response body
 * @param {{ role: string, content: string }[]} messages Request messages
 * @param {string} assistantText Model reply text
 */
export function recordChatTurnUsage(data, messages, assistantText) {
  const fromApi = extractBillableTokensFromResponse(data);
  if (fromApi > 0) {
    addTokensToMonth(fromApi);
    return;
  }
  const est = estimateChatTurnTokens(messages, assistantText);
  if (est > 0) {
    addTokensToMonth(est);
  }
}

/** Add tokens from one successful API response (no-op if usage missing). */
export function recordUsageFromApiResponse(data) {
  const n = extractBillableTokensFromResponse(data);
  if (n <= 0) return;
  addTokensToMonth(n);
}

/** @returns {number} */
export function getTokensRecordedThisMonth() {
  const key = currentMonthKey();
  return Math.floor(loadStore()[key] || 0);
}

/** Short label for footer (e.g. "Mar 2026"). */
export function currentMonthLabel() {
  return new Date().toLocaleString(undefined, { month: "short", year: "numeric" });
}
