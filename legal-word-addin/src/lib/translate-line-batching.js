/**
 * Batch plain-text lines for translation so each API call stays under an input token budget.
 */

import { debugTrace, truncateForLog } from "./debug-trace.js";

/** Conservative chars-per-token for mixed / non-English text. */
const CHARS_PER_TOKEN = 3.2;

/** Max estimated input tokens per request (prompt + JSON payload). */
export const TRANSLATE_MAX_INPUT_TOKENS = 2500;

/** Reserve for system message, instructions, JSON syntax. */
const PROMPT_RESERVE_TOKENS = 400;

/** Estimated JSON / array overhead per string in batch payloads (aligned with group packing). */
export const STRING_JSON_OVERHEAD_TOKENS = 4;

/**
 * @param {string} s
 * @returns {number}
 */
export function estimateTokens(s) {
  if (!s) return 0;
  return Math.max(1, Math.ceil(String(s).length / CHARS_PER_TOKEN));
}

/**
 * @returns {number} budget for user content (strings in JSON array)
 */
export function contentTokenBudget() {
  return Math.max(512, TRANSLATE_MAX_INPUT_TOKENS - PROMPT_RESERVE_TOKENS);
}

/**
 * Per-string cost when packing batches (tokens + JSON overhead).
 * @param {string} s
 */
export function stringEntryCost(s) {
  return estimateTokens(s) + STRING_JSON_OVERHEAD_TOKENS;
}

/**
 * Max estimated tokens per split fragment so one fragment + overhead fits one request row.
 */
export function maxPartTokensForSplit() {
  return Math.max(1, contentTokenBudget() - STRING_JSON_OVERHEAD_TOKENS);
}

/** Default cap on how many pieces one oversized string may be split into (2–3 requests, not dozens). */
export const OVERSIZED_LINE_MAX_PARTS = 3;

/**
 * Split `text` into `n` non-empty segments (2 ≤ n ≤ 3), favoring space boundaries near ideal cuts.
 * @param {string} text
 * @param {number} n
 */
function splitIntoNWordAwareChunks(text, n) {
  const t = String(text);
  const L = t.length;
  if (L === 0) return [];
  n = Math.min(OVERSIZED_LINE_MAX_PARTS, Math.max(2, Math.floor(n)));
  if (n <= 1) return [t];

  /** @type {number[]} cut indices: segment i is [cuts[i], cuts[i+1]) */
  const cuts = [0];
  for (let i = 1; i < n; i++) {
    const ideal = Math.floor((L * i) / n);
    const lo = cuts[i - 1] + 1;
    const hi = L - (n - i);
    let cut = Math.max(lo, Math.min(ideal, hi - 1));
    const spaceBefore = t.lastIndexOf(" ", cut);
    const spaceAfter = t.indexOf(" ", cut);
    if (spaceBefore >= lo - 1 && spaceBefore > cuts[i - 1]) {
      if (Math.abs(spaceBefore + 1 - ideal) <= Math.abs(cut - ideal) + 24) cut = spaceBefore + 1;
    } else if (spaceAfter !== -1 && spaceAfter < hi && spaceAfter >= lo - 1) {
      cut = spaceAfter + 1;
    }
    let next = Math.max(lo, Math.min(cut, hi));
    if (next <= cuts[i - 1]) next = Math.min(hi, cuts[i - 1] + 1);
    cuts.push(next);
  }
  cuts.push(L);

  /** @type {string[]} */
  const parts = [];
  for (let i = 0; i < n; i++) {
    const a = cuts[i];
    const b = cuts[i + 1];
    if (b > a) parts.push(t.slice(a, b));
  }
  return parts.length > 0 ? parts : [t];
}

/**
 * Split one long line into at most {@link OVERSIZED_LINE_MAX_PARTS} parts (minimum 2 when over budget).
 * Avoids dozens of tiny API slices for one paragraph.
 * @param {string} text
 * @param {number} maxTokensPerPart target size hint (used to pick 2 vs 3 parts)
 * @param {number} [maxParts]
 * @returns {string[]}
 */
export function splitOversizedLineMaxParts(text, maxTokensPerPart, maxParts = OVERSIZED_LINE_MAX_PARTS) {
  const t = String(text);
  if (!t) return [];
  if (estimateTokens(t) <= maxTokensPerPart) return [t];
  const limit = Math.min(OVERSIZED_LINE_MAX_PARTS, Math.max(2, Math.floor(maxParts)));
  const totalTok = estimateTokens(t);
  let k = Math.ceil(totalTok / maxTokensPerPart);
  k = Math.min(limit, Math.max(2, k));
  return splitIntoNWordAwareChunks(t, k);
}

/**
 * How many `translateBatch` calls would be made for this ordered list (same packing as {@link translateOrderedStringsInBudget}).
 * @param {string[]} strings
 * @param {number} budget content token budget (same as {@link contentTokenBudget})
 */
export function countApiCallsForOrderedStrings(strings, budget, oversizeDepth = 0) {
  let count = 0;
  let idx = 0;
  while (idx < strings.length) {
    const s = strings[idx];
    const c = stringEntryCost(s);
    if (c > budget) {
      if (oversizeDepth >= 1) {
        count++;
        idx++;
        continue;
      }
      const maxTP = Math.max(1, budget - STRING_JSON_OVERHEAD_TOKENS);
      const parts = splitOversizedLineMaxParts(s, maxTP, OVERSIZED_LINE_MAX_PARTS);
      count += countApiCallsForOrderedStrings(parts, budget, oversizeDepth + 1);
      idx++;
      continue;
    }
    let end = idx;
    let sum = 0;
    while (end < strings.length) {
      const ec = stringEntryCost(strings[end]);
      if (sum + ec > budget) break;
      sum += ec;
      end++;
    }
    count++;
    idx = end;
  }
  return count;
}

/**
 * Calls `translateBatch` one or more times so each call's summed {@link stringEntryCost} is ≤ `budget`.
 * Output length matches `strings.length` (split parts are translated then rejoined per original index).
 * @param {string[]} strings
 * @param {(batch: string[]) => Promise<string[]>} translateBatch
 * @param {number} budget
 * @param {number} [oversizeDepth]
 * @returns {Promise<string[]>}
 */
export async function translateOrderedStringsInBudget(strings, translateBatch, budget, oversizeDepth = 0) {
  /** @type {string[]} */
  const result = [];
  let idx = 0;
  while (idx < strings.length) {
    const s = strings[idx];
    const c = stringEntryCost(s);
    if (c > budget) {
      if (oversizeDepth >= 1) {
        debugTrace("translate.pack", `oversized atom (depth 1): 1 string len=${s.length} — single batch`);
        const trans = await translateBatch([s]);
        if (trans.length !== 1) {
          throw new Error(`Translation batch length mismatch: expected 1, got ${trans.length}.`);
        }
        result.push(String(trans[0] ?? ""));
        idx++;
        continue;
      }
      const maxTP = Math.max(1, budget - STRING_JSON_OVERHEAD_TOKENS);
      const parts = splitOversizedLineMaxParts(s, maxTP, OVERSIZED_LINE_MAX_PARTS);
      const transParts = await translateOrderedStringsInBudget(parts, translateBatch, budget, oversizeDepth + 1);
      result.push(transParts.join(""));
      idx++;
      continue;
    }
    let end = idx;
    let sum = 0;
    while (end < strings.length) {
      const ec = stringEntryCost(strings[end]);
      if (sum + ec > budget) break;
      sum += ec;
      end++;
    }
    const batch = strings.slice(idx, end);
    const estCost = batch.reduce((a, s) => a + stringEntryCost(s), 0);
    debugTrace(
      "translate.pack",
      `sub-batch ${batch.length} string(s), ~${estCost} est. content tokens (budget ${budget})`,
      batch.map((s, i) => ({ i, len: s.length, head: truncateForLog(s, 72) })),
    );
    const trans = await translateBatch(batch);
    if (trans.length !== batch.length) {
      throw new Error(`Translation batch length mismatch: expected ${batch.length}, got ${trans.length}.`);
    }
    for (let j = 0; j < batch.length; j++) {
      result.push(String(trans[j] ?? ""));
    }
    idx = end;
  }
  return result;
}
