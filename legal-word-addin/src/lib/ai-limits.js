/**
 * Some OpenAI-compatible gateways reject max_tokens above 5000.
 * All completions are clamped to this ceiling.
 */
export const MAX_COMPLETION_TOKENS = 5000;

/**
 * @param {number | undefined} requested
 * @returns {number} integer in [1, MAX_COMPLETION_TOKENS]
 */
export function clampMaxTokens(requested) {
  const n = Math.floor(Number(requested));
  const base = Number.isFinite(n) && n >= 1 ? n : 4096;
  return Math.min(base, MAX_COMPLETION_TOKENS);
}
