/**
 * Turn thrown values into text the task pane can show. Office/Word often surfaces
 * `Script error.` with no stack in the embedded webview.
 * @param {unknown} err
 * @returns {string}
 */
export function formatAddinError(err) {
  if (err == null) return "Unknown error.";
  if (typeof err === "object" && err !== null && /** @type {{ name?: string }} */ (err).name === "AbortError") {
    return "Stopped.";
  }
  if (typeof err === "string") {
    if (err === "Script error." || err === "Script error") {
      return (
        "Word reported a generic script error (details hidden by the host). " +
        "Often fixed by: smaller selection, Translate document instead of selection, desktop Word, or updating Word. " +
        "If you use dev server, check the terminal / browser devtools on the machine running webpack."
      );
    }
    return err;
  }
  const o = /** @type {Record<string, unknown>} */ (err);
  const msg = typeof o.message === "string" ? o.message : "";
  if (msg === "Script error." || msg === "Script error") {
    return formatAddinError("Script error.");
  }
  const debugInfo = o.debugInfo;
  if (debugInfo && typeof debugInfo === "object") {
    const di = /** @type {Record<string, unknown>} */ (debugInfo);
    const code = di.code != null ? String(di.code) : "";
    const inner =
      typeof di.errorLocation === "string"
        ? di.errorLocation
        : typeof di.fullStatements === "string"
          ? di.fullStatements
          : "";
    if (inner) {
      return [msg || "Word API error", code && `Code: ${code}`, inner].filter(Boolean).join(" — ");
    }
  }
  if (typeof o.code === "string" || typeof o.code === "number") {
    const trace = typeof o.traceMessages === "string" ? o.traceMessages : "";
    if (msg) return trace ? `${msg} (${trace})` : `${msg} (code: ${o.code})`;
  }
  if (msg) return msg;
  try {
    return JSON.stringify(o).slice(0, 800);
  } catch {
    return String(err);
  }
}
