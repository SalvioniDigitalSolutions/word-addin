/**
 * Single "current run" token for Tools / Chat long operations. Stop button calls {@link cancelActiveRun}.
 */

/** @type {AbortController | null} */
let active = null;

/** Start a cancellable run; replaces any previous controller. @returns {AbortSignal} */
export function beginRun() {
  active = new AbortController();
  return active.signal;
}

/** Clear after a run finishes normally (does not abort). */
export function endRun() {
  active = null;
}

/** User clicked Stop — aborts in-flight fetch and makes checks throw. */
export function cancelActiveRun() {
  active?.abort();
}

/** @returns {AbortSignal | undefined} */
export function getActiveRunSignal() {
  return active?.signal;
}

export function throwIfRunCancelled() {
  const s = active?.signal;
  if (s?.aborted) {
    throw new DOMException("Stopped.", "AbortError");
  }
}
