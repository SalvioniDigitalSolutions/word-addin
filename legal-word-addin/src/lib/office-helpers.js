/* global Word, Office */

import { replacementForItem } from "./redact-fakes.js";
import { appLog, appWarn } from "./app-log.js";
import { sanitizeForWordXmlText } from "./translate-biu-ooxml.js";

/**
 * Word for Mac has been observed to surface bogus automation errors (e.g. `'PrintOutOld' is not a property`)
 * when add-ins touch `Document.trackRevisions`. Prefer `changeTrackingMode` only on that platform.
 * @returns {boolean}
 */
function avoidDocumentTrackRevisionsProperty() {
  try {
    return typeof Office !== "undefined" && Office.context?.platform === Office.PlatformType.Mac;
  } catch {
    return false;
  }
}

/** @param {number} ms */
export function delayMs(ms) {
  return new Promise((resolve) => {
    setTimeout(resolve, Math.max(0, ms));
  });
}

/**
 * Pause between each body `insertOoxml` when translating many paragraphs. Rapid back-to-back writes
 * can overload Word (Word Online, coauthoring, Mac) and surface as `forceSaveFailed` or `GeneralException`.
 */
export const OOXML_REPLACE_GAP_MS = 200;

/**
 * Extra pause before writing **body paragraph index 0** (last step when applying last→first).
 * Helps avoid `GeneralException` on `Range.insertOoxml` on some Mac/desktop builds after many prior inserts.
 */
export const OOXML_BEFORE_PARA_INDEX0_MS = 550;

/** Every N successful bulk paragraph writes, add {@link OOXML_BULK_WRITE_EXTRA_MS} so the host can catch up. */
export const OOXML_BULK_WRITE_EVERY = 12;

export const OOXML_BULK_WRITE_EXTRA_MS = 450;

/**
 * Per-insert gap scales slightly with document size (hundreds of paragraphs stress Word more).
 * @param {number} totalParagraphs
 */
export function ooxmlBulkWriteGapAfterInsert(totalParagraphs) {
  const n = Math.max(0, Math.floor(totalParagraphs));
  const base = OOXML_REPLACE_GAP_MS;
  if (n <= 72) return base;
  return base + Math.min(180, Math.floor((n - 72) / 6));
}

/**
 * Extra pause after every {@link OOXML_BULK_WRITE_EVERY} successful inserts during full-document translate.
 * @param {number} completedWriteCount paragraphs successfully written so far in this write phase
 */
export function ooxmlBulkWriteBreatherMs(completedWriteCount) {
  const c = Math.max(0, Math.floor(completedWriteCount));
  if (c > 0 && c % OOXML_BULK_WRITE_EVERY === 0) return OOXML_BULK_WRITE_EXTRA_MS;
  return 0;
}

const INSERT_OOXML_MAX_RETRIES = 5;

/** @param {number} attempt 1-based */
function insertOoxmlRetryDelayMs(attempt) {
  return Math.min(4000, 450 + 650 * attempt * attempt);
}

/**
 * Strip BOM + XML declaration; `Range.insertOoxml` is more reliable with a bare fragment.
 * @param {string} ooxml
 */
export function normalizeOoxmlFragmentForInsert(ooxml) {
  return String(ooxml ?? "")
    .replace(/^\uFEFF?/, "")
    .replace(/^<\?xml[^?]*\?>\s*/i, "")
    .trim();
}

/** WordprocessingML namespace for `w:p` body paragraph fragments. */
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/**
 * Parse, round-trip serialize, and reject known-bad patterns before `insertOoxml`.
 * @returns {{ ok: true } | { ok: false, reason: string }}
 */
export function validateOoxmlFragmentWellFormed(ooxml) {
  const clean = normalizeOoxmlFragmentForInsert(ooxml);
  if (!clean) return { ok: false, reason: "empty fragment" };
  const d = new DOMParser().parseFromString(clean, "application/xml");
  if (d.getElementsByTagName("parsererror").length) {
    return { ok: false, reason: "not well-formed XML" };
  }
  const root = d.documentElement;
  if (!root) return { ok: false, reason: "no root element" };
  const serialized = new XMLSerializer().serializeToString(root);
  if (/\]\]\>/.test(serialized)) {
    return { ok: false, reason: "illegal ]]> in serialized XML (invalid in character data)" };
  }
  const d2 = new DOMParser().parseFromString(serialized, "application/xml");
  if (d2.getElementsByTagName("parsererror").length) {
    return { ok: false, reason: "round-trip serialize/re-parse failed" };
  }
  return { ok: true };
}

/**
 * Validates a single body paragraph package: must be one `w:p` … `</w:p>`.
 * @returns {{ ok: true } | { ok: false, reason: string }}
 */
export function validateBodyParagraphOoxmlForInsert(ooxmlParagraph) {
  const base = validateOoxmlFragmentWellFormed(ooxmlParagraph);
  if (!base.ok) return base;
  const clean = normalizeOoxmlFragmentForInsert(ooxmlParagraph);
  const d = new DOMParser().parseFromString(clean, "application/xml");
  const root = /** @type {Element | null} */ (d.documentElement);
  if (!root || root.namespaceURI !== W_NS || root.localName !== "p") {
    return { ok: false, reason: "root must be a single w:p element" };
  }
  return { ok: true };
}

/** @param {unknown} err */
function isRetriableInsertOoxmlError(err) {
  if (err == null) return false;
  const o = /** @type {Record<string, unknown>} */ (err);
  if (/** @type {{ name?: string }} */ (err).name === "AbortError") return false;
  const code = String(o.code ?? "");
  if (/generalException/i.test(code)) return true;
  const di = o.debugInfo;
  if (di && typeof di === "object") {
    const d = /** @type {Record<string, unknown>} */ (di);
    if (/generalException/i.test(String(d.code ?? ""))) return true;
    if (/insertOoxml/i.test(String(d.errorLocation ?? ""))) return true;
  }
  const msg = String(o.message ?? err);
  return /GeneralException/i.test(msg) && /insertOoxml/i.test(msg);
}

/**
 * How many paragraphs to fetch per `context.sync()`. A single sync with hundreds of `getOoxml()`
 * payloads can take many minutes (especially Word Online) or time out; smaller batches are usually
 * steadier and often complete sooner in practice.
 */
export const OOXML_READ_BATCH_SIZE = 28;

/** Body paragraphs per **Translate document** chunk (range OOXML + one `insertOoxml`, same pattern as Translate selection). */
export const DOCUMENT_TRANSLATE_BODY_CHUNK_PARAGRAPHS = 12;

/**
 * @returns {Promise<number>} `body.paragraphs.items.length`
 */
export async function getBodyParagraphCount() {
  return Word.run(async (context) => {
    const paras = context.document.body.paragraphs;
    paras.load("items");
    await context.sync();
    return paras.items.length;
  });
}

/**
 * Expand from paragraph `startInclusive` through `endInclusive`, read OOXML, then bookmark that span for async translate + `insertOoxml` (like {@link captureSelectionOoxmlWithBookmark}).
 * @param {number} startInclusive 0-based body index
 * @param {number} endInclusive 0-based body index, ≥ start
 */
export async function captureBodyParagraphRangeOoxmlWithBookmark(startInclusive, endInclusive) {
  return Word.run(async (context) => {
    const paras = context.document.body.paragraphs;
    paras.load("items");
    await context.sync();
    const n = paras.items.length;
    if (n === 0) {
      throw new Error("The document body has no paragraphs.");
    }
    let a = Math.max(0, Math.min(startInclusive, n - 1));
    let b = Math.max(0, Math.min(endInclusive, n - 1));
    if (a > b) {
      const t = a;
      a = b;
      b = t;
    }
    const r0 = paras.items[a].getRange();
    const r1 = paras.items[b].getRange();
    const rng = r0.expandTo(r1);
    rng.load("text");
    const ooxml = rng.getOoxml();
    await context.sync();
    const text = rng.text || "";
    const bookmarkName = makeTranslateSelectionBookmarkName();
    rng.insertBookmark(bookmarkName);
    await context.sync();
    return { bookmarkName, ooxml: ooxml.value, text, startIndex: a, endIndex: b };
  });
}

/**
 * Body paragraphs as OOXML + plain text (same order as `body.paragraphs`).
 * @param {(loaded: number, total: number) => void} [onProgress] `loaded` / `total` paragraph count after each batch.
 * @param {AbortSignal} [signal]
 * @returns {Promise<{ ooxmls: string[], plainTexts: string[] }>}
 */
export async function getBodyParagraphsOoxml(onProgress, signal) {
  return Word.run(async (context) => {
    const paras = context.document.body.paragraphs;
    paras.load("items");
    await context.sync();
    const n = paras.items.length;
    /** @type {string[]} */
    const ooxmls = new Array(n);
    /** @type {string[]} */
    const plainTexts = new Array(n);

    const batch = OOXML_READ_BATCH_SIZE;
    for (let start = 0; start < n; start += batch) {
      if (signal?.aborted) {
        throw new DOMException("Stopped.", "AbortError");
      }
      const end = Math.min(start + batch, n);
      const pending = [];
      for (let i = start; i < end; i++) {
        paras.items[i].load("text");
        pending.push(paras.items[i].getOoxml());
      }
      await context.sync();
      for (let j = 0; j < pending.length; j++) {
        const idx = start + j;
        ooxmls[idx] = pending[j].value;
        plainTexts[idx] = paras.items[idx].text ?? "";
      }
      if (typeof onProgress === "function") {
        onProgress(end, n);
      }
    }

    return { ooxmls, plainTexts };
  });
}

/**
 * OOXML already contains `w:del` / `w:ins`. Host track mode off during insert avoids duplicate visible text.
 * @param {Word.Document} doc
 * @param {Word.Range} range
 * @param {string} ooxml
 * @param {Word.RequestContext} context
 */
async function insertOoxmlReplaceWithEmbeddedTrackOnly(doc, range, ooxml, context) {
  const fragment = normalizeOoxmlFragmentForInsert(ooxml);
  doc.load("changeTrackingMode");
  await context.sync();
  const savedMode = doc.changeTrackingMode;
  const skipTrackRevProp = avoidDocumentTrackRevisionsProperty();

  try {
    if (
      !skipTrackRevProp &&
      typeof Office !== "undefined" &&
      Office.context?.requirements?.isSetSupported("WordApiDesktop", "1.4")
    ) {
      doc.trackRevisions = false;
    }
  } catch {
    /* Word on the web */
  }
  doc.changeTrackingMode = Word.ChangeTrackingMode.off;
  await context.sync();

  range.insertOoxml(fragment, Word.InsertLocation.replace);
  await context.sync();

  doc.changeTrackingMode = savedMode;
  await context.sync();
  try {
    if (
      !skipTrackRevProp &&
      typeof Office !== "undefined" &&
      Office.context?.requirements?.isSetSupported("WordApiDesktop", "1.4")
    ) {
      doc.trackRevisions = true;
    }
  } catch {
    /* Word on the web */
  }
  await context.sync();
}

/**
 * Replace one body paragraph by full `<w:p>…</w:p>` OOXML (tracked revisions may be embedded).
 * @param {number} paragraphIndex 0-based
 * @param {string} ooxmlParagraph
 */
export async function replaceBodyParagraphOoxmlAtIndex(paragraphIndex, ooxmlParagraph) {
  const fragment = normalizeOoxmlFragmentForInsert(ooxmlParagraph);
  const pre = validateBodyParagraphOoxmlForInsert(fragment);
  if (!pre.ok) {
    throw new Error(
      `Paragraph ${paragraphIndex + 1}: ${pre.reason}. Try a smaller section or simplify formatting in that paragraph.`,
    );
  }

  const insertHint =
    "Word rejected this paragraph’s OOXML (insertOoxml). Try: save the document, restart Word, shorten or split a very long paragraph, remove unusual characters, or translate fewer paragraphs at once.";

  /** @type {unknown} */
  let lastErr;
  for (let attempt = 1; attempt <= INSERT_OOXML_MAX_RETRIES; attempt++) {
    appLog("word", "insertOoxml body paragraph attempt", {
      paragraph: paragraphIndex + 1,
      attempt,
      maxAttempts: INSERT_OOXML_MAX_RETRIES,
    });
    try {
      await Word.run(async (context) => {
        await ensureTrackChanges(context.document, context);
        const paras = context.document.body.paragraphs;
        paras.load("items");
        await context.sync();
        const n = paras.items.length;
        if (paragraphIndex < 0 || paragraphIndex >= n) {
          throw new Error(
            `Paragraph ${paragraphIndex + 1} is out of range (${n} in body). The document may have changed during translation — save, undo if needed, and try again.`,
          );
        }
        const rng = paras.items[paragraphIndex].getRange();
        await insertOoxmlReplaceWithEmbeddedTrackOnly(context.document, rng, fragment, context);
      });
      appLog("word", "insertOoxml body paragraph OK", { paragraph: paragraphIndex + 1, attempt });
      return;
    } catch (e) {
      lastErr = e;
      if (typeof e === "object" && e !== null && /** @type {{ name?: string }} */ (e).name === "AbortError") {
        throw e;
      }
      if (!isRetriableInsertOoxmlError(e) || attempt === INSERT_OOXML_MAX_RETRIES) {
        appWarn("word", "insertOoxml body paragraph giving up", {
          paragraph: paragraphIndex + 1,
          attempt,
          error: e instanceof Error ? e.message : String(e),
        });
        break;
      }
      const retryWait = insertOoxmlRetryDelayMs(attempt);
      appWarn("word", "insertOoxml body paragraph retry after error", {
        paragraph: paragraphIndex + 1,
        attempt,
        error: e instanceof Error ? e.message : String(e),
        delayMs: retryWait,
      });
      await delayMs(retryWait);
    }
  }

  const base =
    lastErr instanceof Error
      ? lastErr.message
      : lastErr != null
        ? String(lastErr)
        : "Unknown error";
  throw new Error(`Paragraph ${paragraphIndex + 1}: ${base} ${insertHint}`);
}

/**
 * Replace paragraph content with plain text as a tracked change (no OOXML). Use when `insertOoxml` fails.
 * @param {number} paragraphIndex 0-based body index
 * @param {string} plainText translated plain text (line breaks preserved)
 */
export async function replaceBodyParagraphPlainTextTrackedAtIndex(paragraphIndex, plainText) {
  const safe = sanitizeForWordXmlText(String(plainText ?? ""));
  await Word.run(async (context) => {
    await ensureTrackChanges(context.document, context);
    const paras = context.document.body.paragraphs;
    paras.load("items");
    await context.sync();
    const n = paras.items.length;
    if (paragraphIndex < 0 || paragraphIndex >= n) {
      throw new Error(
        `Paragraph ${paragraphIndex + 1} is out of range (${n} in body). The document may have changed during translation.`,
      );
    }
    const rng = paras.items[paragraphIndex].getRange();
    rng.insertText(safe, Word.InsertLocation.replace);
    await context.sync();
  });
}

/**
 * Try OOXML replace; on failure (invalid fragment or Word `insertOoxml`), apply plain translated text if non-empty.
 * @returns {Promise<"ooxml" | "plain">}
 */
export async function replaceBodyParagraphOoxmlOrPlainFallback(
  paragraphIndex,
  ooxmlParagraph,
  plainTranslatedFallback,
) {
  try {
    await replaceBodyParagraphOoxmlAtIndex(paragraphIndex, ooxmlParagraph);
    return "ooxml";
  } catch (e) {
    if (typeof e === "object" && e !== null && /** @type {{ name?: string }} */ (e).name === "AbortError") {
      throw e;
    }
    const plain = String(plainTranslatedFallback ?? "").trim();
    if (!plain) {
      appWarn("word", "insertOoxml failed; no plain-text fallback (empty preview)", {
        paragraph: paragraphIndex + 1,
        error: e instanceof Error ? e.message : String(e),
      });
      throw e;
    }
    const frag = normalizeOoxmlFragmentForInsert(ooxmlParagraph);
    appWarn("word", "insertOoxml failed; trying plain-text tracked replacement", {
      paragraph: paragraphIndex + 1,
      fragmentChars: frag.length,
      head: frag.slice(0, 96),
      error: e instanceof Error ? e.message : String(e),
    });
    try {
      await replaceBodyParagraphPlainTextTrackedAtIndex(paragraphIndex, plainTranslatedFallback);
      appLog("word", "paragraph updated via plain-text fallback (formatting not preserved)", {
        paragraph: paragraphIndex + 1,
      });
      return "plain";
    } catch (e2) {
      appWarn("word", "plain-text fallback also failed", {
        paragraph: paragraphIndex + 1,
        error: e2 instanceof Error ? e2.message : String(e2),
      });
      throw e2;
    }
  }
}

/** @returns {Promise<string>} */
export async function getSelectedText() {
  return Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("text");
    await context.sync();
    return range.text || "";
  });
}

/**
 * Current selection as OOXML (same shape Word uses for `insertOoxml`) plus plain text.
 * @returns {Promise<{ ooxml: string, text: string }>}
 */
export async function getSelectionOoxml() {
  return Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("text");
    const ooxml = range.getOoxml();
    await context.sync();
    return { ooxml: ooxml.value, text: range.text || "" };
  });
}

/** Word bookmark names: letters/digits/underscore, start with letter, max 40 (WordApi 1.4). */
function makeTranslateSelectionBookmarkName() {
  const raw = `LAI${Date.now().toString(36)}${Math.random().toString(36).slice(2, 10)}`;
  const cleaned = raw.replace(/[^a-zA-Z0-9_]/g, "X");
  let name = cleaned.slice(0, 40);
  if (!/^[a-zA-Z]/.test(name)) name = `L${name}`.slice(0, 40);
  return name;
}

/**
 * Snapshot selection OOXML, then drop a **temporary bookmark** on that range so a later `Word.run`
 * can replace the same span after async work (task pane focus no longer steals selection).
 * Requires WordApi 1.4+. Caller should call {@link tryDeleteBookmark} in `finally` if write fails.
 * @returns {Promise<{ bookmarkName: string, ooxml: string, text: string }>}
 */
export async function captureSelectionOoxmlWithBookmark() {
  return Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("text");
    const ooxml = range.getOoxml();
    await context.sync();
    const text = range.text || "";
    if (!text.trim()) {
      throw new Error("Select text in the document first.");
    }
    const bookmarkName = makeTranslateSelectionBookmarkName();
    range.insertBookmark(bookmarkName);
    await context.sync();
    return { bookmarkName, ooxml: ooxml.value, text };
  });
}

/**
 * Replace the current selection with OOXML (e.g. translated fragment). Tracked revisions embedded in OOXML are preserved.
 * Prefer {@link replaceBookmarkRangeOoxmlTracked} after async gaps so the range is still correct.
 * @param {string} ooxml
 */
export async function replaceSelectionOoxmlTracked(ooxml) {
  await Word.run(async (context) => {
    await ensureTrackChanges(context.document, context);
    const range = context.document.getSelection();
    await insertOoxmlReplaceWithEmbeddedTrackOnly(context.document, range, ooxml, context);
  });
}

/**
 * @param {string} bookmarkName
 * @param {string} ooxml
 */
export async function replaceBookmarkRangeOoxmlTracked(bookmarkName, ooxml) {
  const fragment = normalizeOoxmlFragmentForInsert(ooxml);
  const preInsert = validateOoxmlFragmentWellFormed(fragment);
  if (!preInsert.ok) {
    throw new Error(
      `Translated selection OOXML failed validation (${preInsert.reason}). Try a smaller selection or simplify formatting.`,
    );
  }
  const insertHint =
    "Word rejected the translated OOXML (insertOoxml). Try a smaller selection, save and restart Word, or remove complex content (fields, embedded objects) from the selection.";

  /** @type {unknown} */
  let lastErr;
  for (let attempt = 1; attempt <= INSERT_OOXML_MAX_RETRIES; attempt++) {
    appLog("word", "insertOoxml bookmark range attempt", {
      bookmark: bookmarkName.slice(0, 12),
      attempt,
      maxAttempts: INSERT_OOXML_MAX_RETRIES,
    });
    try {
      await Word.run(async (context) => {
        await ensureTrackChanges(context.document, context);
        const r = context.document.getBookmarkRangeOrNullObject(bookmarkName);
        r.load("isNullObject");
        await context.sync();
        if (r.isNullObject) {
          throw new Error(
            "Could not find the translation anchor in the document. Run Translate selection again without editing the marked text.",
          );
        }
        await insertOoxmlReplaceWithEmbeddedTrackOnly(context.document, r, fragment, context);
      });
      appLog("word", "insertOoxml bookmark range OK", { bookmark: bookmarkName.slice(0, 12), attempt });
      return;
    } catch (e) {
      lastErr = e;
      if (typeof e === "object" && e !== null && /** @type {{ name?: string }} */ (e).name === "AbortError") {
        throw e;
      }
      if (
        e instanceof Error &&
        e.message.includes("Could not find the translation anchor")
      ) {
        throw e;
      }
      if (!isRetriableInsertOoxmlError(e) || attempt === INSERT_OOXML_MAX_RETRIES) {
        appWarn("word", "insertOoxml bookmark range giving up", {
          bookmark: bookmarkName.slice(0, 12),
          attempt,
          error: e instanceof Error ? e.message : String(e),
        });
        break;
      }
      const retryWait = insertOoxmlRetryDelayMs(attempt);
      appWarn("word", "insertOoxml bookmark range retry after error", {
        bookmark: bookmarkName.slice(0, 12),
        attempt,
        error: e instanceof Error ? e.message : String(e),
        delayMs: retryWait,
      });
      await delayMs(retryWait);
    }
  }

  const base =
    lastErr instanceof Error
      ? lastErr.message
      : lastErr != null
        ? String(lastErr)
        : "Unknown error";
  throw new Error(`${base} ${insertHint}`);
}

/**
 * Try OOXML bookmark replace; on failure, plain tracked text if non-empty (sanitized).
 * @returns {Promise<"ooxml" | "plain">}
 */
export async function replaceBookmarkRangeOoxmlOrPlainFallback(bookmarkName, ooxml, plainTranslatedFallback) {
  try {
    await replaceBookmarkRangeOoxmlTracked(bookmarkName, ooxml);
    return "ooxml";
  } catch (e) {
    if (typeof e === "object" && e !== null && /** @type {{ name?: string }} */ (e).name === "AbortError") {
      throw e;
    }
    const plain = String(plainTranslatedFallback ?? "").trim();
    if (!plain) {
      appWarn("word", "Bookmark insertOoxml failed; no plain-text fallback", {
        bookmark: bookmarkName.slice(0, 12),
        error: e instanceof Error ? e.message : String(e),
      });
      throw e;
    }
    appWarn("word", "Bookmark insertOoxml failed; plain-text tracked replacement", {
      bookmark: bookmarkName.slice(0, 12),
      error: e instanceof Error ? e.message : String(e),
    });
    await replaceBookmarkRangeTextTracked(bookmarkName, sanitizeForWordXmlText(plainTranslatedFallback));
    appLog("word", "Bookmark range updated via plain-text fallback", { bookmark: bookmarkName.slice(0, 12) });
    return "plain";
  }
}

/**
 * @param {string} bookmarkName
 * @param {string} newText
 */
export async function replaceBookmarkRangeTextTracked(bookmarkName, newText) {
  await Word.run(async (context) => {
    await ensureTrackChanges(context.document, context);
    const r = context.document.getBookmarkRangeOrNullObject(bookmarkName);
    r.load("isNullObject");
    await context.sync();
    if (r.isNullObject) {
      throw new Error(
        "Could not find the translation anchor in the document. Run Translate selection again without editing the marked text.",
      );
    }
    r.insertText(newText, Word.InsertLocation.replace);
    await context.sync();
  });
}

/**
 * Best-effort cleanup of a temp bookmark (e.g. after failed translate). Ignores missing bookmark.
 * @param {string | null | undefined} bookmarkName
 */
export async function tryDeleteBookmark(bookmarkName) {
  if (!bookmarkName) return;
  try {
    await Word.run(async (context) => {
      context.document.deleteBookmark(bookmarkName);
      await context.sync();
    });
  } catch {
    /* Replaced range often removes the bookmark; host may throw if already gone */
  }
}

/** @returns {Promise<string>} */
export async function getFullDocumentText() {
  return Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();
    return body.text || "";
  });
}

/**
 * Ensure change tracking is on. All write helpers call this first so that
 * nothing ever touches the document outside of review mode.
 * @param {Word.Document} doc - already-loaded document proxy
 * @param {Word.RequestContext} context
 */
async function ensureTrackChanges(doc, context) {
  doc.load("changeTrackingMode");
  await context.sync();
  if (doc.changeTrackingMode !== Word.ChangeTrackingMode.trackAll) {
    doc.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    await context.sync();
  }
  try {
    if (
      !avoidDocumentTrackRevisionsProperty() &&
      typeof Office !== "undefined" &&
      Office.context?.requirements?.isSetSupported("WordApiDesktop", "1.4")
    ) {
      doc.trackRevisions = true;
      await context.sync();
    }
  } catch {
    /* Word on the web: trackRevisions may be unavailable; w:ins wrapper still helps some hosts */
  }
}

/**
 * Replace the current selection with new text.
 * Always applied as a tracked change.
 * @param {string} newText
 */
export async function replaceSelectionTracked(newText) {
  return Word.run(async (context) => {
    await ensureTrackChanges(context.document, context);
    const range = context.document.getSelection();
    range.insertText(newText, Word.InsertLocation.replace);
    await context.sync();
  });
}

/**
 * Replace the entire document body with new text.
 * Always applied as a tracked change.
 * @param {string} newText
 */
export async function replaceBodyTracked(newText) {
  return Word.run(async (context) => {
    await ensureTrackChanges(context.document, context);
    const body = context.document.body;
    body.clear();
    body.insertText(newText, Word.InsertLocation.start);
    await context.sync();
  });
}

/** Normalize text so a matched range can be compared to the expected PII string. */
function normRedactCompare(s) {
  return (s || "")
    .trim()
    .replace(/\s+/g, " ")
    .replace(/[\u2018\u2019\u201B\u2032\u0060]/g, "'")
    .replace(/[\u201C\u201D]/g, '"')
    .toLowerCase();
}

function redactTextMatches(found, expected) {
  return normRedactCompare(found) === normRedactCompare(expected);
}

/**
 * For each entry in `items`, find every occurrence in the given scope and
 * replace it with a random fictitious stand-in (names, companies, addresses, etc.)
 * based on category. Same original+category maps to one stand-in for the whole run.
 * For people (fake mode), tokens are shared: "John Smith", then "Smith" or "John", reuse the same fake surname/given name.
 * Each replacement is its own tracked change.
 * Only replaces when the matched range text exactly equals the PII string
 * (after normalizing spaces and quotes) so Word cannot redact random substrings.
 * Longer strings are processed first.
 * @param {{ original: string, category: string }[]} items
 * @param {'selection' | 'body'} scope
 * @param {(done: number, total: number) => void} [onProgress]
 * @param {'redact' | 'fake'} [replacementMode] `redact` → literal [REDACTED]; `fake` → random fictitious stand-ins per category
 * @param {AbortSignal} [signal]
 * @returns {Promise<number>} total replacements made
 */
export async function redactItems(items, scope, onProgress, replacementMode = "fake", signal) {
  const seen = new Set();
  const normalized = [];
  for (const item of items) {
    const text = item.original?.trim();
    if (!text) continue;
    const key = text.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    normalized.push({ ...item, original: text });
  }

  normalized.sort((a, b) => b.original.length - a.original.length);

  /** @type {Map<string, string>} */
  const fakeCache = new Map();
  /** @type {Map<string, string>} lowercase person token → canonical fake */
  const personTokenMap = new Map();
  const useFake = replacementMode === "fake";

  let totalReplaced = 0;
  let done = 0;

  for (const item of normalized) {
    if (signal?.aborted) {
      throw new DOMException("Stopped.", "AbortError");
    }
    const isPhrase = /\s/.test(item.original);
    const replacement = useFake ? replacementForItem(item, fakeCache, personTokenMap) : "[REDACTED]";
    await Word.run(async (context) => {
      await ensureTrackChanges(context.document, context);

      const searchScope = scope === "selection" ? context.document.getSelection() : context.document.body;
      const results = searchScope.search(item.original, {
        matchCase: false,
        matchWholeWord: !isPhrase,
      });
      results.load("items");
      await context.sync();

      const n = results.items.length;
      for (let i = n - 1; i >= 0; i--) {
        results.items[i].load("text");
      }
      await context.sync();

      for (let i = n - 1; i >= 0; i--) {
        const r = results.items[i];
        if (!redactTextMatches(r.text, item.original)) continue;
        r.insertText(replacement, Word.InsertLocation.replace);
        totalReplaced++;
      }
      await context.sync();
    });

    done++;
    if (onProgress) onProgress(done, normalized.length);
  }

  return totalReplaced;
}
