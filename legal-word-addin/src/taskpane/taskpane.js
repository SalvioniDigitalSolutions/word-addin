import "./taskpane.css";
import { formatAddinError } from "../lib/error-format.js";
import { loadSettings, saveSettings, defaultSettings } from "../lib/ai-settings.js";
import { completeChat } from "../lib/ai-client.js";
import { beginRun, endRun, cancelActiveRun, getActiveRunSignal } from "../lib/process-cancel.js";
import { getTokensRecordedThisMonth, currentMonthLabel } from "../lib/token-usage.js";
import {
  debugTrace,
  debugTraceError,
  initDebugTraceFooter,
  isDebugTraceEnabled,
  setDebugTraceEnabled,
} from "../lib/debug-trace.js";
import { clampMaxTokens } from "../lib/ai-limits.js";
import { appLog, appWarn } from "../lib/app-log.js";
import {
  getSelectedText,
  getFullDocumentText,
  getBodyParagraphCount,
  captureBodyParagraphRangeOoxmlWithBookmark,
  captureSelectionOoxmlWithBookmark,
  DOCUMENT_TRANSLATE_BODY_CHUNK_PARAGRAPHS,
  delayMs,
  OOXML_REPLACE_GAP_MS,
  replaceBookmarkRangeOoxmlTracked,
  replaceBookmarkRangeOoxmlOrPlainFallback,
  replaceBookmarkRangeTextTracked,
  replaceSelectionTracked,
  tryDeleteBookmark,
  replaceBodyTracked,
  redactItems,
} from "../lib/office-helpers.js";
import * as prompts from "../lib/prompts.js";
import { detectPiiItems } from "../lib/presidio-client.js";
import {
  translateEveryParagraphInOoxmlTree,
  paragraphPlainPreviewFromOoxml,
  paragraphTranslatedPlainFromMergedOoxml,
  countTranslatableLinesInOoxmlTree,
  serializeParagraph,
  sanitizeForWordXmlText,
  W_NS,
} from "../lib/translate-biu-ooxml.js";

/** @param {string} raw */
function parseModelJsonStringArray(raw) {
  let t = String(raw).trim();
  t = t.replace(/^```(?:json)?\s*/i, "").replace(/\s*```\s*$/i, "");
  const arr = JSON.parse(t);
  if (!Array.isArray(arr)) throw new Error("Expected a JSON array");
  return arr.map((x) => (x == null ? "" : String(x)));
}

/** @param {unknown} raw */
function normalizeTranslationChunk(raw) {
  return sanitizeForWordXmlText(
    String(raw ?? "")
      .replace(/\r\n/g, "\n")
      .replace(/^["'`]+|["'`]+$/g, "")
      .trimEnd(),
  );
}

/**
 * @param {string} outOoxml
 * @param {string[]} previewLines
 * @param {number} startPara 0-based body index of first `w:p` in chunk
 * @param {number} n body paragraph count
 */
function enrichPreviewLinesFromTranslatedOoxml(outOoxml, previewLines, startPara, n) {
  const clean = String(outOoxml || "")
    .replace(/^\uFEFF?/, "")
    .replace(/^<\?xml[^?]*\?>\s*/i, "")
    .trim();
  const dom = new DOMParser().parseFromString(clean, "application/xml");
  if (dom.getElementsByTagName("parsererror").length) return;
  const ps = [...dom.getElementsByTagNameNS(W_NS, "p")];
  for (let i = 0; i < ps.length; i++) {
    const gi = startPara + i;
    if (gi >= 0 && gi < n) {
      previewLines[gi] =
        paragraphTranslatedPlainFromMergedOoxml(serializeParagraph(ps[i])) || previewLines[gi] || "";
    }
  }
}

/** Plain-text fallback for a translated chunk (for bookmark `insertOoxml` failure). */
function translatedPlainJoinFromOoxmlRoot(outOoxml) {
  const clean = String(outOoxml || "")
    .replace(/^\uFEFF?/, "")
    .replace(/^<\?xml[^?]*\?>\s*/i, "")
    .trim();
  const dom = new DOMParser().parseFromString(clean, "application/xml");
  if (dom.getElementsByTagName("parsererror").length) return "";
  return [...dom.getElementsByTagNameNS(W_NS, "p")]
    .map((p) => paragraphTranslatedPlainFromMergedOoxml(serializeParagraph(p)))
    .filter((x) => x != null && String(x).length > 0)
    .join("\n\n");
}

/**
 * Batch translate (≤ ~2.5k input tokens per API call via prompts).
 * If JSON batch fails or length mismatches, **bisects** the segment list (halves) instead of one LLM call per string.
 * @param {string[]} lines
 * @param {string} targetLang
 * @param {AbortSignal | undefined} signal
 */
async function translateLinesBatchWithFallback(lines, targetLang, signal) {
  if (lines.length === 0) return [];
  appLog("translate", "translateLinesBatchWithFallback start", {
    segments: lines.length,
    targetLang,
  });

  /**
   * @param {string[]} batch
   * @param {number} depth
   * @returns {Promise<string[]>}
   */
  async function recurse(batch, depth) {
    if (signal?.aborted) {
      throw new DOMException("Stopped.", "AbortError");
    }
    if (batch.length === 0) return [];
    if (batch.length === 1) {
      appLog("translate", "single-segment translate", { depth, textLen: batch[0].length });
      debugTrace("translate.fallback", `single-segment translate, len=${batch[0].length}, depth=${depth}`);
      const msgs = prompts.translateSegmentPrompt(batch[0], targetLang);
      const raw = await completeChat(msgs, { maxTokens: 2048, signal });
      return [normalizeTranslationChunk(raw)];
    }

    const approxOut = Math.ceil(batch.reduce((a, s) => a + String(s).length, 0) * 0.45) + 400;
    const maxTok = clampMaxTokens(Math.min(5000, Math.max(1024, approxOut)));

    debugTrace(
      "translate.batch",
      `JSON batch try: ${batch.length} segment(s) → ${targetLang}, maxTokens=${maxTok}, depth=${depth}`,
      batch.map((s, i) => ({ i, len: String(s).length })),
    );

    try {
      const msgs = prompts.translateBatchPrompt(batch, targetLang);
      const raw = await completeChat(msgs, { maxTokens: maxTok, signal });
      const out = parseModelJsonStringArray(raw).map(normalizeTranslationChunk);
      if (out.length === batch.length) {
        appLog("translate", "JSON batch OK", { segments: out.length, depth });
        debugTrace("translate.batch", `JSON batch OK: ${out.length} segment(s), depth=${depth}`);
        return out;
      }
      appWarn("translate", "JSON length mismatch → bisect", {
        expected: batch.length,
        got: out.length,
        depth,
      });
      debugTrace(
        "translate.batch",
        `JSON length mismatch (${out.length} vs ${batch.length}), bisecting`,
      );
    } catch (e) {
      if (typeof e === "object" && e !== null && /** @type {{ name?: string }} */ (e).name === "AbortError") {
        throw e;
      }
      appWarn("translate", "JSON batch error → bisect", {
        segments: batch.length,
        depth,
        error: e instanceof Error ? e.message : String(e),
      });
      debugTraceError("translate.batch", e);
      debugTrace("translate.batch", `JSON batch error — bisecting (${batch.length} segment(s))`);
    }

    const mid = Math.floor(batch.length / 2);
    const cut = mid > 0 ? mid : 1;
    const left = batch.slice(0, cut);
    const right = batch.slice(cut);
    const tLeft = await recurse(left, depth + 1);
    const tRight = await recurse(right, depth + 1);
    const merged = tLeft.concat(tRight);
    if (merged.length !== batch.length) {
      throw new Error(
        `Translation bisect merge bug: expected ${batch.length} segment(s), got ${merged.length}.`,
      );
    }
    return merged;
  }

  return recurse(lines, 0);
}

/** @param {HTMLButtonElement | null} btn */
function translateProgressStart(btn) {
  if (!btn) return;
  btn.classList.add("is-busy");
  translateProgressSet(btn, 0.5);
}

/** @param {HTMLButtonElement | null} btn @param {number} pct 0–100 */
function translateProgressSet(btn, pct) {
  if (!btn) return;
  const fill = btn.querySelector(".btn-translate__progress");
  if (!fill) return;
  const p = Math.min(100, Math.max(0, pct)) / 100;
  fill.style.transform = `scaleX(${Math.max(0.008, p)})`;
}

/** Write phase uses the last 20% of overall progress (after merge ~80%). */
const TRANSLATE_WRITE_PHASE_START_PCT = 80;

/**
 * Document write applies paragraphs last→first; button fill uses **right** origin so growth moves
 * leftward through this phase (matches Word order). Overall fill goes 80%→100% as paragraphs complete.
 * @param {HTMLButtonElement | null} btn
 * @param {number} completedCount paragraphs written or skipped this phase (0…n)
 * @param {number} n total body paragraphs
 */
function translateProgressSetWriteReverse(btn, completedCount, n) {
  if (!btn || n <= 0) return;
  const fill = btn.querySelector(".btn-translate__progress");
  if (!fill) return;
  const span = 100 - TRANSLATE_WRITE_PHASE_START_PCT;
  const pct = TRANSLATE_WRITE_PHASE_START_PCT + (span * Math.min(n, Math.max(0, completedCount))) / n;
  const t = Math.min(1, Math.max(0, pct / 100));
  fill.style.transform = `scaleX(${Math.max(0.008, t)})`;
}

/** @param {HTMLButtonElement | null} btn @param {number} n total body paragraphs */
function translateProgressEnterWritePhase(btn, n) {
  if (!btn) return;
  btn.classList.add("btn-translate--write-reverse");
  translateProgressSetWriteReverse(btn, 0, n);
}

/** @param {HTMLButtonElement | null} btn */
function translateProgressEndWritePhase(btn) {
  if (!btn) return;
  btn.classList.remove("btn-translate--write-reverse");
  translateProgressSet(btn, 100);
}

/** @param {HTMLButtonElement | null} btn */
function translateProgressDone(btn) {
  if (!btn) return;
  btn.classList.remove("is-busy");
  btn.classList.remove("btn-translate--write-reverse");
  const fill = btn.querySelector(".btn-translate__progress");
  if (fill) {
    fill.style.transform = "scaleX(0)";
  }
}

/**
 * Document translate: body is processed in **chunks** of {@link DOCUMENT_TRANSLATE_BODY_CHUNK_PARAGRAPHS} paragraphs.
 * Each chunk uses range OOXML + bookmark → `translateEveryParagraphInOoxmlTree` → **one** `insertOoxml` (same idea as Translate selection).
 * Chunks are applied **last → first** so paragraph indices stay valid after each write.
 * @param {HTMLButtonElement | null} translateBtn
 */
async function runTranslate(translateBtn) {
  const status = $("#status");
  const preview = $("#preview");
  const targetLang = $("#target-language").value;

  setAllButtons(true);
  beginRun();
  const signal = getActiveRunSignal();
  translateProgressStart(translateBtn);
  progressStart("Translate document (chunked)");
  progressSet(0, "Preparing document chunks…");
  preview.value = "";
  setStatus(status, "Preparing document translation…");
  appLog("translate", "Translate document started", { targetLang });
  debugTrace("translate", `runTranslate (document): target=${targetLang}`);

  try {
    const n = await getBodyParagraphCount();
    if (n === 0) {
      setStatus(status, "The document body has no paragraphs.", true);
      progressDone(true);
      return;
    }

    const CHUNK = DOCUMENT_TRANSLATE_BODY_CHUNK_PARAGRAPHS;
    /** @type {{ start: number, end: number }[]} */
    const chunks = [];
    for (let s = 0; s < n; s += CHUNK) {
      chunks.push({ start: s, end: Math.min(s + CHUNK - 1, n - 1) });
    }
    chunks.reverse();

    appLog("translate", "Document translate: chunked bookmark + range OOXML (same write path as selection)", {
      paragraphs: n,
      chunks: chunks.length,
      maxParagraphsPerChunk: CHUNK,
    });
    debugTrace(
      "translate",
      `document chunked: ${n} para(s), ${chunks.length} chunk(s), chunks applied last→first to keep indices stable`,
    );

    /** @type {string[]} */
    const previewLines = new Array(n);
    previewLines.fill("");

    /** @type {number[]} */
    const skippedChunks1Based = [];
    /** @type {number[]} */
    const plainFallbackChunks1Based = [];

    progressSet(6, `Document · ${n} paragraphs · ${chunks.length} chunk(s)`);
    translateProgressSet(translateBtn, 6);
    progressMarkWritePhase(chunks.length);

    for (let ci = 0; ci < chunks.length; ci++) {
      const { start, end } = chunks[ci];
      if (signal?.aborted) {
        throw new DOMException("Stopped.", "AbortError");
      }

      setStatus(
        status,
        `Chunk ${ci + 1}/${chunks.length}: reading paragraphs ${start + 1}–${end + 1} of ${n}…`,
      );
      progressSet(8 + (ci / Math.max(1, chunks.length)) * 4, `Reading chunk ${ci + 1}/${chunks.length}`);
      translateProgressSet(translateBtn, 8 + (ci / Math.max(1, chunks.length)) * 4);

      const cap = await captureBodyParagraphRangeOoxmlWithBookmark(start, end);
      const bm = cap.bookmarkName;
      let chunkLinesEstimate = 1;
      let doneLinesInChunk = 0;

      try {
        const linesInChunk = countTranslatableLinesInOoxmlTree(cap.ooxml);
        chunkLinesEstimate = Math.max(1, linesInChunk);

        if (linesInChunk === 0) {
          appLog("translate", "chunk: no translatable lines", { start: start + 1, end: end + 1 });
        } else {
          const outOoxml = await translateEveryParagraphInOoxmlTree(cap.ooxml, {
            signal,
            onReadyToTranslate: (nLines, nBatches) => {
              appLog("translate", "chunk AI ready", {
                chunk: ci + 1,
                lines: nLines,
                apiCalls: nBatches,
              });
              if (nBatches > 0) {
                progressMarkAiBatches(nBatches);
              }
            },
            onBatchStart: (i, nt, segCount) => {
              appLog("translate", `chunk ${ci + 1}/${chunks.length} · AI ${i}/${nt}`, {
                segmentsInRequest: segCount,
              });
              setStatus(
                status,
                `Chunk ${ci + 1}/${chunks.length} · AI ${i}/${nt} — ${segCount} segment(s)…`,
              );
              progressPhase(`Chunk ${ci + 1}/${chunks.length} · AI ${i}/${nt} · waiting…`);
            },
            onBatchComplete: () => {
              progressMarkAiBatchFinished();
            },
            onTranslatedLine: (_t, meta) => {
              doneLinesInChunk++;
              const lineFrac = Math.min(1, doneLinesInChunk / chunkLinesEstimate);
              const chunkFrac = (ci + lineFrac * 0.65) / Math.max(1, chunks.length);
              const pct = 12 + chunkFrac * 72;
              translateProgressSet(translateBtn, pct);
              const globalPara =
                meta != null && typeof meta.paraIndex === "number"
                  ? start + meta.paraIndex + 1
                  : null;
              progressSet(
                pct,
                globalPara != null
                  ? `Chunk ${ci + 1}/${chunks.length} · line ${doneLinesInChunk} · para ${globalPara}/${n}`
                  : `Chunk ${ci + 1}/${chunks.length} · translating…`,
              );
              setStatus(
                status,
                globalPara != null
                  ? `Chunk ${ci + 1}/${chunks.length} · paragraph ${globalPara} · line ${doneLinesInChunk}…`
                  : `Chunk ${ci + 1}/${chunks.length} · translating…`,
              );
            },
            onMergeProgress: (done, total) => {
              if (done === 1) {
                progressTimingClearEta();
              }
              const mergeFrac = 0.65 + (done / total) * 0.35;
              const chunkFrac = (ci + mergeFrac) / Math.max(1, chunks.length);
              const pct = 12 + chunkFrac * 72;
              progressSet(pct, `Chunk ${ci + 1}/${chunks.length} · merge ${done}/${total}`);
              translateProgressSet(translateBtn, pct);
              setStatus(status, `Chunk ${ci + 1}/${chunks.length}: merging OOXML ${done}/${total}…`);
            },
            translateBatch: (batch) => translateLinesBatchWithFallback(batch, targetLang, signal),
          });

          if (outOoxml == null) {
            setStatus(status, `Chunk ${ci + 1}/${chunks.length}: plain text (no w:p in range)…`);
            const msgs = prompts.translateSegmentPrompt(cap.text, targetLang);
            const raw = await completeChat(msgs, { maxTokens: 8192, signal });
            const translated = normalizeTranslationChunk(raw);
            await replaceBookmarkRangeTextTracked(bm, translated);
            previewLines[start] = translated;
          } else {
            enrichPreviewLinesFromTranslatedOoxml(outOoxml, previewLines, start, n);
            const plainFb = translatedPlainJoinFromOoxmlRoot(outOoxml);
            const mode = await replaceBookmarkRangeOoxmlOrPlainFallback(bm, outOoxml, plainFb);
            if (mode === "plain") {
              plainFallbackChunks1Based.push(ci + 1);
            }
          }
        }

        progressMarkWriteOneDone();
      } catch (e) {
        if (typeof e === "object" && e !== null && /** @type {{ name?: string }} */ (e).name === "AbortError") {
          throw e;
        }
        appWarn(
          "translate",
          `Chunk ${ci + 1}/${chunks.length} (paragraphs ${start + 1}–${end + 1}) failed`,
          e instanceof Error ? e.message : String(e),
        );
        skippedChunks1Based.push(ci + 1);
      } finally {
        await tryDeleteBookmark(bm);
      }

      const afterPct = Math.min(98, 12 + ((ci + 1) / Math.max(1, chunks.length)) * 84);
      translateProgressSet(translateBtn, afterPct);
      progressSet(afterPct, `Chunk ${ci + 1}/${chunks.length} done`);
      if (ci < chunks.length - 1) {
        await delayMs(OOXML_REPLACE_GAP_MS + 120);
      }
    }

    setLivePreview(previewLines.filter((x) => x != null && x !== "").join("\n\n"), preview);

    translateProgressSet(translateBtn, 100);
    progressDone(false);

    const skipN = skippedChunks1Based.length;
    const skipList =
      skipN === 0
        ? ""
        : skipN <= 12
          ? skippedChunks1Based.join(", ")
          : `${skippedChunks1Based.slice(0, 10).join(", ")}… (+${skipN - 10} more)`;
    const allSkipped = skipN > 0 && skipN === chunks.length;
    const fbN = plainFallbackChunks1Based.length;
    const fbNote =
      fbN === 0
        ? ""
        : fbN <= 8
          ? ` ${fbN} chunk(s) used plain text after Word rejected OOXML: ${plainFallbackChunks1Based.join(", ")}.`
          : ` ${fbN} chunk(s) used plain text after Word rejected OOXML.`;
    const baseDone = `Translated ${n} body paragraph(s) to ${targetLang} in ${chunks.length} chunk(s) (up to ${CHUNK} paragraphs each, same insert pattern as Translate selection). Each API call stays within ~2.5k input tokens. Review tracked changes in Word.${fbNote}`;
    if (allSkipped) {
      appWarn("translate", "Translate document — every chunk failed", { chunks: chunks.length });
      setStatus(
        status,
        `No part of the document was updated (${chunks.length} chunk(s) failed). See console [Legal AI] logs. Try a smaller document, save and restart Word, or use Translate selection on sections.`,
        true,
      );
    } else if (skipN > 0) {
      setStatus(
        status,
        `${baseDone} Failed chunk(s) (paragraph ranges unchanged): ${skipList}.`,
      );
      appLog("translate", "Translate document finished (some chunks skipped)", {
        skippedChunks: skippedChunks1Based,
        plainFallbackChunks: plainFallbackChunks1Based,
      });
    } else {
      setStatus(status, baseDone);
      appLog("translate", "Translate document finished OK", {
        plainFallbackChunks: plainFallbackChunks1Based,
      });
    }
    debugTrace("translate", `runTranslate (document): finished OK`);
  } catch (e) {
    progressDone(true);
    appWarn("translate", "Translate document failed", e instanceof Error ? e.message : String(e));
    debugTraceError("translate", e);
    const msg = formatAddinError(e);
    setStatus(status, msg, true);
    preview.value = msg;
  } finally {
    endRun();
    translateProgressDone(translateBtn);
    setAllButtons(false);
  }
}

/** @param {string} ooxml */
function previewPlainFromOoxmlRoot(ooxml) {
  const clean = String(ooxml || "")
    .replace(/^\uFEFF?/, "")
    .replace(/^<\?xml[^?]*\?>\s*/i, "")
    .trim();
  const dom = new DOMParser().parseFromString(clean, "application/xml");
  if (dom.getElementsByTagName("parsererror").length) return "";
  return [...dom.getElementsByTagNameNS(W_NS, "p")]
    .map((p) => paragraphPlainPreviewFromOoxml(serializeParagraph(p)))
    .filter(Boolean)
    .join("\n\n");
}

/**
 * @param {HTMLButtonElement | null} translateBtn
 */
async function runTranslateSelection(translateBtn) {
  const status = $("#status");
  const preview = $("#preview");
  const targetLang = $("#target-language").value;

  setAllButtons(true);
  beginRun();
  const signal = getActiveRunSignal();
  translateProgressStart(translateBtn);
  progressStart("Reading selection in Word");
  progressSet(0, "Capturing selection (bookmark + OOXML)");
  preview.value = "";
  setStatus(status, "Reading selection…");
  appLog("translate", "Translate selection started", { targetLang });
  debugTrace("translate", `runTranslateSelection: target=${targetLang}`);

  /** @type {string | null} */
  let selectionBookmark = null;
  try {
    const cap = await captureSelectionOoxmlWithBookmark();
    selectionBookmark = cap.bookmarkName;
    const { ooxml, text } = cap;
    appLog("translate", "Selection captured", { ooxmlChars: ooxml.length, plainChars: text.length });
    debugTrace("translate", `selection captured: OOXML ~${ooxml.length} chars, plain ~${text.length} chars`);

    const totalLines = countTranslatableLinesInOoxmlTree(ooxml);
    let doneLines = 0;

    progressSet(8, `Translating selection · ${totalLines || "?"} line(s)`);
    translateProgressSet(translateBtn, 8);
    setStatus(
      status,
      totalLines > 0 ? `Translating selection · ${totalLines} line(s)…` : "Translating selection…",
    );

    const outOoxml = await translateEveryParagraphInOoxmlTree(ooxml, {
      signal,
      onReadyToTranslate: (nLines, nBatches) => {
        appLog("translate", "Selection — AI phase ready", { lines: nLines, plannedApiCalls: nBatches });
        if (nBatches > 0) {
          progressMarkAiBatches(nBatches);
        }
      },
      onBatchComplete: () => {
        progressMarkAiBatchFinished();
      },
      onMergeProgress: (done, total) => {
        if (done === 1) {
          progressTimingClearEta();
        }
        const pct = 8 + (done / total) * 72;
        progressSet(pct, `Merging selection OOXML ${done}/${total}`);
        translateProgressSet(translateBtn, pct);
      },
      translateBatch: (batch) => translateLinesBatchWithFallback(batch, targetLang, signal),
      onTranslatedLine: () => {
        doneLines++;
        const pct = totalLines > 0 ? (doneLines / totalLines) * 80 : 40;
        translateProgressSet(translateBtn, pct);
        progressSet(
          pct,
          totalLines > 0 ? `Selection · line ${doneLines}/${totalLines}` : "Translating selection",
        );
        setStatus(
          status,
          totalLines > 0
            ? `Translating selection · line ${doneLines}/${totalLines}…`
            : "Translating selection…",
        );
      },
    });

    if (outOoxml == null) {
      appLog("translate", "Selection → plain-text path (no w:p in OOXML)");
      debugTrace("translate", "selection has no w:p — plain-text translate path");
      setStatus(status, "Translating selection (plain text)…");
      translateProgressSet(translateBtn, 25);
      progressSet(25, "Plain-text translation (no w:p in OOXML)");
      const msgs = prompts.translateSegmentPrompt(text, targetLang);
      const raw = await completeChat(msgs, { maxTokens: 4096, signal });
      const translated = String(raw ?? "")
        .replace(/\r\n/g, "\n")
        .replace(/^["'`]+|["'`]+$/g, "")
        .trimEnd();
      progressMarkWritePhase(1);
      await replaceBookmarkRangeTextTracked(selectionBookmark, translated);
      progressMarkWriteOneDone();
      preview.value = translated.length > 24000 ? `${translated.slice(0, 24000)}\n…` : translated;
      progressSet(100, "Writing plain text to Word");
      translateProgressSet(translateBtn, 100);
      progressDone(false);
      setStatus(
        status,
        `Translated selection to ${targetLang} as plain text (this selection had no paragraph markup in OOXML).`,
      );
      debugTrace("translate", "runTranslateSelection: finished OK (plain-text path)");
      return;
    }

    translateProgressSet(translateBtn, 88);
    progressSet(88, "Writing OOXML selection to Word");
    progressMarkWritePhase(1);
    setStatus(status, "Writing selection to Word…");
    appLog("word", "insertOoxml selection (bookmark range)");
    await replaceBookmarkRangeOoxmlTracked(selectionBookmark, outOoxml);
    progressMarkWriteOneDone();
    setLivePreview(previewPlainFromOoxmlRoot(outOoxml), preview);
    progressSet(100, "Selection translation complete");
    translateProgressSet(translateBtn, 100);
    progressDone(false);
    setStatus(
      status,
      `Translated selection to ${targetLang}. Bold, italic, and underline are kept per segment where applicable. Review tracked changes in Word.`,
    );
    appLog("translate", "Translate selection finished OK (OOXML)");
    debugTrace("translate", "runTranslateSelection: finished OK (OOXML path)");
  } catch (e) {
    progressDone(true);
    appWarn("translate", "Translate selection failed", e instanceof Error ? e.message : String(e));
    debugTraceError("translate", e);
    const msg = formatAddinError(e);
    setStatus(status, msg, true);
    preview.value = msg;
  } finally {
    endRun();
    await tryDeleteBookmark(selectionBookmark);
    translateProgressDone(translateBtn);
    setAllButtons(false);
  }
}

/** @param {string} text @param {HTMLTextAreaElement} el */
function setLivePreview(text, el) {
  const maxLen = 24000;
  el.value = text.length > maxLen ? `${text.slice(0, maxLen)}\n…` : text;
  el.scrollTop = el.scrollHeight;
}

function $(sel, root = document) {
  return root.querySelector(sel);
}

function setStatus(el, text, isError = false) {
  el.textContent = text || "";
  el.classList.toggle("is-error", Boolean(isError));
}

function showView(name) {
  document.querySelectorAll(".view").forEach((v) => {
    const id = v.id.replace("view-", "");
    const active = id === name;
    v.hidden = !active;
    v.classList.toggle("is-visible", active);
  });
  document.querySelectorAll(".tab").forEach((t) => {
    const active = t.dataset.view === name;
    t.classList.toggle("is-active", active);
    t.setAttribute("aria-selected", active ? "true" : "false");
  });
}

function formToSettings() {
  const s = loadSettings();
  s.requestMode = $("#request-mode").value === "direct" ? "direct" : "proxy";
  s.proxyUrl = $("#proxy-url").value.trim();
  s.piiEngine = $("#pii-engine").value === "presidio" ? "presidio" : "llm";
  s.presidioBackendUrl = ($("#presidio-backend-url")?.value || "").trim();
  const intervalRaw = $("#llm-min-interval-ms").value.trim();
  if (intervalRaw === "") {
    delete s.llmMinIntervalMs;
  } else {
    const n = parseInt(intervalRaw, 10);
    if (Number.isFinite(n)) s.llmMinIntervalMs = Math.min(120_000, Math.max(0, n));
    else delete s.llmMinIntervalMs;
  }
  s.activeProvider = /** @type {typeof s.activeProvider} */ ($("#active-provider").value);

  s.providers.openai.enabled = $("#openai-enabled").checked;
  s.providers.openai.apiKey = $("#openai-key").value;
  s.providers.openai.model = $("#openai-model").value.trim() || s.providers.openai.model;
  s.providers.openai.baseUrl = $("#openai-base").value.trim();
  s.providers.openai.organizationId = $("#openai-org").value.trim();
  s.providers.openai.projectId = $("#openai-project").value.trim();

  if (s.activeProvider === "openai" && s.providers.openai.enabled === false) {
    s.activeProvider = "anthropic";
  }

  s.providers.anthropic.apiKey = $("#anthropic-key").value;
  s.providers.anthropic.model = $("#anthropic-model").value.trim() || s.providers.anthropic.model;
  s.providers.anthropic.baseUrl = $("#anthropic-base").value.trim();

  s.providers.azureOpenai.apiKey = $("#azure-key").value;
  s.providers.azureOpenai.baseUrl = $("#azure-endpoint").value.trim();
  s.providers.azureOpenai.deployment = $("#azure-deployment").value.trim();
  s.providers.azureOpenai.apiVersion = $("#azure-version").value.trim() || s.providers.azureOpenai.apiVersion;

  s.providers.customOpenAI.apiKey = $("#custom-key").value;
  s.providers.customOpenAI.baseUrl = $("#custom-base").value.trim() || s.providers.customOpenAI.baseUrl;
  s.providers.customOpenAI.model = $("#custom-model").value.trim() || s.providers.customOpenAI.model;

  return s;
}

function syncOpenAiAllowUi() {
  const enabled = $("#openai-enabled").checked;
  const creds = document.querySelector(".openai-credentials");
  if (creds) creds.classList.toggle("is-disabled", !enabled);
  const opt = $("#active-provider")?.querySelector('option[value="openai"]');
  if (opt) opt.disabled = !enabled;
  if (!enabled && $("#active-provider").value === "openai") {
    $("#active-provider").value = "anthropic";
  }
}

function settingsToForm() {
  const s = loadSettings();
  $("#request-mode").value = s.requestMode === "direct" ? "direct" : "proxy";
  $("#proxy-url").value = s.proxyUrl || "";
  $("#pii-engine").value = s.piiEngine === "presidio" ? "presidio" : "llm";
  const presidioUrlEl = $("#presidio-backend-url");
  if (presidioUrlEl) presidioUrlEl.value = s.presidioBackendUrl || "";
  $("#llm-min-interval-ms").value =
    s.llmMinIntervalMs !== undefined && s.llmMinIntervalMs !== null ? String(s.llmMinIntervalMs) : "";
  let ap = s.activeProvider;
  if (ap === "openai" && s.providers.openai.enabled === false) ap = "anthropic";
  $("#active-provider").value = ap;

  $("#openai-enabled").checked = s.providers.openai.enabled !== false;
  $("#openai-key").value = s.providers.openai.apiKey || "";
  $("#openai-model").value = s.providers.openai.model || "";
  $("#openai-base").value = s.providers.openai.baseUrl || "";
  $("#openai-org").value = s.providers.openai.organizationId || "";
  $("#openai-project").value = s.providers.openai.projectId || "";
  syncOpenAiAllowUi();

  $("#anthropic-key").value = s.providers.anthropic.apiKey || "";
  $("#anthropic-model").value = s.providers.anthropic.model || "";
  $("#anthropic-base").value = s.providers.anthropic.baseUrl || "";

  $("#azure-key").value = s.providers.azureOpenai.apiKey || "";
  $("#azure-endpoint").value = s.providers.azureOpenai.baseUrl || "";
  $("#azure-deployment").value = s.providers.azureOpenai.deployment || "";
  $("#azure-version").value = s.providers.azureOpenai.apiVersion || "";

  $("#custom-key").value = s.providers.customOpenAI.apiKey || "";
  $("#custom-base").value = s.providers.customOpenAI.baseUrl || "";
  $("#custom-model").value = s.providers.customOpenAI.model || "";

  const debugCb = $("#debug-trace-enabled");
  if (debugCb) debugCb.checked = isDebugTraceEnabled();
}

// ---------------------------------------------------------------------------
// Progress timing (elapsed + estimated remaining)
// ---------------------------------------------------------------------------

/** @type {{ startedAt: number, tickerId: ReturnType<typeof setInterval> | 0, etaType: 'none' | 'ai_batches' | 'write' | 'redact', etaTotal: number, etaDone: number, etaPhaseStartAt: number }} */
let progressClock = {
  startedAt: 0,
  tickerId: 0,
  etaType: "none",
  etaTotal: 0,
  etaDone: 0,
  etaPhaseStartAt: 0,
};

/** Heuristic before the first translation batch returns (network + throttle + JSON). */
const ROUGH_MS_PER_TRANSLATE_BATCH = 28_000;

/** @param {number} ms */
function formatDurationShort(ms) {
  if (ms < 0 || !Number.isFinite(ms)) return "—";
  const s = Math.floor(ms / 1000);
  if (s < 60) return `${s}s`;
  const m = Math.floor(s / 60);
  const rs = s % 60;
  if (m < 60) return rs > 0 ? `${m}m ${rs}s` : `${m}m`;
  const h = Math.floor(m / 60);
  const rm = m % 60;
  return `${h}h ${rm}m`;
}

function progressTimingClearEta() {
  progressClock.etaType = "none";
  progressClock.etaTotal = 0;
  progressClock.etaDone = 0;
  progressClock.etaPhaseStartAt = 0;
}

/** @param {number} totalBatches */
function progressMarkAiBatches(totalBatches) {
  if (totalBatches <= 0) return;
  progressClock.etaType = "ai_batches";
  progressClock.etaTotal = totalBatches;
  progressClock.etaDone = 0;
  progressClock.etaPhaseStartAt = Date.now();
  progressTickTiming();
}

function progressMarkAiBatchFinished() {
  if (progressClock.etaType === "ai_batches") {
    progressClock.etaDone += 1;
  }
  progressTickTiming();
}

/** @param {number} totalParas */
function progressMarkWritePhase(totalParas) {
  if (totalParas <= 0) return;
  progressClock.etaType = "write";
  progressClock.etaTotal = totalParas;
  progressClock.etaDone = 0;
  progressClock.etaPhaseStartAt = Date.now();
  progressTickTiming();
}

function progressMarkWriteOneDone() {
  if (progressClock.etaType === "write") {
    progressClock.etaDone += 1;
  }
  progressTickTiming();
}

/** @param {number} totalItems */
function progressMarkRedactPhase(totalItems) {
  if (totalItems <= 0) return;
  progressClock.etaType = "redact";
  progressClock.etaTotal = totalItems;
  progressClock.etaDone = 0;
  progressClock.etaPhaseStartAt = Date.now();
  progressTickTiming();
}

/** @param {number} done */
function progressMarkRedactProgress(done) {
  if (progressClock.etaType === "redact") {
    progressClock.etaDone = done;
  }
  progressTickTiming();
}

function estimateRemainingMs() {
  const now = Date.now();
  const { etaType, etaTotal, etaDone, etaPhaseStartAt } = progressClock;

  if (etaType === "ai_batches" && etaTotal > 0) {
    if (etaDone >= etaTotal) return null;
    if (etaDone > 0) {
      const phaseElapsed = now - etaPhaseStartAt;
      const avg = phaseElapsed / etaDone;
      return Math.max(0, Math.round(avg * (etaTotal - etaDone)));
    }
    return etaTotal * ROUGH_MS_PER_TRANSLATE_BATCH;
  }
  if (etaType === "write" && etaTotal > 0) {
    if (etaDone >= etaTotal) return null;
    if (etaDone > 0) {
      const phaseElapsed = now - etaPhaseStartAt;
      const avg = phaseElapsed / etaDone;
      return Math.max(0, Math.round(avg * (etaTotal - etaDone)));
    }
    return null;
  }
  if (etaType === "redact" && etaTotal > 0) {
    if (etaDone >= etaTotal) return null;
    if (etaDone > 0) {
      const phaseElapsed = now - etaPhaseStartAt;
      const avg = phaseElapsed / etaDone;
      return Math.max(0, Math.round(avg * (etaTotal - etaDone)));
    }
    return null;
  }
  return null;
}

function progressTickTiming() {
  const timingEl = $("#progress-timing");
  const bar = $("#progress");
  if (!timingEl || !bar || bar.classList.contains("is-hidden")) return;

  const elapsed = progressClock.startedAt ? Date.now() - progressClock.startedAt : 0;
  const rem = estimateRemainingMs();
  const elapsedStr = formatDurationShort(elapsed);

  let line = `Elapsed ${elapsedStr}`;
  if (rem != null) {
    const rough =
      progressClock.etaType === "ai_batches" && progressClock.etaDone === 0 && progressClock.etaTotal > 0;
    const remStr = formatDurationShort(rem);
    line += rough
      ? ` · Est. remaining ~${remStr} (rough) · Est. total job ~${formatDurationShort(elapsed + rem)} (rough)`
      : ` · Est. remaining ~${remStr} · Est. total ~${formatDurationShort(elapsed + rem)}`;
  } else {
    line += " · Est. remaining / total —";
  }

  timingEl.textContent = line;
}

function progressStartClock() {
  progressClock.startedAt = Date.now();
  progressTimingClearEta();
  if (progressClock.tickerId) {
    clearInterval(progressClock.tickerId);
    progressClock.tickerId = 0;
  }
  progressClock.tickerId = setInterval(() => progressTickTiming(), 1000);
  progressTickTiming();
}

function progressStopClock() {
  if (progressClock.tickerId) {
    clearInterval(progressClock.tickerId);
    progressClock.tickerId = 0;
  }
  progressClock.startedAt = 0;
  progressTimingClearEta();
  const timingEl = $("#progress-timing");
  if (timingEl) timingEl.textContent = "";
}

// ---------------------------------------------------------------------------
// Actions
// ---------------------------------------------------------------------------

function setAllButtons(disabled) {
  document.querySelectorAll("[data-action]").forEach((b) => {
    b.disabled = disabled;
  });
  const chatSend = $("#btn-chat-send");
  const chatApply = $("#btn-chat-apply");
  if (chatSend) chatSend.disabled = disabled;
  if (chatApply) chatApply.disabled = disabled;
  const stopBtn = $("#btn-stop");
  if (stopBtn) stopBtn.disabled = !disabled;
}

/** @param {string} [phase] Short description of current step (shown under the bar). */
function progressSetLabel(phase) {
  const lab = $("#progress-label");
  if (lab) lab.textContent = phase || "";
}

/**
 * @param {string} [phase] Initial phase label while work is indeterminate.
 */
function progressStart(phase) {
  document.body.classList.add("progress-active");
  const bar = $("#progress");
  const fill = $("#progress-fill");
  bar.classList.remove("is-hidden");
  fill.className = "progress-bar__fill is-indeterminate";
  fill.style.width = "";
  progressStartClock();
  progressSetLabel(phase ?? "Starting…");
  bar?.setAttribute("aria-valuenow", "0");
  bar?.setAttribute("aria-valuetext", phase ?? "Starting…");
  progressTickTiming();
}

/**
 * @param {number} pct 0–100
 * @param {string} [phase] Current step; combined with percentage in the label.
 */
function progressSet(pct, phase) {
  const bar = $("#progress");
  const fill = $("#progress-fill");
  fill.className = "progress-bar__fill";
  const p = Math.min(100, Math.max(0, pct));
  fill.style.width = `${p}%`;
  const label =
    phase != null && String(phase).trim() !== ""
      ? `${phase} · ${Math.round(p)}%`
      : `${Math.round(p)}%`;
  progressSetLabel(label);
  bar?.setAttribute("aria-valuenow", String(Math.round(p)));
  bar?.setAttribute("aria-valuetext", label);
  progressTickTiming();
}

/** Update the label only, keeping the current fill width (for sub-steps like “waiting on AI”). */
function progressPhase(detail) {
  const bar = $("#progress");
  const fill = $("#progress-fill");
  const w = fill?.style?.width || "";
  const m = /^([\d.]+)%$/.exec(w);
  const p = m ? Math.min(100, Math.max(0, Number(m[1]))) : 0;
  const label =
    detail != null && String(detail).trim() !== ""
      ? `${detail} · ${Math.round(p)}%`
      : `${Math.round(p)}%`;
  progressSetLabel(label);
  bar?.setAttribute("aria-valuetext", label);
  progressTickTiming();
}

function progressDone(isError = false) {
  if (progressClock.tickerId) {
    clearInterval(progressClock.tickerId);
    progressClock.tickerId = 0;
  }
  const elapsed = progressClock.startedAt ? Date.now() - progressClock.startedAt : 0;
  progressClock.startedAt = 0;
  progressTimingClearEta();

  const timingEl = $("#progress-timing");
  if (timingEl) {
    if (elapsed > 500) {
      timingEl.textContent = isError
        ? `Elapsed ${formatDurationShort(elapsed)} · stopped`
        : `Total ${formatDurationShort(elapsed)}`;
    } else {
      timingEl.textContent = "";
    }
  }

  const bar = $("#progress");
  const fill = $("#progress-fill");
  if (!isError) {
    progressSetLabel("Done");
    fill.style.width = "100%";
    fill.className = "progress-bar__fill is-done";
  } else {
    progressSetLabel("Error or stopped");
    fill.style.width = "100%";
    fill.className = "progress-bar__fill is-error";
  }
  setTimeout(() => {
    bar?.classList.add("is-hidden");
    fill.className = "progress-bar__fill";
    fill.style.width = "";
    document.body.classList.remove("progress-active");
    progressSetLabel("");
    progressStopClock();
    bar?.setAttribute("aria-valuenow", "0");
    bar?.setAttribute("aria-valuetext", "");
  }, 1200);
}

/** @param {string} result */
function parsePiiArrayFromModel(result) {
  const cleaned = result.replace(/^```json?\s*/i, "").replace(/```\s*$/, "").trim();
  const items = JSON.parse(cleaned);
  if (!Array.isArray(items)) throw new Error("AI did not return a JSON array.");
  return items;
}

/**
 * @param {string} text
 * @param {AbortSignal | undefined} signal
 */
async function scanDocumentTextForPii(text, signal) {
  const s = loadSettings();
  if (s.piiEngine === "presidio") {
    return detectPiiItems(text, { signal });
  }
  const result = await completeChat(prompts.anonymizePrompt(text), { maxTokens: 8192, signal });
  return parsePiiArrayFromModel(result);
}

/** Index after last `}` of the object starting at `startIdx`, or -1. Braces inside JSON strings are ignored. */
function findJsonObjectEnd(text, startIdx) {
  let depth = 0;
  let inString = false;
  let escape = false;
  for (let i = startIdx; i < text.length; i++) {
    const c = text[i];
    if (escape) {
      escape = false;
      continue;
    }
    if (inString) {
      if (c === "\\") escape = true;
      else if (c === '"') inString = false;
      continue;
    }
    if (c === '"') {
      inString = true;
      continue;
    }
    if (c === "{") depth++;
    else if (c === "}") {
      depth--;
      if (depth === 0) return i + 1;
    }
  }
  return -1;
}

/**
 * @param {string} action
 * @param {HTMLButtonElement | null} [sourceButton]
 */
async function runAction(action, sourceButton) {
  if (action === "translate") {
    await runTranslate(sourceButton ?? null);
    return;
  }
  if (action === "translate-selection") {
    await runTranslateSelection(sourceButton ?? null);
    return;
  }

  const status = $("#status");
  const preview = $("#preview");
  const anonymizeMatch = /^anonymize-(selection|file)-(redact|fake)$/.exec(action);
  const isFile = anonymizeMatch ? anonymizeMatch[1] === "file" : false;
  const scope = isFile ? "file" : "selection";
  /** @type {'redact' | 'fake' | null} */
  const replacementMode = anonymizeMatch ? (anonymizeMatch[2] === "redact" ? "redact" : "fake") : null;

  preview.value = "";
  setAllButtons(true);
  progressStart(isFile ? "Anonymize · reading full document" : "Anonymize · reading selection");
  setStatus(status, `Reading ${scope}…`);

  let text;
  try {
    text = isFile ? await getFullDocumentText() : await getSelectedText();
  } catch (e) {
    progressDone(true);
    setStatus(status, formatAddinError(e), true);
    setAllButtons(false);
    return;
  }

  if (!text?.trim()) {
    progressDone(true);
    setStatus(status, isFile ? "The document appears empty." : "Select text in the document first.", true);
    setAllButtons(false);
    return;
  }

  const isAnonymize = Boolean(anonymizeMatch);
  const usePresidio = loadSettings().piiEngine === "presidio";

  beginRun();
  const signal = getActiveRunSignal();
  progressPhase(usePresidio ? "Presidio · detecting PII" : "AI · detecting PII…");
  setStatus(status, usePresidio ? "Calling Presidio…" : "Calling AI…");

  try {
    if (isAnonymize) {
      let items;
      try {
        items = await scanDocumentTextForPii(text, signal);
      } catch (e) {
        const raw = e instanceof Error ? e.message : String(e);
        throw new Error(usePresidio ? `Presidio PII scan failed: ${raw}` : `AI did not return valid PII JSON: ${raw}`);
      }
      if (items.length === 0) {
        progressDone(false);
        setStatus(status, "No PII found in the text.");
        preview.value = "No personally identifiable information detected.";
        return;
      }

      preview.value = items.map((i) => `${i.category}: ${i.original}`).join("\n");
      setStatus(status, `Redacting ${items.length} item(s)…`);
      progressMarkRedactPhase(items.length);
      progressSet(0, `Word · replacing ${items.length} PII pattern(s)`);

      const redactScope = isFile ? "body" : "selection";
      const mode = /** @type {'redact' | 'fake'} */ (replacementMode || "fake");
      const count = await redactItems(
        items,
        redactScope,
        (done, total) => {
          progressMarkRedactProgress(done);
          progressSet((done / total) * 100, `Redacting ${done}/${total} in Word`);
          setStatus(status, `Redacting ${done}/${total}…`);
        },
        mode,
        signal,
      );

      progressDone(false);
      setStatus(
        status,
        mode === "redact"
          ? `Replaced ${count} occurrence(s) with [REDACTED]. Review tracked changes in Word.`
          : `Replaced ${count} occurrence(s) with random fictitious names/data. Review tracked changes in Word.`,
      );
    }
  } catch (e) {
    const msg = formatAddinError(e);
    progressDone(true);
    setStatus(status, msg, true);
    preview.value = msg;
  } finally {
    endRun();
    setAllButtons(false);
  }
}

// ---------------------------------------------------------------------------
// Chat
// ---------------------------------------------------------------------------

/** @type {{ role: string, content: string }[]} */
let chatHistory = [];
let lastApplicableContent = "";
let pendingAction = null;

function addChatBubble(type, text) {
  const container = $("#chat-messages");
  const el = document.createElement("div");
  el.className = `chat-bubble chat-bubble--${type}`;
  el.textContent = text;
  container.appendChild(el);
  container.scrollTop = container.scrollHeight;
  return el;
}

function scrollChat() {
  const container = $("#chat-messages");
  container.scrollTop = container.scrollHeight;
}

function showConfirmation(description) {
  return new Promise((resolve) => {
    const container = $("#chat-messages");
    const bubble = document.createElement("div");
    bubble.className = "chat-bubble chat-bubble--confirm";

    const msg = document.createElement("div");
    msg.textContent = description;
    bubble.appendChild(msg);

    const actions = document.createElement("div");
    actions.className = "confirm-actions";

    const yesBtn = document.createElement("button");
    yesBtn.className = "btn btn-confirm-yes";
    yesBtn.textContent = "Yes, apply";
    yesBtn.type = "button";

    const noBtn = document.createElement("button");
    noBtn.className = "btn btn-confirm-no";
    noBtn.textContent = "No, cancel";
    noBtn.type = "button";

    actions.appendChild(yesBtn);
    actions.appendChild(noBtn);
    bubble.appendChild(actions);
    container.appendChild(bubble);
    scrollChat();

    function cleanup(result) {
      yesBtn.disabled = true;
      noBtn.disabled = true;
      yesBtn.remove();
      noBtn.remove();
      const status = document.createElement("div");
      status.style.marginTop = "4px";
      status.style.fontSize = "11px";
      status.style.fontStyle = "italic";
      status.textContent = result ? "Confirmed." : "Cancelled.";
      bubble.appendChild(status);
      resolve(result);
    }

    yesBtn.addEventListener("click", () => cleanup(true));
    noBtn.addEventListener("click", () => cleanup(false));
  });
}

function normalizeAnonymizeAction(action) {
  if (action === "anonymize_selection") return "anonymize_selection_fake";
  if (action === "anonymize_document") return "anonymize_document_fake";
  return action;
}

function parseAiResponse(raw) {
  let text = raw.trim();
  text = text.replace(/^```json?\s*/i, "").replace(/```\s*$/, "").trim();
  try {
    const parsed = JSON.parse(text);
    if (parsed && typeof parsed.action === "string") {
      if (typeof parsed.content !== "string") {
        parsed.content = JSON.stringify(parsed.content);
      }
      parsed.action = normalizeAnonymizeAction(parsed.action);
      return parsed;
    }
  } catch { /* fall through */ }

  const match = text.match(/\{[\s\S]*"action"\s*:\s*"[^"]+"/);
  if (match) {
    try {
      const idx = text.indexOf(match[0]);
      const end = findJsonObjectEnd(text, idx);
      if (end < 0) throw new Error("unclosed");
      const parsed = JSON.parse(text.slice(idx, end));
      if (parsed && typeof parsed.action === "string") {
        if (typeof parsed.content !== "string") {
          parsed.content = JSON.stringify(parsed.content);
        }
        parsed.action = normalizeAnonymizeAction(parsed.action);
        return parsed;
      }
    } catch { /* fall through */ }
  }

  return { action: "display", content: raw };
}

function describeAction(action, content) {
  if (action === "anonymize_selection_redact")
    return "Redact PII in the selection — each hit replaced with [REDACTED] (tracked changes).";
  if (action === "anonymize_selection_fake")
    return "Anonymize the selection — each PII item replaced with random fictitious names, companies, addresses, etc. (tracked changes).";
  if (action === "anonymize_document_redact")
    return "Redact PII in the entire document — each hit replaced with [REDACTED] (tracked changes).";
  if (action === "anonymize_document_fake")
    return "Anonymize the entire document — each PII item replaced with random fictitious stand-ins (tracked changes).";
  if (action === "apply_selection") {
    const preview = content.length > 80 ? content.slice(0, 80) + "…" : content;
    return `Apply changes to the selection as tracked revisions:\n"${preview}"`;
  }
  if (action === "apply_document") {
    const preview = content.length > 80 ? content.slice(0, 80) + "…" : content;
    return `Apply changes to the entire document as tracked revisions:\n"${preview}"`;
  }
  return "Apply changes to the document.";
}

/** @param {string} action anonymize_selection_redact | anonymize_selection_fake | anonymize_document_redact | anonymize_document_fake */
async function runChatAnonymizeScan(action, selectedText, fullDocument) {
  const isSelection = action.includes("selection");
  const replacementMode = action.endsWith("_redact") ? "redact" : "fake";
  const redactScope = isSelection ? "selection" : "body";
  const text = isSelection ? selectedText : fullDocument;
  if (!text?.trim()) {
    addChatBubble(
      "error",
      isSelection
        ? "Select text in the document first, or ask to anonymize the whole document."
        : "The document appears empty.",
    );
    return;
  }

  const usePresidio = loadSettings().piiEngine === "presidio";
  addChatBubble("system", usePresidio ? "Scanning for PII via Presidio…" : "Scanning for PII (same process as the Tools anonymize buttons)…");

  const signal = getActiveRunSignal();
  progressPhase(usePresidio ? "Presidio · detecting PII" : "AI · detecting PII in text");
  let items;
  try {
    items = await scanDocumentTextForPii(text, signal);
  } catch (e) {
    const raw = e instanceof Error ? e.message : String(e);
    throw new Error(usePresidio ? `Presidio PII scan failed: ${raw}` : `PII scan failed: ${raw}`);
  }
  if (items.length === 0) {
    addChatBubble("ai", "No personally identifiable information detected.");
    return;
  }

  addChatBubble("ai", `Found ${items.length} PII item(s): ${items.map((i) => `${i.category}: ${i.original}`).join(", ")}`);

  progressMarkRedactPhase(items.length);
  progressSet(0, `Chat · redacting ${items.length} pattern(s) in Word`);
  const count = await redactItems(
    items,
    redactScope,
    (done, total) => {
      progressMarkRedactProgress(done);
      progressSet((done / total) * 100, `Chat redact ${done}/${total}`);
    },
    replacementMode,
    signal,
  );

  addChatBubble(
    "system",
    replacementMode === "redact"
      ? `Redacted ${count} occurrence(s) with [REDACTED] (tracked changes). Review in Word.`
      : `Redacted ${count} occurrence(s) with random fictitious stand-ins (tracked changes). Review in Word.`,
  );
}

async function executeAction(action, content) {
  if (action === "apply_selection") {
    progressMarkWritePhase(1);
    progressSet(45, "Chat · applying to selection (tracked)");
    await replaceSelectionTracked(content);
    progressMarkWriteOneDone();
    addChatBubble("system", "Applied to selection as tracked changes. Review in Word.");
    lastApplicableContent = content;

  } else if (action === "apply_document") {
    progressMarkWritePhase(1);
    progressSet(45, "Chat · applying to full document (tracked)");
    await replaceBodyTracked(content);
    progressMarkWriteOneDone();
    addChatBubble("system", "Applied to full document as tracked changes. Review in Word.");
    lastApplicableContent = content;
  }
}

async function sendChat() {
  const input = $("#chat-input");
  const userMsg = input.value.trim();
  if (!userMsg) return;

  input.value = "";
  addChatBubble("user", userMsg);
  progressStart("Chat · loading document and selection");
  setAllButtons(true);
  beginRun();
  const signal = getActiveRunSignal();

  try {
    let selectedText = "";
    let fullDocument = "";
    try {
      [selectedText, fullDocument] = await Promise.all([
        getSelectedText().catch(() => ""),
        getFullDocumentText(),
      ]);
    } catch (e) {
      progressDone(true);
      addChatBubble("error", formatAddinError(e));
      return;
    }

    if (!fullDocument?.trim()) {
      progressDone(true);
      addChatBubble("error", "The document appears empty.");
      return;
    }

    const messages = prompts.chatMessages(selectedText, fullDocument, chatHistory, userMsg);

    let rawReply;
    try {
      progressSet(12, "Chat · waiting for AI response");
      rawReply = await completeChat(messages, { maxTokens: 4096, signal });
    } catch (e) {
      addChatBubble("error", formatAddinError(e));
      progressDone(true);
      return;
    }

    chatHistory.push({ role: "user", content: userMsg });
    chatHistory.push({ role: "assistant", content: rawReply });

    const { action, content } = parseAiResponse(rawReply);
    const isModifying = [
      "apply_selection",
      "apply_document",
      "anonymize_selection_redact",
      "anonymize_selection_fake",
      "anonymize_document_redact",
      "anonymize_document_fake",
    ].includes(action);

    if (isModifying) {
      progressDone(false);
      const confirmed = await showConfirmation(describeAction(action, content));
      if (!confirmed) {
        addChatBubble("system", "Action cancelled.");
        return;
      }
      progressStart("Chat · running confirmed action");
      progressSet(20, "Working in Word…");
      try {
        if (
          action === "anonymize_selection_redact" ||
          action === "anonymize_selection_fake" ||
          action === "anonymize_document_redact" ||
          action === "anonymize_document_fake"
        ) {
          progressPhase("PII scan + redaction");
          await runChatAnonymizeScan(action, selectedText, fullDocument);
        } else {
          await executeAction(action, content);
        }
        progressSet(100, "Complete");
        progressDone(false);
      } catch (e) {
        addChatBubble("error", formatAddinError(e));
        progressDone(true);
      }
    } else {
      addChatBubble("ai", content);
      progressSet(100, "Reply received");
      progressDone(false);
    }
  } finally {
    endRun();
    setAllButtons(false);
  }
}

async function applyChatResponse() {
  if (!lastApplicableContent) {
    addChatBubble("error", "No applicable AI response yet. The last response was informational.");
    return;
  }
  const confirmed = await showConfirmation("Re-apply last change to current selection as tracked revisions?");
  if (!confirmed) return;
  progressStart("Apply · tracked replace in Word");
  setAllButtons(true);
  beginRun();
  try {
    progressSet(55, "Inserting tracked revision");
    progressMarkWritePhase(1);
    await replaceSelectionTracked(lastApplicableContent);
    progressMarkWriteOneDone();
    addChatBubble("system", "Re-applied last change to current selection as tracked changes.");
    progressSet(100, "Re-applied");
    progressDone(false);
  } catch (e) {
    addChatBubble("error", formatAddinError(e));
    progressDone(true);
  } finally {
    endRun();
    setAllButtons(false);
  }
}

function clearChat() {
  chatHistory = [];
  lastApplicableContent = "";
  pendingAction = null;
  $("#chat-messages").innerHTML = "";
}

// ---------------------------------------------------------------------------
// Wire UI
// ---------------------------------------------------------------------------

function wireUi() {
  document.querySelectorAll(".tab").forEach((tab) => {
    tab.addEventListener("click", () => {
      const v = tab.dataset.view;
      if (v === "settings") settingsToForm();
      showView(v);
    });
  });

  $("#openai-enabled").addEventListener("change", syncOpenAiAllowUi);

  $("#btn-save").addEventListener("click", () => {
    const s = formToSettings();
    saveSettings(s);
    syncOpenAiAllowUi();
    setStatus($("#settings-status"), "Settings saved.", false);
  });

  $("#btn-reset").addEventListener("click", () => {
    saveSettings(defaultSettings());
    settingsToForm();
    setStatus($("#settings-status"), "Reset to defaults.", false);
  });

  const debugTraceCb = $("#debug-trace-enabled");
  if (debugTraceCb) {
    debugTraceCb.addEventListener("change", () => {
      setDebugTraceEnabled(debugTraceCb.checked);
      setStatus(
        $("#settings-status"),
        debugTraceCb.checked
          ? "Debug trace on — open the browser devtools console (F12) to see detailed logs."
          : "Debug trace off.",
        false,
      );
    });
  }

  document.querySelectorAll("[data-action]").forEach((btn) => {
    btn.addEventListener("click", () => runAction(btn.getAttribute("data-action"), btn));
  });

  $("#btn-chat-send").addEventListener("click", sendChat);
  $("#btn-chat-apply").addEventListener("click", applyChatResponse);
  $("#btn-chat-clear").addEventListener("click", clearChat);

  const stopBtn = $("#btn-stop");
  if (stopBtn) {
    stopBtn.addEventListener("click", () => cancelActiveRun());
  }

  $("#chat-input").addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      sendChat();
    }
  });
}

function installGlobalErrorToStatus() {
  const statusEl = () => document.getElementById("status");
  window.addEventListener("error", (ev) => {
    const err = ev.error != null ? ev.error : ev.message;
    debugTraceError("runtime", err);
    const el = statusEl();
    if (!el) return;
    el.textContent = formatAddinError(err);
    el.classList.add("is-error");
  });
  window.addEventListener("unhandledrejection", (ev) => {
    debugTraceError("runtime", ev.reason);
    const el = statusEl();
    if (!el) return;
    el.textContent = formatAddinError(ev.reason);
    el.classList.add("is-error");
  });
}

function wireTokenUsageFooter() {
  const el = document.getElementById("token-usage-month");
  if (!el) return;
  const refresh = () => {
    const n = getTokensRecordedThisMonth();
    const label = currentMonthLabel();
    el.textContent = `Tokens this month (${label}): ${n.toLocaleString()}`;
    el.title =
      "Tokens this month: provider-reported usage when the API includes it; otherwise a rough estimate from your prompts and the model reply (same length heuristic as translation batching). Stored in this browser only; resets each calendar month (local time).";
  };
  refresh();
  window.addEventListener("legal-ai-tokens-updated", refresh);
}

Office.onReady(() => {
  appLog("app", "Office.onReady — task pane boot");
  installGlobalErrorToStatus();
  const s = loadSettings();
  if (s.activeProvider === "openai" && s.providers.openai.enabled === false) {
    s.activeProvider = "anthropic";
    saveSettings(s);
  }
  settingsToForm();
  wireUi();
  wireTokenUsageFooter();
  initDebugTraceFooter(document.getElementById("debug-trace-line"));
  showView("assistant");
});
