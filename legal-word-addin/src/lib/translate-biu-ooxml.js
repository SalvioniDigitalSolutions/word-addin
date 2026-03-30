/**
 * Paragraph translation with **bold / italic / underline** preserved per run group,
 * and each changed phrase emitted as `w:del` + `w:ins` for Word tracked revisions.
 */

import {
  contentTokenBudget,
  splitOversizedLineMaxParts,
  OVERSIZED_LINE_MAX_PARTS,
  stringEntryCost,
  maxPartTokensForSplit,
  countApiCallsForOrderedStrings,
  translateOrderedStringsInBudget,
} from "./translate-line-batching.js";
import { debugTrace } from "./debug-trace.js";

export const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const XML_SPACE_NS = "http://www.w3.org/XML/1998/namespace";

export const TRACK_TRANSLATION_AUTHOR = "Legal AI Assistant";

/**
 * Keep only XML 1.0 `Char` + common whitespace; normalize / drop marks Word’s OOXML loader often rejects.
 * @param {string} s
 */
function filterToXml10LegalText(s) {
  let out = "";
  for (const ch of s) {
    const cp = ch.codePointAt(0);
    if (cp === 0x2028 || cp === 0x2029) {
      out += "\n";
      continue;
    }
    if ((cp >= 0x202a && cp <= 0x202e) || (cp >= 0x2066 && cp <= 0x2069)) continue;
    if (
      cp === 0x9 ||
      cp === 0xa ||
      cp === 0xd ||
      (cp >= 0x20 && cp <= 0xd7ff) ||
      (cp >= 0xe000 && cp <= 0xfffd) ||
      (cp >= 0x10000 && cp <= 0x10ffff)
    ) {
      out += ch;
    }
  }
  return out;
}

/**
 * Strip / normalize characters that break strict XML 1.0 or Word’s OOXML temp package for `insertOoxml`.
 * - C0 controls (except tab / LF / CR, allowed in XML character data)
 * - BOM, U+FFFE / U+FFFF (non-characters in XML 1.0)
 * - Lone UTF-16 surrogates
 * - `]]>` (illegal in XML character data; browsers often serialize it raw → “can’t open … contents”, ~line 1 col 72)
 * - Bidi embedding marks (U+202A–U+202E, U+2066–U+2069) and line/paragraph separators (U+2028/U+2029)
 * - Any code point outside XML 1.0 `Char` (belt-and-suspenders vs. model junk)
 * @param {string} s
 */
export function sanitizeForWordXmlText(s) {
  let t = String(s ?? "")
    .replace(/\uFEFF/g, "")
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "");
  t = t.replace(/\uFFFE/g, "").replace(/\uFFFF/g, "");
  t = stripDefectiveUtf16Surrogates(t);
  t = t.replace(/]]>/g, "]]&gt;");
  t = filterToXml10LegalText(t);
  return t;
}

/**
 * Re-parse a serialized `w:p` and re-run {@link sanitizeForWordXmlText} on every `w:t` and `w:delText`
 * (new translation + text copied from cloned runs / `pPr`). Catches oddities after `XMLSerializer`.
 * @param {string} ooxmlParagraph
 */
function scrubParagraphTextElements(ooxmlParagraph) {
  const clean = stripXmlDecl(ooxmlParagraph);
  if (!clean) return ooxmlParagraph;
  const d = new DOMParser().parseFromString(clean, "application/xml");
  if (d.getElementsByTagName("parsererror").length) return ooxmlParagraph;
  const p = d.getElementsByTagNameNS(W_NS, "p")[0];
  if (!p) return ooxmlParagraph;
  for (const local of ["t", "delText"]) {
    const els = [...p.getElementsByTagNameNS(W_NS, local)];
    for (const el of els) {
      const raw = el.textContent ?? "";
      const fixed = sanitizeForWordXmlText(raw);
      if (fixed === raw) continue;
      el.textContent = fixed;
      if (/^\s|\s$|[\t\n]/.test(fixed)) {
        el.setAttributeNS(XML_SPACE_NS, "xml:space", "preserve");
      } else {
        el.removeAttribute("xml:space");
      }
    }
  }
  return serializeParagraph(p);
}

/** Remove lone surrogate code units (invalid XML / UTF-8). */
function stripDefectiveUtf16Surrogates(str) {
  if (!/[\uD800-\uDFFF]/.test(str)) return str;
  let out = "";
  for (let i = 0; i < str.length; i++) {
    const c = str.charCodeAt(i);
    if (c >= 0xd800 && c <= 0xdbff) {
      const lo = i + 1 < str.length ? str.charCodeAt(i + 1) : 0;
      if (lo >= 0xdc00 && lo <= 0xdfff) {
        out += str[i] + str[i + 1];
        i++;
      }
    } else if (c >= 0xdc00 && c <= 0xdfff) {
      /* lone low surrogate */
    } else {
      out += str[i];
    }
  }
  return out;
}

/** @param {string} s */
export function normText(s) {
  return String(s ?? "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
}

function stripXmlDecl(s) {
  return String(s || "")
    .replace(/^\uFEFF?/, "")
    .replace(/^<\?xml[^?]*\?>\s*/i, "")
    .trim();
}

/** @param {string} ooxmlParagraph */
export function parseParagraphOoxml(ooxmlParagraph) {
  const clean = stripXmlDecl(ooxmlParagraph);
  const doc = new DOMParser().parseFromString(clean, "application/xml");
  if (doc.getElementsByTagName("parsererror").length) {
    throw new Error("OOXML parse error");
  }
  const p = doc.getElementsByTagNameNS(W_NS, "p")[0];
  if (!p) throw new Error("No w:p in fragment");
  return p;
}

/** @param {Element} p */
export function serializeParagraph(p) {
  return new XMLSerializer().serializeToString(p);
}

/** @param {string} ooxmlParagraph */
export function paragraphPlainPreviewFromOoxml(ooxmlParagraph) {
  try {
    const p = parseParagraphOoxml(ooxmlParagraph);
    return p.textContent ?? "";
  } catch {
    return "";
  }
}

/**
 * Plain text for Word `insertText` fallback after failed `insertOoxml`: visible / “after” wording only.
 * Skips {@link paragraphPlainPreviewFromOoxml}, which concatenates `w:del` + `w:ins` and would duplicate source + translation.
 * @param {string} ooxmlParagraph
 */
export function paragraphTranslatedPlainFromMergedOoxml(ooxmlParagraph) {
  try {
    const p = parseParagraphOoxml(ooxmlParagraph);
    return collectParagraphPlainSkippingDeletes(p);
  } catch {
    return "";
  }
}

/** @param {Element} pEl w:p */
function collectParagraphPlainSkippingDeletes(pEl) {
  let s = "";
  for (const child of pEl.childNodes) {
    if (child.nodeType !== 1) continue;
    if (child.namespaceURI !== W_NS) continue;
    if (child.localName === "pPr") continue;
    if (child.localName === "del") continue;
    if (child.localName === "r") {
      s += getRunInlinePlain(child);
      continue;
    }
    s += elementPlainSkippingDelSubtree(child);
  }
  return s;
}

/** @param {Element} el */
function elementPlainSkippingDelSubtree(el) {
  if (el.namespaceURI === W_NS && el.localName === "del") {
    return "";
  }
  let s = "";
  for (const child of el.childNodes) {
    if (child.nodeType !== 1) continue;
    if (child.namespaceURI !== W_NS) continue;
    switch (child.localName) {
      case "t":
        s += child.textContent ?? "";
        break;
      case "delText":
        break;
      case "tab":
        s += "\t";
        break;
      case "br":
      case "cr":
        s += "\n";
        break;
      default:
        s += elementPlainSkippingDelSubtree(child);
        break;
    }
  }
  return s;
}

/** @param {Element} run */
function getRunInlinePlain(run) {
  let s = "";
  for (const child of run.childNodes) {
    if (child.nodeType !== 1) continue;
    if (child.namespaceURI !== W_NS) continue;
    switch (child.localName) {
      case "t":
        s += child.textContent ?? "";
        break;
      case "br":
      case "cr":
        s += "\n";
        break;
      case "tab":
        s += "\t";
        break;
      default:
        break;
    }
  }
  return s;
}

/** @param {Element} run */
function shouldSkipRun(run) {
  if (run.getElementsByTagNameNS(W_NS, "drawing").length) return true;
  if (run.getElementsByTagNameNS(W_NS, "object").length) return true;
  if (run.getElementsByTagNameNS(W_NS, "pict").length) return true;
  return false;
}

/** @param {Element} run @param {Element} paragraphEl */
function isInsideRevisionDelete(run, paragraphEl) {
  let n = run.parentElement;
  while (n && n !== paragraphEl) {
    if (n.namespaceURI === W_NS && n.localName === "del") return true;
    n = n.parentElement;
  }
  return false;
}

/** @param {Element} run */
function biuFromRun(run) {
  const rPr = run.getElementsByTagNameNS(W_NS, "rPr")[0];
  if (!rPr) {
    return { bold: false, italic: false, underline: false, key: "|||" };
  }
  const bold =
    !!rPr.getElementsByTagNameNS(W_NS, "b")[0] || !!rPr.getElementsByTagNameNS(W_NS, "bCs")[0];
  const italic =
    !!rPr.getElementsByTagNameNS(W_NS, "i")[0] || !!rPr.getElementsByTagNameNS(W_NS, "iCs")[0];
  const uEl = rPr.getElementsByTagNameNS(W_NS, "u")[0];
  const underline = uEl != null && uEl.getAttributeNS(W_NS, "val") !== "none" && uEl.getAttributeNS(W_NS, "val") !== "false";
  const key = `${bold}|${italic}|${underline}`;
  return { bold, italic, underline, key };
}

/**
 * Ordered chunks: merged text runs that share bold/italic/underline, or opaque runs (e.g. drawings).
 * @param {Element} pEl
 */
function buildChunks(pEl) {
  /** @type {({ type: 'text', plain: string, bold: boolean, italic: boolean, underline: boolean, key: string } | { type: 'opaque', el: Element })[]} */
  const chunks = [];
  const runs = [...pEl.getElementsByTagNameNS(W_NS, "r")];
  for (const r of runs) {
    if (isInsideRevisionDelete(r, pEl)) continue;
    if (shouldSkipRun(r)) {
      chunks.push({ type: "opaque", el: r });
      continue;
    }
    const plain = getRunInlinePlain(r);
    const { bold, italic, underline, key } = biuFromRun(r);
    const last = chunks[chunks.length - 1];
    if (last && last.type === "text" && last.key === key) {
      last.plain += plain;
    } else {
      chunks.push({ type: "text", plain, bold, italic, underline, key });
    }
  }
  return chunks;
}

/**
 * Non-empty logical lines across all text chunks (for progress totals).
 * @param {string} ooxmlParagraph
 */
export function countTranslatableLinesInParagraph(ooxmlParagraph) {
  try {
    const p = parseParagraphOoxml(ooxmlParagraph);
    let n = 0;
    for (const c of buildChunks(p)) {
      if (c.type !== "text") continue;
      for (const line of normText(c.plain).split("\n")) {
        if (line.trim()) n++;
      }
    }
    return n;
  } catch {
    return 0;
  }
}

function createRevisionIdAllocator() {
  const base = Math.floor(Math.random() * 2_000_000_000) + 1_000_000;
  let n = 0;
  return () => String(base + ++n);
}

/** @param {Element} el @param {string} author @param {() => string} nextId */
function stampRevisionMeta(el, author, nextId) {
  el.setAttributeNS(W_NS, "id", nextId());
  el.setAttributeNS(W_NS, "author", author);
  el.setAttributeNS(W_NS, "date", new Date().toISOString());
}

/**
 * @param {Document} doc
 * @param {boolean} bold
 * @param {boolean} italic
 * @param {boolean} underline
 * @param {string} text
 */
function makeRun(doc, bold, italic, underline, text) {
  const safe = sanitizeForWordXmlText(text);
  const r = doc.createElementNS(W_NS, "r");
  if (bold || italic || underline) {
    const rPr = doc.createElementNS(W_NS, "rPr");
    if (bold) rPr.appendChild(doc.createElementNS(W_NS, "b"));
    if (italic) rPr.appendChild(doc.createElementNS(W_NS, "i"));
    if (underline) {
      const u = doc.createElementNS(W_NS, "u");
      u.setAttributeNS(W_NS, "val", "single");
      rPr.appendChild(u);
    }
    r.appendChild(rPr);
  }
  const t = doc.createElementNS(W_NS, "t");
  if (/^\s|\s$|[\t\n]/.test(safe)) {
    t.setAttributeNS(XML_SPACE_NS, "xml:space", "preserve");
  }
  t.textContent = safe;
  r.appendChild(t);
  return r;
}

/**
 * Translate each logical line in order; batches consecutive single-part lines into `translateBatch` calls (≤ token budget per call).
 * Oversized single lines are split into sub-parts; each API call is capped by {@link translateOrderedStringsInBudget}.
 * @param {string} plain
 * @param {(batch: string[]) => Promise<string[]>} translateBatch
 * @param {(translatedLine: string) => void} [onTranslatedLine] once per original line (after rejoin if split)
 */
async function translatePlainByBatchedLines(plain, translateBatch, onTranslatedLine) {
  const lines = normText(plain).split("\n");
  const out = new Array(lines.length);
  const budget = contentTokenBudget();

  /** @type {{ lineIndex: number, text: string }[]} */
  let pending = [];
  let pendingCost = 0;

  async function flushPending() {
    if (!pending.length) return;
    const texts = pending.map((p) => p.text);
    const trans = await translateOrderedStringsInBudget(texts, translateBatch, budget);
    if (trans.length !== texts.length) {
      throw new Error(`Translation batch length mismatch: expected ${texts.length}, got ${trans.length}.`);
    }
    for (let j = 0; j < pending.length; j++) {
      const idx = pending[j].lineIndex;
      const t = String(trans[j] ?? "");
      out[idx] = t;
      if (onTranslatedLine) onTranslatedLine(t);
    }
    pending = [];
    pendingCost = 0;
  }

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (!line.trim()) {
      await flushPending();
      out[i] = line;
      continue;
    }

    const parts = splitOversizedLineMaxParts(line, maxPartTokensForSplit(), OVERSIZED_LINE_MAX_PARTS);
    if (parts.length > 1) {
      await flushPending();
      const transParts = await translateOrderedStringsInBudget(parts, translateBatch, budget);
      const joined = transParts.join("");
      out[i] = joined;
      if (onTranslatedLine) onTranslatedLine(joined);
      continue;
    }

    const cost = stringEntryCost(line);
    if (pending.length && pendingCost + cost > budget) {
      await flushPending();
    }
    pending.push({ lineIndex: i, text: line });
    pendingCost += cost;
  }

  await flushPending();
  return out.join("\n");
}

/**
 * @param {Element} origP parsed `w:p`
 * @param {Array<{ type: string, plain?: string, bold?: boolean, italic?: boolean, underline?: boolean, el?: Element }>} chunks
 * @param {{ orig: string, trans: string, bold: boolean, italic: boolean, underline: boolean }[]} textChunkPairs same order as text chunks only
 * @param {string} author
 */
function buildTrackedParagraphOoxmlFromChunks(origP, chunks, textChunkPairs, author) {
  const doc = origP.ownerDocument;
  const outP = doc.createElementNS(W_NS, "p");
  const pPr = origP.getElementsByTagNameNS(W_NS, "pPr")[0];
  if (pPr) {
    outP.appendChild(pPr.cloneNode(true));
  }

  const nextId = createRevisionIdAllocator();
  let pi = 0;
  for (const c of chunks) {
    if (c.type === "opaque") {
      outP.appendChild(doc.importNode(c.el, true));
      continue;
    }
    const step = textChunkPairs[pi++];
    if (!step) continue;
    if (normText(step.orig) === normText(step.trans)) {
      outP.appendChild(makeRun(doc, step.bold, step.italic, step.underline, step.trans));
      continue;
    }
    const del = doc.createElementNS(W_NS, "del");
    stampRevisionMeta(del, author, nextId);
    const dr = doc.createElementNS(W_NS, "r");
    const dt = doc.createElementNS(W_NS, "delText");
    dt.setAttributeNS(XML_SPACE_NS, "xml:space", "preserve");
    dt.textContent = sanitizeForWordXmlText(step.orig);
    dr.appendChild(dt);
    del.appendChild(dr);
    outP.appendChild(del);

    const ins = doc.createElementNS(W_NS, "ins");
    stampRevisionMeta(ins, author, nextId);
    ins.appendChild(makeRun(doc, step.bold, step.italic, step.underline, step.trans));
    outP.appendChild(ins);
  }

  let out = serializeParagraph(outP);
  out = scrubParagraphTextElements(out);
  if (!isWellFormedWordFragment(out)) {
    throw new Error(
      "Translation produced invalid OOXML after sanitizing control characters. Try translating a smaller selection, or use plain-text translation for this passage.",
    );
  }
  return out;
}

/**
 * @param {{ paraIndex: number, chunkIndex: number, lineIndex: number, strings: string[] }[]} lineGroups
 * @param {number} budget
 */
function packLineGroupsIntoBatches(lineGroups, budget) {
  /** @type {typeof lineGroups[]} */
  const batches = [];
  /** @type {typeof lineGroups} */
  let current = [];
  let currentCost = 0;

  /** @param {typeof lineGroups[0]} lg */
  function groupCost(lg) {
    return lg.strings.reduce((a, s) => a + stringEntryCost(s), 0);
  }

  for (const lg of lineGroups) {
    const gc = groupCost(lg);
    if (current.length > 0 && currentCost + gc > budget) {
      batches.push(current);
      current = [];
      currentCost = 0;
    }
    current.push(lg);
    currentCost += gc;
  }
  if (current.length) batches.push(current);
  return batches;
}

/**
 * Translate many body paragraphs with line batches spanning paragraphs (same token budget as per-chunk batching).
 * @param {string[]} ooxmlParagraphs OOXML `w:p` fragments in document order
 * @param {{
 *   translateBatch: (lines: string[]) => Promise<string[]>,
 *   author?: string,
 *   onTranslatedLine?: (translatedLine: string, meta?: { paraIndex: number, chunkIndex: number, lineIndex: number }) => void,
 *   signal?: AbortSignal,
 *   onExtractProgress?: (doneParas: number, totalParas: number) => void,
 *   onReadyToTranslate?: (lineCount: number, apiCallCount: number) => void,
 *   onBatchStart?: (requestIndex1: number, totalApiCalls: number, stringCount: number) => void,
 *   onBatchComplete?: (requestIndex1: number, totalApiCalls: number) => void,
 *   onMergeProgress?: (doneParas: number, totalParas: number) => void,
 * }} options
 * @returns {Promise<{ mergedOoxmls: string[], totalLineGroups: number }>}
 */
export async function translateParagraphsOoxmlGlobally(ooxmlParagraphs, options) {
  const {
    translateBatch,
    author = TRACK_TRANSLATION_AUTHOR,
    onTranslatedLine,
    signal,
    onExtractProgress,
    onReadyToTranslate,
    onBatchStart,
    onBatchComplete,
    onMergeProgress,
  } = options;
  const budget = contentTokenBudget();

  debugTrace("translate", `document translate: start, ${ooxmlParagraphs.length} paragraph(s), content budget ${budget}`);

  /** @type {{ paraIndex: number, chunkIndex: number, lineIndex: number, strings: string[] }[]} */
  const lineGroups = [];
  const parseOk = new Array(ooxmlParagraphs.length).fill(false);
  const totalParas = ooxmlParagraphs.length;

  for (let p = 0; p < totalParas; p++) {
    if (signal?.aborted) {
      throw new DOMException("Stopped.", "AbortError");
    }
    let origP;
    try {
      origP = parseParagraphOoxml(ooxmlParagraphs[p]);
    } catch {
      onExtractProgress?.(p + 1, totalParas);
      await new Promise((r) => setTimeout(r, 0));
      continue;
    }
    parseOk[p] = true;
    const chunks = buildChunks(origP);
    const hasTranslatableText = chunks.some((c) => c.type === "text" && c.plain.length > 0);
    if (!hasTranslatableText) {
      onExtractProgress?.(p + 1, totalParas);
      await new Promise((r) => setTimeout(r, 0));
      continue;
    }

    for (let ci = 0; ci < chunks.length; ci++) {
      const c = chunks[ci];
      if (c.type !== "text") continue;
      const lines = normText(c.plain).split("\n");
      for (let li = 0; li < lines.length; li++) {
        const line = lines[li];
        if (!line.trim()) continue;
        const parts = splitOversizedLineMaxParts(line, maxPartTokensForSplit(), OVERSIZED_LINE_MAX_PARTS);
        lineGroups.push({ paraIndex: p, chunkIndex: ci, lineIndex: li, strings: parts });
      }
    }
    onExtractProgress?.(p + 1, totalParas);
    await new Promise((r) => setTimeout(r, 0));
  }

  /** @type {Map<string, string>} */
  const lineTranslationMap = new Map();
  const batches = packLineGroupsIntoBatches(lineGroups, budget);

  let totalApiCalls = 0;
  for (const batch of batches) {
    const flat = [];
    for (const lg of batch) {
      for (const s of lg.strings) flat.push(s);
    }
    totalApiCalls += countApiCallsForOrderedStrings(flat, budget);
  }
  debugTrace(
    "translate",
    `plan: ${lineGroups.length} line group(s), ${batches.length} logical batch(es), ${totalApiCalls} API call(s)`,
  );
  onReadyToTranslate?.(lineGroups.length, totalApiCalls);

  let apiIndex = 0;
  for (let bi = 0; bi < batches.length; bi++) {
    const batch = batches[bi];
    if (signal?.aborted) {
      throw new DOMException("Stopped.", "AbortError");
    }
    const flat = [];
    for (const lg of batch) {
      for (const s of lg.strings) flat.push(s);
    }
    const translated = await translateOrderedStringsInBudget(flat, async (sub) => {
      apiIndex += 1;
      onBatchStart?.(apiIndex, totalApiCalls, sub.length);
      try {
        return await translateBatch(sub);
      } finally {
        onBatchComplete?.(apiIndex, totalApiCalls);
      }
    }, budget);
    if (translated.length !== flat.length) {
      throw new Error(`Translation batch length mismatch: expected ${flat.length}, got ${translated.length}.`);
    }
    let offset = 0;
    for (const lg of batch) {
      const n = lg.strings.length;
      const slice = translated.slice(offset, offset + n).map((x) => String(x ?? ""));
      offset += n;
      const joined = slice.join("");
      const key = `${lg.paraIndex}\t${lg.chunkIndex}\t${lg.lineIndex}`;
      lineTranslationMap.set(key, joined);
      if (onTranslatedLine) {
        onTranslatedLine(joined, {
          paraIndex: lg.paraIndex,
          chunkIndex: lg.chunkIndex,
          lineIndex: lg.lineIndex,
        });
      }
    }
  }

  /** @param {number} p */
  function mergedParagraphOoxml(p) {
    if (!parseOk[p]) {
      return ooxmlParagraphs[p];
    }
    let origP;
    try {
      origP = parseParagraphOoxml(ooxmlParagraphs[p]);
    } catch {
      return ooxmlParagraphs[p];
    }
    const chunks = buildChunks(origP);
    const hasTranslatableText = chunks.some((c) => c.type === "text" && c.plain.length > 0);
    if (!hasTranslatableText) {
      return ooxmlParagraphs[p];
    }

    /** @type {{ orig: string, trans: string, bold: boolean, italic: boolean, underline: boolean }[]} */
    const pairs = [];
    for (let ci = 0; ci < chunks.length; ci++) {
      const c = chunks[ci];
      if (c.type !== "text") continue;
      const lines = normText(c.plain).split("\n");
      const outLines = [];
      for (let li = 0; li < lines.length; li++) {
        const line = lines[li];
        if (!line.trim()) {
          outLines.push(line);
        } else {
          const key = `${p}\t${ci}\t${li}`;
          if (!lineTranslationMap.has(key)) {
            throw new Error(`Missing translation for paragraph ${p + 1}, chunk ${ci}, line ${li}.`);
          }
          outLines.push(lineTranslationMap.get(key) ?? "");
        }
      }
      pairs.push({
        orig: c.plain,
        trans: outLines.join("\n"),
        bold: c.bold,
        italic: c.italic,
        underline: c.underline,
      });
    }

    return buildTrackedParagraphOoxmlFromChunks(origP, chunks, pairs, author);
  }

  /** @type {string[]} */
  const mergedOoxmls = new Array(ooxmlParagraphs.length);
  for (let p = 0; p < ooxmlParagraphs.length; p++) {
    if (signal?.aborted) {
      throw new DOMException("Stopped.", "AbortError");
    }
    mergedOoxmls[p] = mergedParagraphOoxml(p);
    onMergeProgress?.(p + 1, ooxmlParagraphs.length);
    await new Promise((r) => setTimeout(r, 0));
  }

  return { mergedOoxmls, totalLineGroups: lineGroups.length };
}

/**
 * Build a new `w:p` OOXML string: per text chunk (BIU group), lines batched for translation,
 * `w:del`+`w:ins` when text changes; opaque runs copied unchanged.
 * @param {string} ooxmlParagraph
 * @param {{
 *   translateBatch: (lines: string[]) => Promise<string[]>,
 *   author?: string,
 *   onTranslatedLine?: (translatedLine: string) => void,
 * }} options
 * @returns {Promise<string>}
 */
export async function translateParagraphBiuTracked(ooxmlParagraph, options) {
  const {
    translateBatch,
    author = TRACK_TRANSLATION_AUTHOR,
    onTranslatedLine,
  } = options;
  let origP;
  try {
    origP = parseParagraphOoxml(ooxmlParagraph);
  } catch {
    return ooxmlParagraph;
  }
  const chunks = buildChunks(origP);
  const hasTranslatableText = chunks.some((c) => c.type === "text" && c.plain.length > 0);
  if (!hasTranslatableText) {
    return ooxmlParagraph;
  }

  /** @type {{ orig: string, trans: string, bold: boolean, italic: boolean, underline: boolean }[]} */
  const pairs = [];
  for (const c of chunks) {
    if (c.type !== "text") continue;
    const trans = await translatePlainByBatchedLines(c.plain, translateBatch, onTranslatedLine);
    pairs.push({
      orig: c.plain,
      trans,
      bold: c.bold,
      italic: c.italic,
      underline: c.underline,
    });
  }

  return buildTrackedParagraphOoxmlFromChunks(origP, chunks, pairs, author);
}

/**
 * Quick parse check so we never hand Word a string that breaks its temp document.
 * @param {string} xml
 */
function isWellFormedWordFragment(xml) {
  const clean = stripXmlDecl(String(xml || "").trim());
  if (!clean) return false;
  const d = new DOMParser().parseFromString(clean, "application/xml");
  return d.getElementsByTagName("parsererror").length === 0;
}

/**
 * Total non-empty lines across all `w:p` in an OOXML fragment (for progress).
 * @param {string} ooxml
 */
export function countTranslatableLinesInOoxmlTree(ooxml) {
  const clean = stripXmlDecl(ooxml);
  const dom = new DOMParser().parseFromString(clean, "application/xml");
  if (dom.getElementsByTagName("parsererror").length) return 0;
  const paras = [...dom.getElementsByTagNameNS(W_NS, "p")];
  let n = 0;
  for (const p of paras) {
    n += countTranslatableLinesInParagraph(serializeParagraph(p));
  }
  return n;
}

/**
 * Translate every `w:p` in an OOXML tree in document order (e.g. range or package from `Range.getOoxml()`),
 * replacing each paragraph node in the DOM and serializing the root element back.
 * @param {string} ooxml
 * @param {{
 *   translateBatch: (lines: string[]) => Promise<string[]>,
 *   author?: string,
 *   onTranslatedLine?: (translatedLine: string) => void,
 *   signal?: AbortSignal,
 *   onExtractProgress?: (doneParas: number, totalParas: number) => void,
 *   onReadyToTranslate?: (lineCount: number, apiCallCount: number) => void,
 *   onBatchStart?: (requestIndex1: number, totalApiCalls: number, stringCount: number) => void,
 *   onBatchComplete?: (requestIndex1: number, totalApiCalls: number) => void,
 *   onMergeProgress?: (doneParas: number, totalParas: number) => void,
 * }} options same as {@link translateParagraphsOoxmlGlobally}
 * @returns {Promise<string | null>} Serialized XML, or `null` if the tree contains no `w:p` (use plain-text path).
 */
export async function translateEveryParagraphInOoxmlTree(ooxml, options) {
  const clean = stripXmlDecl(ooxml);
  const dom = new DOMParser().parseFromString(clean, "application/xml");
  if (dom.getElementsByTagName("parsererror").length) {
    throw new Error("OOXML parse error");
  }
  const paras = [...dom.getElementsByTagNameNS(W_NS, "p")];
  if (paras.length === 0) {
    return null;
  }
  const ser = paras.map((p) => serializeParagraph(p));
  const { mergedOoxmls: mergedList } = await translateParagraphsOoxmlGlobally(ser, options);
  for (let i = 0; i < paras.length; i++) {
    const merged = mergedList[i];
    const newP = parseParagraphOoxml(merged);
    paras[i].parentNode.replaceChild(dom.importNode(newP, true), paras[i]);
  }
  const finalXml = new XMLSerializer().serializeToString(dom.documentElement);
  const reparsed = new DOMParser().parseFromString(finalXml, "application/xml");
  if (reparsed.getElementsByTagName("parsererror").length) {
    throw new Error(
      "Translated selection produced invalid OOXML. Try a smaller selection, or remove unusual characters from the source text.",
    );
  }
  return finalXml;
}
