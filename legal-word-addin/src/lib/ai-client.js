import { loadSettings } from "./ai-settings.js";
import { clampMaxTokens } from "./ai-limits.js";
import { recordChatTurnUsage, extractBillableTokensFromResponse } from "./token-usage.js";
import { debugTrace, debugTraceError, truncateForLog } from "./debug-trace.js";
import { appLog, appWarn } from "./app-log.js";

/** Infomaniak LLM API: ~1 req / 2s; see developer.infomaniak.com LLM docs. */
const INFA_MANIAK_DEFAULT_INTERVAL_MS = 2000;

let lastLlmRequestEndedAt = 0;
/** @type {Promise<unknown>} */
let llmRequestChain = Promise.resolve();

/** @param {Record<string, unknown>} settings @returns {number} */
function effectiveLlmMinIntervalMs(settings) {
  const raw = settings.llmMinIntervalMs;
  if (raw !== undefined && raw !== null) {
    const n = Number(raw);
    if (Number.isFinite(n) && n >= 0) return Math.min(120_000, Math.floor(n));
  }
  const ap = settings.activeProvider;
  const p = settings.providers?.[ap];
  const urls = [String(settings.proxyUrl || ""), String(p?.baseUrl || "")];
  if (urls.some((u) => /infomaniak\.com/i.test(u))) {
    return INFA_MANIAK_DEFAULT_INTERVAL_MS;
  }
  return 0;
}

/** @param {Record<string, unknown>} settings @param {AbortSignal | undefined} signal */
async function throttleBeforeNextLlmCall(settings, signal) {
  const minGap = effectiveLlmMinIntervalMs(settings);
  if (minGap <= 0) return;
  if (lastLlmRequestEndedAt <= 0) return;
  const elapsed = Date.now() - lastLlmRequestEndedAt;
  const wait = minGap - elapsed;
  if (wait > 0) {
    appLog("llm", `throttle ${Math.round(wait)}ms before next request`, { minGapMs: minGap });
    debugTrace("llm", `throttle: waiting ${Math.round(wait)}ms (min gap ${minGap}ms since last call)`);
    if (signal?.aborted) {
      throw new DOMException("Stopped.", "AbortError");
    }
    await new Promise((resolve, reject) => {
      const t = setTimeout(resolve, wait);
      if (!signal) return;
      const onAbort = () => {
        clearTimeout(t);
        reject(new DOMException("Stopped.", "AbortError"));
      };
      if (signal.aborted) {
        clearTimeout(t);
        reject(new DOMException("Stopped.", "AbortError"));
        return;
      }
      signal.addEventListener("abort", onAbort, { once: true });
    });
  }
}

/**
 * @param {{ role: string, content: string }[]} messages
 * @param {{ temperature?: number, maxTokens?: number, signal?: AbortSignal }} [opts]
 */
export async function completeChat(messages, opts = {}) {
  const run = async () => {
    const settings = loadSettings();
    const signal = opts.signal;
    if (signal?.aborted) {
      throw new DOMException("Stopped.", "AbortError");
    }
    await throttleBeforeNextLlmCall(settings, signal);
    try {
      return await executeCompleteChat(messages, opts, settings);
    } finally {
      lastLlmRequestEndedAt = Date.now();
    }
  };
  const p = llmRequestChain.then(run, run);
  llmRequestChain = p.catch(() => {});
  return p;
}

/**
 * @param {{ role: string, content: string }[]} messages
 * @param {{ temperature?: number, maxTokens?: number, signal?: AbortSignal }} opts
 * @param {Record<string, unknown>} settings
 */
async function executeCompleteChat(messages, opts, settings) {
  const signal = opts.signal;
  if (signal?.aborted) {
    throw new DOMException("Stopped.", "AbortError");
  }
  const { activeProvider, requestMode, proxyUrl, providers } = settings;
  const p = providers[activeProvider];
  if (activeProvider === "openai" && providers.openai.enabled === false) {
    throw new Error('OpenAI is turned off in AI connections. Enable "Allow OpenAI" or pick another provider.');
  }
  if (!p?.apiKey?.trim()) {
    throw new Error("Add an API key in Settings for the active provider.");
  }

  const maxTokens = clampMaxTokens(opts.maxTokens);
  const inputChars = messages.reduce((a, m) => a + String(m.content ?? "").length, 0);
  appLog("llm", "completeChat → provider", {
    provider: activeProvider,
    mode: requestMode,
    model: p.model,
    messageCount: messages.length,
    inputChars,
    maxTokens,
  });
  debugTrace("llm", `completeChat start`, {
    provider: activeProvider,
    mode: requestMode,
    model: p.model,
    messages: messages.length,
    inputChars,
    maxTokens,
  });

  const body = {
    provider: activeProvider,
    model: p.model,
    messages,
    temperature: opts.temperature ?? 0.2,
    max_tokens: maxTokens,
    apiKey: p.apiKey,
    baseUrl: p.baseUrl || undefined,
    deployment: p.deployment || undefined,
    apiVersion: p.apiVersion || undefined,
    openaiAllowed: providers.openai.enabled !== false,
    organizationId: activeProvider === "openai" ? providers.openai.organizationId?.trim() || undefined : undefined,
    projectId: activeProvider === "openai" ? providers.openai.projectId?.trim() || undefined : undefined,
  };

  try {
    if (requestMode === "proxy") {
      const origin = (proxyUrl || "").trim().replace(/\/$/, "");
      const url = origin ? `${origin}/api/legal-ai/complete` : "/api/legal-ai/complete";
      debugTrace("llm", `POST proxy ${url}`);
      const res = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
        signal,
      });
      if (!res.ok) {
        const t = await res.text();
        debugTrace("llm", `proxy HTTP ${res.status}`, truncateBodyForLog(t));
        throw formatProviderHttpError(t, res.statusText);
      }
      const data = await res.json();
      ensureChatSuccessPayload(data);
      const reply = data.text || data.content || "";
      const apiUsage = extractBillableTokensFromResponse(data);
      debugTrace("llm", `proxy OK`, {
        http: res.status,
        replyChars: reply.length,
        replyPreview: truncateForLog(reply, 200),
        usageFromJson: apiUsage || "(none — estimate may be used)",
      });
      recordChatTurnUsage(data, messages, reply);
      appLog("llm", "completeChat ← OK (proxy)", { replyChars: reply.length, http: res.status });
      return reply;
    }

    const out = await directComplete(activeProvider, p, messages, { ...opts, maxTokens });
    return out;
  } catch (e) {
    appWarn("llm", "completeChat failed", e instanceof Error ? e.message : String(e));
    debugTraceError("llm", e);
    throw e;
  }
}

/** @param {string} t */
function truncateBodyForLog(t) {
  const s = String(t || "");
  return s.length > 500 ? `${s.slice(0, 500)}…` : s;
}

/** @param {{ code?: string, description?: string }} err */
function providerErrorToError(err) {
  const code = err?.code != null ? String(err.code) : "error";
  const desc = err?.description != null ? String(err.description) : JSON.stringify(err);
  return new Error(`${code}: ${desc}`);
}

/** @param {unknown} data */
function ensureChatSuccessPayload(data) {
  const d = /** @type {Record<string, unknown>} */ (data);
  if (d?.result === "error" && d.error) {
    throw providerErrorToError(/** @type {{ code?: string, description?: string }} */ (d.error));
  }
}

/**
 * @param {string} bodyText
 * @param {string} statusText
 */
function formatProviderHttpError(bodyText, statusText) {
  const t = (bodyText || statusText || "").trim();
  if (t) {
    try {
      const j = JSON.parse(t);
      if (j?.result === "error" && j.error) return providerErrorToError(/** @type {any} */ (j.error));
      if (j?.error && (typeof j.error === "object") && (j.error.code || j.error.description)) {
        return providerErrorToError(/** @type {any} */ (j.error));
      }
    } catch {
      /* not JSON */
    }
  }
  return new Error(t || statusText || "Request failed");
}

/**
 * @param {string} activeProvider
 * @param {{ apiKey: string, model: string, baseUrl?: string, deployment?: string, apiVersion?: string, organizationId?: string, projectId?: string }} p
 * @param {{ role: string, content: string }[]} messages
 * @param {{ temperature?: number, maxTokens?: number, signal?: AbortSignal }} opts
 */
async function directComplete(activeProvider, p, messages, opts) {
  const maxTokens = clampMaxTokens(opts.maxTokens);
  const signal = opts.signal;
  if (activeProvider === "anthropic") {
    const url = (p.baseUrl && p.baseUrl.trim()) || "https://api.anthropic.com/v1/messages";
    debugTrace("llm", `POST anthropic ${url}`);
    const res = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": p.apiKey,
        "anthropic-version": "2023-06-01",
      },
      signal,
      body: JSON.stringify({
        model: p.model,
        max_tokens: maxTokens,
        temperature: opts.temperature ?? 0.2,
        messages: messages.filter((m) => m.role !== "system").map((m) => ({ role: m.role, content: m.content })),
        system: messages.find((m) => m.role === "system")?.content,
      }),
    });
    if (!res.ok) throw formatProviderHttpError(await res.text(), res.statusText);
    const data = await res.json();
    const block = data.content?.[0];
    const reply = block?.type === "text" ? block.text : JSON.stringify(data);
    const apiUsage = extractBillableTokensFromResponse(data);
    debugTrace("llm", `anthropic OK`, {
      http: res.status,
      replyChars: reply.length,
      replyPreview: truncateForLog(reply, 200),
      usageFromJson: apiUsage || "(none — estimate may be used)",
    });
    recordChatTurnUsage(data, messages, reply);
    appLog("llm", "completeChat ← OK (anthropic)", { replyChars: reply.length, http: res.status });
    return reply;
  }

  if (activeProvider === "azureOpenai") {
    const endpoint = (p.baseUrl || "").replace(/\/$/, "");
    const deployment = p.deployment || p.model;
    const ver = p.apiVersion || "2024-02-15-preview";
    const url = `${endpoint}/openai/deployments/${deployment}/chat/completions?api-version=${ver}`;
    debugTrace("llm", `POST azure ${url}`);
    const res = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "api-key": p.apiKey,
      },
      signal,
      body: JSON.stringify({
        messages,
        temperature: opts.temperature ?? 0.2,
        max_tokens: maxTokens,
      }),
    });
    if (!res.ok) throw formatProviderHttpError(await res.text(), res.statusText);
    const data = await res.json();
    const reply = data.choices?.[0]?.message?.content ?? "";
    const apiUsage = extractBillableTokensFromResponse(data);
    debugTrace("llm", `azure OK`, {
      http: res.status,
      replyChars: reply.length,
      replyPreview: truncateForLog(reply, 200),
      usageFromJson: apiUsage || "(none — estimate may be used)",
    });
    recordChatTurnUsage(data, messages, reply);
    appLog("llm", "completeChat ← OK (azure)", { replyChars: reply.length, http: res.status });
    return reply;
  }

  const base =
    activeProvider === "customOpenAI"
      ? (p.baseUrl || "").replace(/\/$/, "")
      : (p.baseUrl || "").trim()
        ? String(p.baseUrl).replace(/\/$/, "")
        : "https://api.openai.com/v1";
  const url = `${base}/chat/completions`;
  debugTrace("llm", `POST openai-compatible ${url}`, { model: p.model });
  /** @type {Record<string, string>} */
  const headers = {
    "Content-Type": "application/json",
    Authorization: `Bearer ${p.apiKey}`,
  };
  if (activeProvider === "openai") {
    if (p.organizationId?.trim()) headers["OpenAI-Organization"] = p.organizationId.trim();
    if (p.projectId?.trim()) headers["OpenAI-Project"] = p.projectId.trim();
  }
  const res = await fetch(url, {
    method: "POST",
    headers,
    signal,
    body: JSON.stringify({
      model: p.model,
      messages,
      temperature: opts.temperature ?? 0.2,
      max_tokens: maxTokens,
    }),
  });
  if (!res.ok) throw formatProviderHttpError(await res.text(), res.statusText);
  const data = await res.json();
  ensureChatSuccessPayload(data);
  const reply = data.choices?.[0]?.message?.content ?? "";
  const apiUsage = extractBillableTokensFromResponse(data);
  debugTrace("llm", `openai-compatible OK`, {
    http: res.status,
    replyChars: reply.length,
    replyPreview: truncateForLog(reply, 200),
    usageFromJson: apiUsage || "(none — estimate may be used)",
  });
  recordChatTurnUsage(data, messages, reply);
  appLog("llm", "completeChat ← OK (openai-compatible)", {
    replyChars: reply.length,
    http: res.status,
    model: p.model,
  });
  return reply;
}
