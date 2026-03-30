const STORAGE_KEY = "legal-word-addin:ai-settings-v1";

/**
 * @typedef {Object} OpenAiLikeConfig
 * @property {string} apiKey
 * @property {string} model
 * @property {string} [baseUrl]
 * @property {string} [deployment]
 * @property {string} [apiVersion]
 */

/**
 * @typedef {Object} OpenAiProviderConfig
 * @property {boolean} enabled Allow routing requests to OpenAI (platform API).
 * @property {string} apiKey
 * @property {string} model
 * @property {string} baseUrl Optional; default https://api.openai.com/v1
 * @property {string} organizationId Optional OpenAI-Organization header (org-…).
 * @property {string} projectId Optional OpenAI-Project header (proj_…).
 */

/**
 * @typedef {Object} AppAiSettings
 * @property {'openai'|'anthropic'|'azureOpenai'|'customOpenAI'} activeProvider
 * @property {'proxy'|'direct'} requestMode
 * @property {string} proxyUrl Backend base URL for proxy mode (empty = same origin, e.g. webpack → local server).
 * @property {number} [llmMinIntervalMs] Min ms between LLM HTTP requests (0 = none). Omit for auto (e.g. 2000 for Infomaniak).
 * @property {'llm'|'presidio'} piiEngine Which engine to use for PII detection in anonymize. "llm" = active LLM provider, "presidio" = Microsoft Presidio sidecar.
 * @property {string} [presidioBackendUrl] Direct-mode URL for the Presidio Python sidecar (e.g. http://localhost:3549). Leave empty in proxy mode.
 * @property {{ openai: OpenAiProviderConfig, anthropic: OpenAiLikeConfig, azureOpenai: OpenAiLikeConfig, customOpenAI: OpenAiLikeConfig }} providers
 */

/** @returns {AppAiSettings} */
export function defaultSettings() {
  return {
    activeProvider: "openai",
    requestMode: "proxy",
    proxyUrl: "",
    piiEngine: "llm",
    presidioBackendUrl: "",
    providers: {
      openai: {
        enabled: true,
        apiKey: "",
        model: "gpt-4o",
        baseUrl: "",
        organizationId: "",
        projectId: "",
      },
      anthropic: {
        apiKey: "",
        model: "claude-sonnet-4-20250514",
        baseUrl: "",
      },
      azureOpenai: {
        apiKey: "",
        model: "",
        baseUrl: "",
        deployment: "",
        apiVersion: "2024-02-15-preview",
      },
      customOpenAI: {
        apiKey: "",
        model: "gpt-4o",
        baseUrl: "https://api.example.com/v1",
      },
    },
  };
}

function mergeProviders(parsedProviders) {
  const def = defaultSettings().providers;
  /** @type {typeof def} */
  const out = { ...def };
  for (const key of Object.keys(def)) {
    out[key] = { ...def[key], ...(parsedProviders?.[key] && typeof parsedProviders[key] === "object" ? parsedProviders[key] : {}) };
  }
  return out;
}

export function loadSettings() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return defaultSettings();
    const parsed = JSON.parse(raw);
    const base = defaultSettings();
    return {
      ...base,
      ...parsed,
      providers: mergeProviders(parsed.providers),
    };
  } catch {
    return defaultSettings();
  }
}

/** @param {AppAiSettings} s */
export function saveSettings(s) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(s));
}

export function maskKey(key) {
  if (!key || key.length < 8) return key ? "••••••••" : "";
  return `${key.slice(0, 4)}…${key.slice(-4)}`;
}
