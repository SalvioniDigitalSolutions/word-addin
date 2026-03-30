import "dotenv/config";
import express from "express";
import cors from "cors";

const MAX_COMPLETION_TOKENS = 5000;

const PORT = parseInt(process.env.PORT || "3548", 10);
const app = express();

app.use(cors({ origin: true }));
app.use(express.json({ limit: "4mb" }));

/**
 * @param {import("express").Request} req
 * @param {import("express").Response} res
 */
app.post("/api/legal-ai/complete", async (req, res) => {
  try {
    let {
      provider,
      model,
      messages,
      temperature = 0.2,
      max_tokens: maxTokensRaw = 4096,
      apiKey,
      baseUrl,
      deployment,
      apiVersion,
      openaiAllowed,
      organizationId,
      projectId,
    } = req.body || {};

    if (provider === "openai" && openaiAllowed === false) {
      res.status(400).type("text").send('OpenAI is turned off in AI connections.');
      return;
    }

    if (!provider || !Array.isArray(messages)) {
      res.status(400).type("text").send("Invalid request body.");
      return;
    }

    const key = typeof apiKey === "string" ? apiKey.trim() : "";
    if (!key) {
      res.status(400).type("text").send("Missing API key.");
      return;
    }

    const max_tokens = Math.min(
      MAX_COMPLETION_TOKENS,
      Math.max(1, Math.floor(Number(maxTokensRaw)) || 4096),
    );

    let text = "";
    /** @type {unknown} */
    let usagePayload;

    if (provider === "anthropic") {
      const url = (baseUrl && String(baseUrl).trim()) || "https://api.anthropic.com/v1/messages";
      const r = await fetch(url, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": key,
          "anthropic-version": "2023-06-01",
        },
        body: JSON.stringify({
          model,
          max_tokens,
          temperature,
          messages: messages.filter((m) => m.role !== "system").map((m) => ({ role: m.role, content: m.content })),
          system: messages.find((m) => m.role === "system")?.content,
        }),
      });
      if (!r.ok) {
        res.status(r.status).type("text").send(await r.text());
        return;
      }
      const data = await r.json();
      const block = data.content?.[0];
      text = block?.type === "text" ? block.text : JSON.stringify(data);
      usagePayload = data.usage;
    } else if (provider === "azureOpenai") {
      const endpoint = String(baseUrl || "").replace(/\/$/, "");
      const dep = deployment || model;
      const ver = apiVersion || "2024-02-15-preview";
      const url = `${endpoint}/openai/deployments/${dep}/chat/completions?api-version=${ver}`;
      const r = await fetch(url, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "api-key": key,
        },
        body: JSON.stringify({
          messages,
          temperature,
          max_tokens,
        }),
      });
      if (!r.ok) {
        res.status(r.status).type("text").send(await r.text());
        return;
      }
      const data = await r.json();
      text = data.choices?.[0]?.message?.content ?? "";
      usagePayload = data.usage;
    } else {
      const base =
        provider === "customOpenAI"
          ? String(baseUrl || "").replace(/\/$/, "")
          : String(baseUrl || "").trim()
            ? String(baseUrl).replace(/\/$/, "")
            : "https://api.openai.com/v1";
      const url = `${base}/chat/completions`;
      /** @type {Record<string, string>} */
      const headers = {
        "Content-Type": "application/json",
        Authorization: `Bearer ${key}`,
      };
      if (provider === "openai") {
        if (organizationId?.trim()) headers["OpenAI-Organization"] = organizationId.trim();
        if (projectId?.trim()) headers["OpenAI-Project"] = projectId.trim();
      }
      const r = await fetch(url, {
        method: "POST",
        headers,
        body: JSON.stringify({
          model,
          messages,
          temperature,
          max_tokens,
        }),
      });
      if (!r.ok) {
        res.status(r.status).type("text").send(await r.text());
        return;
      }
      const data = await r.json();
      text = data.choices?.[0]?.message?.content ?? "";
      usagePayload = data.usage;
    }

    res.json({ text, content: text, usage: usagePayload });
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    res.status(500).type("text").send(msg);
  }
});

const PRESIDIO_PORT = parseInt(process.env.PRESIDIO_PORT || "3549", 10);
const PRESIDIO_SCORE_THRESHOLD = parseFloat(process.env.PRESIDIO_SCORE_THRESHOLD || "0.32");

app.post("/api/legal-ai/presidio-pii", async (req, res) => {
  try {
    const { text } = req.body || {};
    if (!text || typeof text !== "string") {
      res.status(400).json({ error: "Missing 'text' field." });
      return;
    }
    const url = `http://127.0.0.1:${PRESIDIO_PORT}/scan`;
    const r = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text, score_threshold: PRESIDIO_SCORE_THRESHOLD }),
    });
    if (!r.ok) {
      const body = await r.text();
      res.status(r.status).type("text").send(body || `Presidio returned HTTP ${r.status}`);
      return;
    }
    const data = await r.json();
    res.json(data);
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    res.status(502).json({ error: `Could not reach Presidio sidecar on port ${PRESIDIO_PORT}: ${msg}` });
  }
});

app.listen(PORT, "127.0.0.1", () => {
  console.log(`Legal add-in proxy listening on http://127.0.0.1:${PORT}`);
});
