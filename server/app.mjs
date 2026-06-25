import fs from "node:fs";
import http from "node:http";
import https from "node:https";
import os from "node:os";
import path from "node:path";
import express from "express";

import { ChatService } from "./chat-service.mjs";
import { config, getBaseUrl } from "./config.mjs";
import { ExcelMcpClient } from "./mcp-client.mjs";
import { ProviderStore } from "./provider-store.mjs";

const app = express();
const mcpClient = new ExcelMcpClient({
  command: config.mcpCommand,
  args: config.mcpArgs,
  cwd: config.mcpCwd,
});
const providerStore = new ProviderStore({
  projectRoot: config.projectRoot,
  env: process.env,
});
const chatService = new ChatService({
  mcpClient,
  providerStore,
});

app.use(express.json({ limit: "1mb" }));
app.use(express.static(config.publicDir));

app.get("/health", async (_req, res) => {
  try {
    const tools = await mcpClient.listTools();
    res.json({
      ok: true,
      openaiConfigured: chatService.isConfigured(),
      model: providerStore.read().model,
      baseUrl: getBaseUrl(),
      mcpTools: tools.length,
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      message: error.message,
    });
  }
});

app.get("/api/config", async (_req, res) => {
  const tools = await mcpClient.listTools();
  res.json({
    openaiConfigured: chatService.isConfigured(),
    model: providerStore.read().model,
    provider: providerStore.publicSettings(),
    baseUrl: getBaseUrl(),
    manifestPath: config.manifestPath,
    mcpTools: tools.length,
  });
});

app.get("/api/provider-settings", (req, res) => {
  if (req.query.provider) {
    const profile = providerStore.publicProfile(String(req.query.provider));
    res.json(profile || { providerName: String(req.query.provider), configured: false });
    return;
  }

  res.json(providerStore.publicSettings());
});

app.get("/api/provider-profiles", (_req, res) => {
  res.json({
    activeProvider: providerStore.read().providerName,
    profiles: providerStore.publicProfiles(),
  });
});

app.get("/api/model-options", async (_req, res) => {
  const profiles = providerStore.privateProfiles();
  const active = providerStore.read();
  const sourceProfiles = Object.keys(profiles).length ? profiles : { [active.providerName]: active };
  const groups = [];

  for (const settings of Object.values(sourceProfiles)) {
    if (!settings.apiKey && !isLocalProvider(settings)) {
      continue;
    }

    try {
      const rawModels = await fetchProviderModels(settings);
      const models = displayModelsForProvider(settings, rawModels);
      groups.push({
        providerName: settings.providerName,
        baseURL: settings.baseURL,
        currentModel: settings.model,
        models,
      });
    } catch (error) {
      groups.push({
        providerName: settings.providerName,
        baseURL: settings.baseURL,
        currentModel: settings.model,
        error: error.message,
        models: settings.model
          ? [
              {
                model: settings.model,
                label: labelForModel(settings, settings.model),
              },
            ]
          : [],
      });
    }
  }

  res.json({
    activeProvider: active.providerName,
    activeModel: active.model,
    groups,
  });
});

app.get("/api/models", async (_req, res) => {
  try {
    const settings = providerStore.read();
    const models = await fetchProviderModels(settings);
    res.json({
      ok: true,
      providerName: settings.providerName,
      baseURL: settings.baseURL,
      currentModel: settings.model,
      models,
    });
  } catch (error) {
    res.status(502).json({
      ok: false,
      error: error.message,
      currentModel: providerStore.read().model,
      models: [],
    });
  }
});

app.post("/api/provider-settings", (req, res) => {
  const current = providerStore.read();
  const incoming = req.body || {};
  const existingProfile = incoming.providerName ? providerStore.privateProfile(incoming.providerName) : null;
  const next = providerStore.write({
    ...incoming,
    apiKey:
      incoming.apiKey && !isMaskedSecret(incoming.apiKey)
        ? incoming.apiKey
        : (existingProfile?.apiKey || current.apiKey),
  });

  res.json({
    ...providerStore.publicSettings(),
    model: next.model,
  });
});

app.post("/api/active-model", (req, res) => {
  try {
    const { providerName, model } = req.body || {};
    if (!providerName || !model) {
      res.status(400).json({ error: "providerName and model are required." });
      return;
    }

    const next = providerStore.activateModel({ providerName, model });
    res.json(publicProviderResponse(next));
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

app.post("/api/models", async (req, res) => {
  try {
    const current = providerStore.read();
    const incoming = req.body || {};
    const existingProfile = incoming.providerName ? providerStore.privateProfile(incoming.providerName) : null;
    const settings = {
      ...current,
      ...(existingProfile || {}),
      ...incoming,
      apiKey:
        incoming.apiKey && !isMaskedSecret(incoming.apiKey)
          ? incoming.apiKey
          : (existingProfile?.apiKey || current.apiKey),
    };
    const models = await fetchProviderModels(settings);
    res.json({
      ok: true,
      providerName: settings.providerName,
      baseURL: settings.baseURL,
      currentModel: settings.model,
      models,
    });
  } catch (error) {
    res.status(502).json({
      ok: false,
      error: error.message,
      currentModel: providerStore.read().model,
      models: [],
    });
  }
});

app.post("/api/chat", async (req, res) => {
  const { history = [], message = "", workbookContext = {}, images = [], providerSettings = null } = req.body || {};

  if (!message.trim()) {
    return res.status(400).json({ error: "message is required" });
  }

  try {
    const result = await chatService.runConversation({
      history,
      message,
      workbookContext,
      images,
      providerSettings,
    });
    res.json(result);
  } catch (error) {
    res.status(500).json({
      error: error.message,
    });
  }
});

app.get("/manifest.xml", (_req, res) => {
  res.type("application/xml");
  res.send(fs.readFileSync(config.manifestPath, "utf8"));
});

app.get("/", (_req, res) => {
  res.redirect("/taskpane.html");
});

async function start() {
  const server =
    config.protocol === "https"
      ? https.createServer(getStaticHttpsOptions(), app)
      : http.createServer(app);

  server.listen(config.port, config.host, () => {
    console.log(`Excel chatbot add-in server running at ${getBaseUrl()}`);
    console.log(`Manifest: ${path.join(config.projectRoot, "manifest.xml")}`);
  });
}

function getStaticHttpsOptions() {
  const localCertDir = path.join(config.projectRoot, ".certs");
  const pfxPath = path.join(localCertDir, "localhost.pfx");
  const pfxPasswordPath = path.join(localCertDir, "pfx-password.txt");

  if (fs.existsSync(pfxPath) && fs.existsSync(pfxPasswordPath)) {
    return {
      pfx: fs.readFileSync(pfxPath),
      passphrase: fs.readFileSync(pfxPasswordPath, "utf8").trim(),
    };
  }

  const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
  const caPath = path.join(certDir, "ca.crt");
  const certPath = path.join(certDir, "localhost.crt");
  const keyPath = path.join(certDir, "localhost.key");

  for (const requiredPath of [caPath, certPath, keyPath]) {
    if (!fs.existsSync(requiredPath)) {
      throw new Error(`Missing HTTPS certificate file: ${requiredPath}`);
    }
  }

  return {
    ca: fs.readFileSync(caPath),
    cert: fs.readFileSync(certPath),
    key: fs.readFileSync(keyPath),
  };
}

function publicProviderResponse(settings) {
  return {
    ...settings,
    apiKey: settings.apiKey ? maskSecret(settings.apiKey) : "",
    configured: Boolean(settings.apiKey && settings.model),
  };
}

function maskSecret(value) {
  if (!value) {
    return "";
  }

  if (value.length <= 8) {
    return "********";
  }

  return `${value.slice(0, 4)}...${value.slice(-4)}`;
}

function isMaskedSecret(value) {
  return String(value || "").includes("*") || String(value || "").includes("...");
}

function isLocalProvider(settings) {
  return /^https?:\/\/(127\.0\.0\.1|localhost)(:\d+)?\//i.test(`${settings.baseURL}/`);
}

function displayModelsForProvider(settings, rawModels) {
  const providerKey = settings.providerName.toLowerCase();
  const routingCombos = providerKey === "9router" ? ["FREE", "coder", "codex"] : null;
  const models = routingCombos ? routingCombos.filter((model) => rawModels.includes(model)) : rawModels;

  return models.map((model) => ({
    model,
    label: labelForModel(settings, model),
  }));
}

function labelForModel(settings, model) {
  if (settings.providerName.toLowerCase() === "9router") {
    return `9router/${model}`;
  }

  return model;
}

async function fetchProviderModels(settings) {
  const url = new URL("models", `${settings.baseURL.replace(/\/$/, "")}/`);
  const headers = {
    Accept: "application/json",
  };

  if (settings.apiKey) {
    headers.Authorization = `Bearer ${settings.apiKey}`;
  }

  const response = await fetch(url, { headers });
  const text = await response.text();

  if (!response.ok) {
    throw new Error(`Model list failed (${response.status}): ${text.slice(0, 240)}`);
  }

  const payload = JSON.parse(text);
  const rawModels = Array.isArray(payload.data)
    ? payload.data
    : Array.isArray(payload.models)
      ? payload.models
      : Array.isArray(payload)
        ? payload
        : [];

  return rawModels
    .map((item) => normalizeProviderModelId(settings, typeof item === "string" ? item : item.id || item.name))
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b));
}

function normalizeProviderModelId(settings, model) {
  if (!model) {
    return "";
  }

  if (settings.providerName.toLowerCase().includes("gemini") || settings.providerName.toLowerCase().includes("google")) {
    return model.replace(/^models\//, "");
  }

  return model;
}

start().catch((error) => {
  console.error(error);
  process.exit(1);
});
