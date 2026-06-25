import fs from "node:fs";
import path from "node:path";

const DEFAULT_SETTINGS = {
  providerName: "OpenAI",
  apiKey: "",
  baseURL: "https://api.openai.com/v1",
  model: "gpt-4.1-mini",
  apiMode: "responses",
};

export class ProviderStore {
  constructor({ projectRoot, env }) {
    this.settingsPath = path.join(projectRoot, ".provider-settings.json");
    this.env = env;
  }

  read() {
    const saved = readJson(this.settingsPath);
    const activeProvider = saved.activeProvider || saved.providerName || DEFAULT_SETTINGS.providerName;
    const profiles = normalizeProfiles(saved.profiles || {});
    const activeProfile = profiles[activeProvider] || (saved.providerName ? saved : {});

    return normalizeSettings({
      ...DEFAULT_SETTINGS,
      apiKey: this.env.OPENAI_API_KEY || "",
      model: this.env.OPENAI_MODEL || DEFAULT_SETTINGS.model,
      ...activeProfile,
      providerName: activeProvider,
    });
  }

  write(settings) {
    const normalized = normalizeSettings({
      ...this.read(),
      ...settings,
    });
    const saved = readJson(this.settingsPath);
    const profiles = normalizeProfiles(saved.profiles || {});
    profiles[normalized.providerName] = normalized;

    fs.writeFileSync(
      this.settingsPath,
      JSON.stringify(
        {
          activeProvider: normalized.providerName,
          profiles,
        },
        null,
        2,
      ),
      "utf8",
    );
    return normalized;
  }

  publicSettings() {
    const settings = this.read();
    return publicSettings(settings);
  }

  publicProfile(providerName) {
    const settings = this.privateProfile(providerName);
    if (!settings) {
      return null;
    }

    return publicSettings(settings);
  }

  privateProfile(providerName) {
    const profiles = this.privateProfiles();
    return profiles[providerName] || null;
  }

  privateProfiles() {
    const saved = readJson(this.settingsPath);
    return normalizeProfiles(saved.profiles || {});
  }

  publicProfiles() {
    return Object.fromEntries(
      Object.entries(this.privateProfiles()).map(([providerName, settings]) => [
        providerName,
        publicSettings(settings),
      ]),
    );
  }

  activateModel({ providerName, model }) {
    const profile = this.privateProfile(providerName);
    if (!profile) {
      throw new Error(`Provider profile is not configured: ${providerName}`);
    }

    return this.write({
      ...profile,
      model,
    });
  }
}

function publicSettings(settings) {
  const configured = Boolean(settings.apiKey && settings.model && !isMaskedSecret(settings.apiKey));
  return {
    ...settings,
    apiKey: settings.apiKey ? maskSecret(settings.apiKey) : "",
    configured,
  };
}

function normalizeProfiles(profiles) {
  const normalized = {};
  for (const [providerName, settings] of Object.entries(profiles || {})) {
    normalized[providerName] = normalizeSettings({
      ...settings,
      providerName,
    });
  }
  return normalized;
}

function readJson(filePath) {
  if (!fs.existsSync(filePath)) {
    return {};
  }

  try {
    return JSON.parse(fs.readFileSync(filePath, "utf8").replace(/^\uFEFF/, ""));
  } catch {
    return {};
  }
}

function normalizeSettings(settings) {
  const providerName = String(settings.providerName || DEFAULT_SETTINGS.providerName).trim();
  const apiKey = String(settings.apiKey || "").trim();
  return {
    providerName,
    apiKey: isMaskedSecret(apiKey) ? "" : apiKey,
    baseURL: normalizeBaseURL(providerName, String(settings.baseURL || DEFAULT_SETTINGS.baseURL).trim()),
    model: normalizeModel(providerName, String(settings.model || DEFAULT_SETTINGS.model).trim()),
    apiMode: settings.apiMode === "chat_completions" ? "chat_completions" : "responses",
  };
}

function normalizeBaseURL(providerName, baseURL) {
  const clean = baseURL.replace(/\/$/, "");
  const normalizedName = providerName.toLowerCase();

  if (normalizedName.includes("gemini") || normalizedName.includes("google")) {
    if (/^https:\/\/generativelanguage\.googleapis\.com(\/)?$/i.test(clean)) {
      return "https://generativelanguage.googleapis.com/v1beta/openai";
    }
    if (/^https:\/\/generativelanguage\.googleapis\.com\/v1beta$/i.test(clean)) {
      return "https://generativelanguage.googleapis.com/v1beta/openai";
    }
  }

  if (normalizedName.includes("alibaba") || normalizedName.includes("qwen")) {
    if (/dashscope.*aliyuncs\.com\/?$/i.test(clean) && !clean.includes("compatible-mode")) {
      return "https://dashscope-intl.aliyuncs.com/compatible-mode/v1";
    }
  }

  return clean;
}

function normalizeModel(providerName, model) {
  const normalizedName = providerName.toLowerCase();
  if (normalizedName.includes("gemini") || normalizedName.includes("google")) {
    return model.replace(/^models\//, "");
  }

  return model;
}

function maskSecret(value) {
  if (value.length <= 8) {
    return "********";
  }

  return `${value.slice(0, 4)}...${value.slice(-4)}`;
}

function isMaskedSecret(value) {
  const text = String(value || "");
  return text.includes("*") || text.includes("...");
}
