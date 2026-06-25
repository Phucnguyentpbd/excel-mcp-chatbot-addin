const subtitleEl = document.getElementById("subtitle");
const readyPillEl = document.getElementById("ready-pill");
const messagesEl = document.getElementById("messages");
const formEl = document.getElementById("chat-form");
const messageInputEl = document.getElementById("message-input");
const sendButtonEl = document.getElementById("send-button");
const modelButtonEl = document.getElementById("model-button");
const newSessionButtonEl = document.getElementById("new-session-button");
const sessionsButtonEl = document.getElementById("sessions-button");
const providerButtonEl = document.getElementById("provider-button");
const modelSelectEl = document.getElementById("model-select");
const imageInputEl = document.getElementById("image-input");
const imageStripEl = document.getElementById("image-strip");
const settingsDialogEl = document.getElementById("settings-dialog");
const settingsFormEl = document.getElementById("settings-form");
const sessionsDialogEl = document.getElementById("sessions-dialog");
const sessionsListEl = document.getElementById("sessions-list");
const providerSearchEl = document.getElementById("provider-search");
const providerListEl = document.getElementById("provider-list");
const providerNameEl = document.getElementById("provider-name");
const apiModeEl = document.getElementById("api-mode");
const baseUrlEl = document.getElementById("base-url");
const apiKeyEl = document.getElementById("api-key");
const modelNameEl = document.getElementById("model-name");

const PROVIDERS = [
  {
    id: "9router",
    section: "Popular",
    name: "9Router",
    icon: "9",
    description: "Local routing, combo models, round robin",
    badge: "Recommended",
    baseURL: "http://127.0.0.1:20128/v1",
    apiMode: "chat_completions",
    model: "FREE",
    routingCombos: ["FREE", "coder", "codex"],
  },
  {
    id: "openai",
    section: "Popular",
    name: "OpenAI",
    icon: "O",
    description: "GPT models for general Excel work",
    baseURL: "https://api.openai.com/v1",
    apiMode: "responses",
    model: "gpt-4.1-mini",
  },
  {
    id: "gemini",
    section: "Popular",
    name: "Google Gemini",
    icon: "G",
    description: "Gemini via OpenAI-compatible endpoint",
    baseURL: "https://generativelanguage.googleapis.com/v1beta/openai",
    apiMode: "chat_completions",
    model: "gemini-2.5-flash",
  },
  {
    id: "openrouter",
    section: "Popular",
    name: "OpenRouter",
    icon: "OR",
    description: "Many hosted models through one API",
    baseURL: "https://openrouter.ai/api/v1",
    apiMode: "chat_completions",
    model: "openai/gpt-4.1-mini",
  },
  {
    id: "opencode",
    section: "Popular",
    name: "OpenCode",
    icon: "OC",
    description: "Coding-focused provider or gateway",
    baseURL: "https://api.opencode.ai/v1",
    apiMode: "chat_completions",
    model: "opencode/coder",
  },
  {
    id: "glm",
    section: "Popular",
    name: "GLM / Z.ai",
    icon: "AI",
    description: "GLM models with OpenAI-compatible API",
    baseURL: "https://open.bigmodel.cn/api/paas/v4",
    apiMode: "chat_completions",
    model: "glm-4.5",
  },
  {
    id: "alibaba",
    section: "Popular",
    name: "Alibaba Qwen",
    icon: "Q",
    description: "DashScope OpenAI-compatible Qwen models",
    baseURL: "https://dashscope-intl.aliyuncs.com/compatible-mode/v1",
    apiMode: "chat_completions",
    model: "qwen-plus",
  },
  {
    id: "deepseek",
    section: "Popular",
    name: "DeepSeek",
    icon: "DS",
    description: "DeepSeek chat and reasoning models",
    baseURL: "https://api.deepseek.com/v1",
    apiMode: "chat_completions",
    model: "deepseek-chat",
  },
  {
    id: "anthropic",
    section: "Other",
    name: "Anthropic",
    icon: "A",
    description: "Use through an OpenAI-compatible proxy",
    baseURL: "https://api.anthropic.com/v1",
    apiMode: "chat_completions",
    model: "claude-sonnet-4-5",
  },
  {
    id: "groq",
    section: "Other",
    name: "Groq",
    icon: "GQ",
    description: "Fast inference for open models",
    baseURL: "https://api.groq.com/openai/v1",
    apiMode: "chat_completions",
    model: "llama-3.3-70b-versatile",
  },
  {
    id: "moonshot",
    section: "Other",
    name: "Moonshot / Kimi",
    icon: "K",
    description: "Kimi and Moonshot OpenAI-compatible API",
    baseURL: "https://api.moonshot.ai/v1",
    apiMode: "chat_completions",
    model: "kimi-k2-0905-preview",
  },
  {
    id: "copilot",
    section: "Other",
    name: "GitHub Copilot",
    icon: "GH",
    description: "Use through your Copilot-compatible local bridge",
    baseURL: "http://127.0.0.1:4141/v1",
    apiMode: "chat_completions",
    model: "gpt-4.1",
  },
  {
    id: "vercel",
    section: "Other",
    name: "Vercel AI Gateway",
    icon: "V",
    description: "Vercel gateway for multiple providers",
    baseURL: "https://ai-gateway.vercel.sh/v1",
    apiMode: "chat_completions",
    model: "openai/gpt-4.1-mini",
  },
  {
    id: "custom",
    section: "Other",
    name: "Custom",
    icon: "+",
    description: "Any OpenAI-compatible provider",
    baseURL: "https://your-provider.example/v1",
    apiMode: "chat_completions",
    model: "custom-model",
  },
];

const SESSION_STORAGE_KEY = "excel-mcp-chat-sessions-v1";
const ACTIVE_SESSION_KEY = "excel-mcp-active-session-v1";
const history = [];
const chatMessages = [];
let workbookContext = {};
let officeReady = false;
let providerSettings = null;
let attachedImages = [];
let sessions = [];
let activeSessionId = "";

bootstrap();

async function bootstrap() {
  bindEvents();
  loadSessions();
  renderActiveSession();
  await refreshBridgeConfig();

  if (window.Office) {
    Office.onReady(async (info) => {
      officeReady = true;
      appendMessage("system", `Excel connected: ${info.host || "Workbook"}`);
      await refreshWorkbookContext();
    });
  } else {
    appendMessage("system", "Office.js unavailable. Browser preview mode only.");
  }
}

function bindEvents() {
  newSessionButtonEl.addEventListener("click", () => {
    createNewSession();
    appendMessage("system", "New chat session started.");
  });
  sessionsButtonEl.addEventListener("click", openSessionsDialog);
  modelButtonEl.addEventListener("click", () => {
    modelSelectEl.focus();
    if (typeof modelSelectEl.showPicker === "function") {
      modelSelectEl.showPicker();
    }
  });
  providerButtonEl.addEventListener("click", openSettingsDialog);
  providerSearchEl.addEventListener("input", () => renderProviderList(providerSearchEl.value));
  imageInputEl.addEventListener("change", handleImageFiles);
  messageInputEl.addEventListener("paste", handleMessagePaste);
  modelSelectEl.addEventListener("change", async () => {
    await activateSelectedModel();
  });

  document.querySelectorAll("[data-preset]").forEach((button) => {
    button.addEventListener("click", () => applyPreset(button.dataset.preset));
  });

  settingsFormEl.addEventListener("submit", async (event) => {
    if (event.submitter?.value === "cancel") {
      return;
    }

    event.preventDefault();
    await saveProviderSettings();
    settingsDialogEl.close();
  });

  formEl.addEventListener("submit", sendMessage);
}

async function refreshBridgeConfig() {
  try {
    const response = await fetch("/api/config");
    const payload = await response.json();
    providerSettings = payload.provider;
    fillSettingsForm(providerSettings);
    await loadAllModelOptions();
    setStatus(payload.openaiConfigured ? "Ready" : "Need key", payload.openaiConfigured);
    subtitleEl.textContent = `${providerSettings.providerName} / ${providerSettings.model} / MCP tools: ${payload.mcpTools}`;
  } catch (error) {
    setStatus("Offline", false);
    subtitleEl.textContent = error.message;
  }
}

async function saveProviderSettings({ silent = false, skipModelReload = false } = {}) {
  const next = collectProviderFormSettings();
  const response = await fetch("/api/provider-settings", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(next),
  });
  providerSettings = await response.json();
  fillSettingsForm(providerSettings);
  if (!skipModelReload) {
    await loadAllModelOptions();
  }
  setStatus(providerSettings.configured ? "Ready" : "Need key", providerSettings.configured);
  subtitleEl.textContent = `${providerSettings.providerName} / ${providerSettings.model}`;
  if (!silent) {
    appendMessage("system", `Provider saved: ${providerSettings.providerName} -> ${providerSettings.model}`);
  }
}

function openSettingsDialog() {
  fillSettingsForm(providerSettings || {});
  providerSearchEl.value = "";
  renderProviderList();
  settingsDialogEl.showModal();
}

function fillSettingsForm(settings) {
  providerNameEl.value = settings.providerName || "OpenAI";
  apiModeEl.value = settings.apiMode || "responses";
  baseUrlEl.value = settings.baseURL || "https://api.openai.com/v1";
  apiKeyEl.value = settings.apiKey || "";
  modelNameEl.value = settings.model || "gpt-4.1-mini";
}

function applyPreset(preset) {
  const provider = PROVIDERS.find((item) => item.id === preset) || PROVIDERS.find((item) => item.id === "custom");
  applyProvider(provider);
}

async function loadAllModelOptions() {
  modelSelectEl.disabled = true;

  try {
    const response = await fetch("/api/model-options");
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || "Cannot load model options.");
    }

    renderModelOptions(payload.groups || [], payload.activeProvider, payload.activeModel);
    subtitleEl.textContent = `${payload.activeProvider} / ${displaySelectedModelLabel()}`;
  } catch (error) {
    renderFallbackModelOption();
    subtitleEl.textContent = error.message;
  } finally {
    modelSelectEl.disabled = false;
  }
}

function renderModelOptions(groups = [], activeProvider = "", activeModel = "") {
  modelSelectEl.innerHTML = "";

  for (const group of groups) {
    if (!group.models?.length) {
      continue;
    }

    const optgroup = document.createElement("optgroup");
    optgroup.label = group.providerName;
    for (const model of group.models) {
      const option = document.createElement("option");
      option.value = JSON.stringify({
        providerName: group.providerName,
        model: model.model,
      });
      option.textContent = model.label;
      option.selected = group.providerName === activeProvider && model.model === activeModel;
      optgroup.append(option);
    }
    modelSelectEl.append(optgroup);
  }

  if (!modelSelectEl.options.length) {
    renderFallbackModelOption();
  }
}

function renderFallbackModelOption() {
  const provider = providerNameEl.value || providerSettings?.providerName || "Provider";
  const model = modelNameEl.value || providerSettings?.model || "";
  modelSelectEl.innerHTML = "";
  const option = document.createElement("option");
  option.value = JSON.stringify({ providerName: provider, model });
  option.textContent = model ? `${provider}/${model}` : "No configured models";
  modelSelectEl.append(option);
}

async function activateSelectedModel() {
  const selected = parseSelectedModel();
  if (!selected.providerName || !selected.model) {
    return;
  }

  const response = await fetch("/api/active-model", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(selected),
  });
  const payload = await response.json();
  if (!response.ok) {
    appendMessage("system", `Model switch failed: ${payload.error || "unknown error"}`);
    return;
  }

  providerSettings = payload;
  fillSettingsForm(providerSettings);
  subtitleEl.textContent = `${providerSettings.providerName} / ${displaySelectedModelLabel()}`;
  setStatus(providerSettings.configured ? "Ready" : "Need key", providerSettings.configured);
}

function parseSelectedModel() {
  try {
    return JSON.parse(modelSelectEl.value || "{}");
  } catch {
    return { providerName: providerNameEl.value, model: modelSelectEl.value };
  }
}

function displaySelectedModelLabel() {
  return modelSelectEl.selectedOptions?.[0]?.textContent || providerSettings?.model || "";
}

function renderProviderList(query = "") {
  const normalizedQuery = query.trim().toLowerCase();
  const filtered = PROVIDERS.filter((provider) =>
    [provider.name, provider.description, provider.baseURL, provider.model]
      .join(" ")
      .toLowerCase()
      .includes(normalizedQuery),
  );
  providerListEl.innerHTML = "";

  for (const section of ["Popular", "Other"]) {
    const sectionProviders = filtered.filter((provider) => provider.section === section);
    if (!sectionProviders.length) {
      continue;
    }

    const title = document.createElement("h3");
    title.textContent = section;
    providerListEl.append(title);

    for (const provider of sectionProviders) {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "provider-row";
      button.innerHTML = `
        <span class="provider-icon">${escapeHtml(provider.icon)}</span>
        <span class="provider-copy">
          <strong>${escapeHtml(provider.name)}</strong>
          <small>${escapeHtml(provider.description)}</small>
        </span>
        ${provider.badge ? `<span class="provider-badge">${escapeHtml(provider.badge)}</span>` : ""}
      `;
      button.addEventListener("click", () => applyProvider(provider));
      providerListEl.append(button);
    }
  }
}

async function applyProvider(provider) {
  let saved = null;
  try {
    const response = await fetch(`/api/provider-settings?provider=${encodeURIComponent(provider.name)}`);
    saved = await response.json();
  } catch {
    saved = null;
  }

  const settings = saved?.configured || saved?.baseURL ? { ...provider, ...saved } : provider;
  providerNameEl.value = settings.providerName || settings.name;
  apiModeEl.value = settings.apiMode;
  baseUrlEl.value = settings.baseURL;
  apiKeyEl.value = settings.apiKey || "";
  modelNameEl.value = settings.model;
}

function setStatus(text, ready) {
  readyPillEl.textContent = text;
  readyPillEl.classList.toggle("ready", Boolean(ready));
  readyPillEl.classList.toggle("warn", !ready);
}

async function handleImageFiles(event) {
  const files = Array.from(event.target.files || []).slice(0, 4);
  await addPendingImages(files);
  imageInputEl.value = "";
}

async function handleMessagePaste(event) {
  const items = Array.from(event.clipboardData?.items || []);
  const files = items
    .filter((item) => item.kind === "file" && item.type.startsWith("image/"))
    .map((item) => item.getAsFile())
    .filter(Boolean);

  if (!files.length) {
    return;
  }

  event.preventDefault();
  await addPendingImages(files);
}

async function addPendingImages(files) {
  const room = Math.max(0, 4 - attachedImages.length);
  if (!room) {
    return;
  }

  const nextImages = await Promise.all(
    files.slice(0, room).map((file, index) => readImageFile(file, createScreenshotName(index, file.type))),
  );
  attachedImages = [...attachedImages, ...nextImages].slice(0, 4);
  renderImages();
}

function readImageFile(file, fallbackName = "pasted-image.png") {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve({ name: file.name || fallbackName, dataUrl: reader.result });
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

function createScreenshotName(index, mimeType = "image/png") {
  const extension = mimeType.includes("jpeg") ? "jpg" : mimeType.includes("webp") ? "webp" : "png";
  const stamp = new Date().toISOString().replace(/[-:T]/g, "").slice(0, 14);
  return `screenshot-${stamp}${index ? `-${index + 1}` : ""}.${extension}`;
}

function renderImages() {
  imageStripEl.hidden = attachedImages.length === 0;
  imageStripEl.innerHTML = "";
  attachedImages.forEach((image, index) => {
    const chip = document.createElement("div");
    chip.className = "image-chip";
    chip.innerHTML = `<img alt="" src="${image.dataUrl}" /><span>${escapeHtml(image.name)}</span>`;
    const removeButton = document.createElement("button");
    removeButton.type = "button";
    removeButton.textContent = "×";
    removeButton.addEventListener("click", () => {
      attachedImages.splice(index, 1);
      renderImages();
    });
    chip.append(removeButton);
    imageStripEl.append(chip);
  });
}

async function sendMessage(event) {
  event.preventDefault();
  const message = messageInputEl.value.trim();
  if (!message && !attachedImages.length) {
    return;
  }

  sendButtonEl.disabled = true;
  const outgoingImages = attachedImages.slice();
  appendMessage("user", message || "(image only)", [], outgoingImages);
  history.push({ role: "user", content: message || "(image only)" });
  messageInputEl.value = "";
  attachedImages = [];
  renderImages();

  try {
    await refreshWorkbookContext();
    const response = await fetch("/api/chat", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        history,
        message,
        workbookContext,
        images: outgoingImages,
        providerSettings: collectProviderSettings(),
      }),
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || "Chat request failed.");
    }

    appendMessage("assistant", payload.text || "(No text returned)", payload.toolTrace || []);
    history.push({ role: "assistant", content: payload.text || "" });
    subtitleEl.textContent = `${payload.providerName || providerSettings.providerName} / ${payload.model || providerSettings.model}`;
  } catch (error) {
    appendMessage("assistant", `Error: ${error.message}`);
  } finally {
    sendButtonEl.disabled = false;
  }
}

function collectProviderSettings() {
  const selected = parseSelectedModel();
  return {
    providerName: selected.providerName || providerNameEl.value,
    apiMode: apiModeEl.value,
    baseURL: baseUrlEl.value,
    apiKey: apiKeyEl.value,
    model: selected.model || modelNameEl.value,
  };
}

function collectProviderFormSettings() {
  return {
    providerName: providerNameEl.value,
    apiMode: apiModeEl.value,
    baseURL: baseUrlEl.value,
    apiKey: apiKeyEl.value,
    model: modelNameEl.value || modelSelectEl.value,
  };
}

async function refreshWorkbookContext() {
  workbookContext = officeReady ? await getWorkbookContext() : fallbackContext();
}

async function getWorkbookContext() {
  try {
    return await Excel.run(async (context) => {
      const workbook = context.workbook;
      const sheet = workbook.worksheets.getActiveWorksheet();
      const selectedRange = workbook.getSelectedRange();
      const activeCell = workbook.getActiveCell();

      workbook.load("name");
      sheet.load("name");
      selectedRange.load(["address", "values", "rowCount", "columnCount"]);
      activeCell.load("address");
      await context.sync();

      return {
        workbookName: workbook.name,
        workbookUrl: Office.context.document?.url || "",
        activeWorksheet: sheet.name,
        activeCellAddress: activeCell.address,
        selectedRangeAddress: selectedRange.address,
        selectionShape: {
          rows: selectedRange.rowCount,
          columns: selectedRange.columnCount,
        },
        selectedValuesPreview: truncateGrid(selectedRange.values || []),
      };
    });
  } catch (error) {
    appendMessage("system", `Excel context error: ${error.message}`);
    return fallbackContext(error.message);
  }
}

function truncateGrid(values) {
  return values.slice(0, 12).map((row) => row.slice(0, 8));
}

function fallbackContext(error = "") {
  return {
    workbookName: "",
    workbookUrl: "",
    activeWorksheet: "",
    activeCellAddress: "",
    selectedRangeAddress: "",
    selectedValuesPreview: [],
    error,
  };
}

function appendMessage(role, content, toolTrace = [], images = []) {
  const timestamp = new Date().toISOString();
  chatMessages.push({
    role,
    content,
    toolTrace,
    images,
    timestamp,
  });
  saveActiveSession();
  renderMessage({ role, content, toolTrace, images, timestamp });
}

function splitVisibleReasoning(content) {
  const thoughts = [];
  const finalText = String(content || "")
    .replace(/<(thought|thinking|reasoning)>([\s\S]*?)<\/\1>/gi, (_match, _tag, thought) => {
      const trimmed = thought.trim();
      if (trimmed) {
        thoughts.push(trimmed);
      }
      return "";
    })
    .trim();

  return {
    finalText,
    reasoning: thoughts.join("\n\n"),
  };
}

async function copyMessageContent(button, content, images = []) {
  const imageNames = images.map((image) => image.name).filter(Boolean);
  const visibleText = splitVisibleReasoning(content).finalText || String(content || "");
  const text = [visibleText, imageNames.length ? `Images: ${imageNames.join(", ")}` : ""]
    .filter(Boolean)
    .join("\n");

  try {
    await navigator.clipboard.writeText(text);
    button.textContent = "✓";
  } catch {
    const textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.setAttribute("readonly", "");
    textarea.style.position = "fixed";
    textarea.style.opacity = "0";
    document.body.append(textarea);
    textarea.select();
    document.execCommand("copy");
    textarea.remove();
    button.textContent = "✓";
  }

  window.setTimeout(() => {
    button.textContent = "⧉";
  }, 1200);
}

function editUserMessage(content, images = []) {
  messageInputEl.value = String(content || "").replace(/^\(image only\)$/, "");
  attachedImages = images.map((image) => ({ ...image }));
  renderImages();
  messageInputEl.focus();
  messageInputEl.setSelectionRange(messageInputEl.value.length, messageInputEl.value.length);
}

function renderImageGallery(images = []) {
  if (!images.length) {
    return null;
  }

  const gallery = document.createElement("div");
  gallery.className = "message-images";
  for (const image of images) {
    const figure = document.createElement("figure");
    const img = document.createElement("img");
    img.alt = image.name || "Attached image";
    img.src = image.dataUrl;
    const caption = document.createElement("figcaption");
    caption.textContent = image.name || "Attached image";
    figure.append(img, caption);
    gallery.append(figure);
  }
  return gallery;
}

function renderMessage({ role, content, toolTrace = [], images = [], timestamp }) {
  const article = document.createElement("article");
  article.className = `message ${role}`;

  const header = document.createElement("div");
  header.className = "message-header";
  const roleLabel = document.createElement("span");
  roleLabel.textContent = role === "assistant" ? "Assistant" : role === "system" ? "System" : "You";
  header.append(roleLabel);

  const body = document.createElement("div");
  body.className = "message-body";
  const parsedContent =
    role === "assistant"
      ? splitVisibleReasoning(content)
      : { finalText: String(content || ""), reasoning: "" };
  body.textContent = parsedContent.finalText || (parsedContent.reasoning ? "(No final answer)" : "");

  article.append(header, body);

  const gallery = renderImageGallery(images);
  if (gallery) {
    article.append(gallery);
  }

  if (parsedContent.reasoning) {
    const details = document.createElement("details");
    details.className = "reasoning-trace";
    details.open = true;
    const summary = document.createElement("summary");
    summary.textContent = "Reasoning";
    const pre = document.createElement("pre");
    pre.textContent = parsedContent.reasoning;
    details.append(summary, pre);
    article.insertBefore(details, body);
  }

  if (toolTrace.length) {
    const details = document.createElement("details");
    details.className = "tool-trace";
    const summary = document.createElement("summary");
    summary.textContent = `Tool trace (${toolTrace.length})`;
    const pre = document.createElement("pre");
    pre.textContent = JSON.stringify(toolTrace, null, 2);
    details.append(summary, pre);
    article.append(details);
  }

  const footer = document.createElement("div");
  footer.className = "message-actions";
  const time = document.createElement("span");
  time.className = "message-time";
  time.textContent = new Date(timestamp || Date.now()).toLocaleTimeString([], {
    hour: "numeric",
    minute: "2-digit",
  });
  footer.append(time);

  if (role === "assistant" || role === "user") {
    const copyButton = document.createElement("button");
    copyButton.type = "button";
    copyButton.className = "message-action-button";
    copyButton.textContent = "⧉";
    copyButton.title = "Copy message";
    copyButton.addEventListener("click", () => copyMessageContent(copyButton, content, images));
    footer.append(copyButton);
  }

  if (role === "user") {
    const editButton = document.createElement("button");
    editButton.type = "button";
    editButton.className = "message-action-button";
    editButton.textContent = "✎";
    editButton.title = "Edit and resend";
    editButton.addEventListener("click", () => editUserMessage(content, images));
    footer.append(editButton);
  }

  article.append(footer);
  messagesEl.append(article);
  article.scrollIntoView({ behavior: "smooth", block: "end" });
}

function loadSessions() {
  sessions = readSessions();
  activeSessionId = localStorage.getItem(ACTIVE_SESSION_KEY) || "";
  if (!sessions.some((session) => session.id === activeSessionId)) {
    activeSessionId = sessions[0]?.id || "";
  }
  if (!activeSessionId) {
    createNewSession({ silent: true });
  }
}

function readSessions() {
  try {
    const parsed = JSON.parse(localStorage.getItem(SESSION_STORAGE_KEY) || "[]");
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function writeSessions() {
  localStorage.setItem(SESSION_STORAGE_KEY, JSON.stringify(sessions.slice(0, 60)));
  localStorage.setItem(ACTIVE_SESSION_KEY, activeSessionId);
}

function createNewSession({ silent = false } = {}) {
  const now = new Date().toISOString();
  const session = {
    id: `session-${Date.now()}-${Math.random().toString(16).slice(2)}`,
    title: "New chat",
    createdAt: now,
    updatedAt: now,
    providerName: providerSettings?.providerName || providerNameEl.value || "",
    model: providerSettings?.model || modelSelectEl.selectedOptions?.[0]?.textContent || "",
    messages: [],
  };
  sessions.unshift(session);
  activeSessionId = session.id;
  history.length = 0;
  chatMessages.length = 0;
  messagesEl.innerHTML = "";
  writeSessions();
  if (!silent) {
    renderSessionsList();
  }
}

function renderActiveSession() {
  const session = getActiveSession();
  history.length = 0;
  chatMessages.length = 0;
  messagesEl.innerHTML = "";
  for (const message of session?.messages || []) {
    chatMessages.push(message);
    if (message.role === "user" || message.role === "assistant") {
      history.push({ role: message.role, content: message.content });
    }
    renderMessage(message);
  }
  if (!chatMessages.length) {
    appendMessage("system", "OpenCode connected. MCP `excel` is registered.");
  }
}

function saveActiveSession() {
  const session = getActiveSession();
  if (!session) {
    return;
  }
  session.messages = chatMessages.slice(-200);
  session.updatedAt = new Date().toISOString();
  session.providerName = providerSettings?.providerName || providerNameEl.value || session.providerName;
  session.model = modelSelectEl.selectedOptions?.[0]?.textContent || providerSettings?.model || session.model;
  const firstUserMessage = chatMessages.find((message) => message.role === "user")?.content;
  if (firstUserMessage) {
    session.title = firstUserMessage.slice(0, 48);
  }
  writeSessions();
}

function getActiveSession() {
  return sessions.find((session) => session.id === activeSessionId) || null;
}

function openSessionsDialog() {
  renderSessionsList();
  sessionsDialogEl.showModal();
}

function renderSessionsList() {
  sessionsListEl.innerHTML = "";
  if (!sessions.length) {
    sessionsListEl.textContent = "No sessions yet.";
    return;
  }

  for (const session of sessions) {
    const row = document.createElement("article");
    row.className = `session-row ${session.id === activeSessionId ? "active" : ""}`;
    const title = document.createElement("div");
    title.className = "session-title";
    title.innerHTML = `<strong>${escapeHtml(session.title || "Untitled")}</strong><small>${escapeHtml(session.providerName || "")} ${escapeHtml(session.model || "")} · ${new Date(session.updatedAt).toLocaleString()}</small>`;
    const openButton = document.createElement("button");
    openButton.type = "button";
    openButton.textContent = "Open";
    openButton.addEventListener("click", () => {
      activeSessionId = session.id;
      writeSessions();
      renderActiveSession();
      sessionsDialogEl.close();
    });
    const deleteButton = document.createElement("button");
    deleteButton.type = "button";
    deleteButton.textContent = "Delete";
    deleteButton.addEventListener("click", () => {
      sessions = sessions.filter((item) => item.id !== session.id);
      if (activeSessionId === session.id) {
        activeSessionId = sessions[0]?.id || "";
        if (!activeSessionId) {
          createNewSession({ silent: true });
        } else {
          renderActiveSession();
        }
      }
      writeSessions();
      renderSessionsList();
    });
    row.append(title, openButton, deleteButton);
    sessionsListEl.append(row);
  }
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    "\"": "&quot;",
    "'": "&#039;",
  })[char]);
}
