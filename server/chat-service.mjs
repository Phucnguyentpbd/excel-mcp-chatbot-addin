import OpenAI from "openai";

const SYSTEM_PROMPT = `You are an Excel copilot running inside a task pane add-in.

Your main job is to help the user work with the workbook that is currently open in Excel.

Important rules:
- Prefer live-workbook operations over offline file edits whenever the user is clearly talking about the workbook currently open in Excel.
- For live-workbook operations, prefer the MCP tool "run_python" and use Python COM automation through win32com.client against the running Excel instance.
- The current workbook context will be included in the request. When you use "run_python", always assume the backend will merge that workbook context into the tool's "data" argument automatically.
- Use the workbook name from data["workbookName"] to find the open workbook in Excel COM.
- Use data["activeWorksheet"], data["selectedRangeAddress"], data["activeCellAddress"], and data["selectedValuesPreview"] when they help.
- If the task is better suited to the built-in file-based Excel MCP tools and the user supplied a workbook path, you may use those file-based tools.
- Be explicit about what you changed.
- Keep answers concise and action-oriented.

Useful Python COM pattern for run_python:
import win32com.client as win32
excel = win32.GetActiveObject("Excel.Application")
wb = None
for candidate in excel.Workbooks:
    if candidate.Name.lower() == data["workbookName"].lower():
        wb = candidate
        break
if wb is None:
    raise RuntimeError(f'Workbook {data["workbookName"]} is not open in Excel.')
ws = wb.Worksheets(data["activeWorksheet"])

When returning structured data from run_python, assign it to the variable "output".
`;

export class ChatService {
  constructor({ mcpClient, providerStore }) {
    this.mcpClient = mcpClient;
    this.providerStore = providerStore;
  }

  isConfigured() {
    const settings = this.providerStore.read();
    return Boolean(settings.apiKey && settings.model);
  }

  async getToolDefinitions() {
    const tools = await this.mcpClient.listTools();
    return tools.map((tool) => ({
      type: "function",
      name: tool.name,
      description: tool.description || `Call the MCP tool ${tool.name}`,
      parameters: tool.inputSchema || {
        type: "object",
        properties: {},
        additionalProperties: true,
      },
    }));
  }

  async runConversation({ history = [], message, workbookContext = {}, images = [], providerSettings = null }) {
    const settings = providerSettings
      ? this.providerStore.write(mergeProviderSettings(this.providerStore.read(), providerSettings))
      : this.providerStore.read();

    if (!settings.apiKey) {
      throw new Error("Provider API key is not configured.");
    }

    const openai = new OpenAI({
      apiKey: settings.apiKey,
      baseURL: settings.baseURL,
    });

    if (settings.apiMode === "chat_completions") {
      return this.runChatCompletions({
        openai,
        settings,
        history,
        message,
        workbookContext,
        images,
      });
    }

    return this.runResponses({
      openai,
      settings,
      history,
      message,
      workbookContext,
      images,
    });
  }

  async runResponses({ openai, settings, history = [], message, workbookContext = {}, images = [] }) {
    const tools = await this.getToolDefinitions();
    const input = [
      {
        role: "system",
        content: [{ type: "input_text", text: SYSTEM_PROMPT }],
      },
      ...history.map(toResponseMessage),
      {
        role: "user",
        content: buildResponsesUserContent(message, workbookContext, images),
      },
    ];

    let response = await openai.responses.create({
      model: settings.model,
      input,
      tools,
      tool_choice: "auto",
    });

    const toolTrace = [];

    for (let iteration = 0; iteration < 8; iteration += 1) {
      const toolCalls = (response.output || []).filter((item) => item.type === "function_call");
      if (!toolCalls.length) {
        break;
      }

      const toolOutputs = [];

      for (const call of toolCalls) {
        const parsedArgs = safeParseJson(call.arguments);
        const mergedArgs = injectWorkbookContext(call.name, parsedArgs, workbookContext);
        const toolResult = await this.mcpClient.callTool(call.name, mergedArgs);

        toolTrace.push({
          toolName: call.name,
          arguments: mergedArgs,
          isError: toolResult.isError,
          text: toolResult.text,
        });

        toolOutputs.push({
          type: "function_call_output",
          call_id: call.call_id,
          output: JSON.stringify(toolResult),
        });
      }

      response = await openai.responses.create({
        model: settings.model,
        previous_response_id: response.id,
        input: toolOutputs,
        tools,
        tool_choice: "auto",
      });
    }

    return {
      text: response.output_text || extractResponseText(response.output || []),
      toolTrace,
      model: settings.model,
      providerName: settings.providerName,
    };
  }

  async runChatCompletions({ openai, settings, history = [], message, workbookContext = {}, images = [] }) {
    const tools = await this.getToolDefinitions();
    const chatTools = tools.map((tool) => ({
      type: "function",
      function: {
        name: tool.name,
        description: tool.description,
        parameters: tool.parameters,
      },
    }));
    const messages = [
      { role: "system", content: SYSTEM_PROMPT },
      ...history.map((item) => ({
        role: item.role === "assistant" ? "assistant" : "user",
        content: item.content || "",
      })),
      {
        role: "user",
        content: buildChatUserContent(message, workbookContext, images),
      },
    ];
    const toolTrace = [];

    for (let iteration = 0; iteration < 8; iteration += 1) {
      const response = await openai.chat.completions.create({
        model: settings.model,
        messages,
        tools: chatTools,
        tool_choice: "auto",
      });
      const assistantMessage = response.choices?.[0]?.message || {};
      messages.push(assistantMessage);

      const toolCalls = assistantMessage.tool_calls || [];
      if (!toolCalls.length) {
        return {
          text: assistantMessage.content || "",
          toolTrace,
          model: settings.model,
          providerName: settings.providerName,
        };
      }

      for (const call of toolCalls) {
        const parsedArgs = safeParseJson(call.function?.arguments);
        const mergedArgs = injectWorkbookContext(call.function?.name, parsedArgs, workbookContext);
        const toolResult = await this.mcpClient.callTool(call.function.name, mergedArgs);
        toolTrace.push({
          toolName: call.function.name,
          arguments: mergedArgs,
          isError: toolResult.isError,
          text: toolResult.text,
        });
        messages.push({
          role: "tool",
          tool_call_id: call.id,
          content: JSON.stringify(toolResult),
        });
      }
    }

    return {
      text: "Tool loop stopped after reaching the safety limit.",
      toolTrace,
      model: settings.model,
      providerName: settings.providerName,
    };
  }
}

function mergeProviderSettings(current, incoming) {
  return {
    ...incoming,
    baseURL: normalizeChatBaseURL(incoming.providerName, incoming.baseURL),
    model: normalizeChatModel(incoming.providerName, incoming.model),
    apiKey: incoming.apiKey && !isMaskedSecret(incoming.apiKey) ? incoming.apiKey : current.apiKey,
  };
}

function isMaskedSecret(value) {
  return String(value || "").includes("*") || String(value || "").includes("...");
}

function normalizeChatBaseURL(providerName, baseURL) {
  const clean = String(baseURL || "").trim().replace(/\/$/, "");
  const name = String(providerName || "").toLowerCase();

  if (name.includes("gemini") || name.includes("google")) {
    if (/^https:\/\/generativelanguage\.googleapis\.com(\/)?$/i.test(clean)) {
      return "https://generativelanguage.googleapis.com/v1beta/openai";
    }
    if (/^https:\/\/generativelanguage\.googleapis\.com\/v1beta$/i.test(clean)) {
      return "https://generativelanguage.googleapis.com/v1beta/openai";
    }
  }

  return clean;
}

function normalizeChatModel(providerName, model) {
  const name = String(providerName || "").toLowerCase();
  const clean = String(model || "").trim();
  if (name.includes("gemini") || name.includes("google")) {
    return clean.replace(/^models\//, "");
  }

  return clean;
}

function toResponseMessage(item) {
  return {
    role: item.role === "assistant" ? "assistant" : "user",
    content: [{ type: "input_text", text: item.content || "" }],
  };
}

function buildUserMessage(message, workbookContext) {
  return [
    "Workbook context:",
    JSON.stringify(workbookContext || {}, null, 2),
    "",
    "User request:",
    message,
  ].join("\n");
}

function buildResponsesUserContent(message, workbookContext, images) {
  return [
    {
      type: "input_text",
      text: buildUserMessage(message, workbookContext),
    },
    ...normalizeImages(images).map((image) => ({
      type: "input_image",
      image_url: image.dataUrl,
    })),
  ];
}

function buildChatUserContent(message, workbookContext, images) {
  const normalizedImages = normalizeImages(images);
  if (!normalizedImages.length) {
    return buildUserMessage(message, workbookContext);
  }

  return [
    {
      type: "text",
      text: buildUserMessage(message, workbookContext),
    },
    ...normalizedImages.map((image) => ({
      type: "image_url",
      image_url: {
        url: image.dataUrl,
      },
    })),
  ];
}

function normalizeImages(images) {
  return (images || [])
    .filter((image) => image?.dataUrl?.startsWith("data:image/"))
    .slice(0, 4);
}

function safeParseJson(value) {
  if (!value) {
    return {};
  }

  try {
    return JSON.parse(value);
  } catch {
    return {};
  }
}

function injectWorkbookContext(toolName, args, workbookContext) {
  if (toolName !== "run_python") {
    return args;
  }

  return {
    ...args,
    data: {
      ...(workbookContext || {}),
      ...(args.data || {}),
    },
  };
}

function extractResponseText(outputItems) {
  const parts = [];
  for (const item of outputItems) {
    if (item.type === "message") {
      for (const content of item.content || []) {
        if (content.type === "output_text" && content.text) {
          parts.push(content.text);
        }
      }
    }
  }
  return parts.join("\n\n").trim();
}
