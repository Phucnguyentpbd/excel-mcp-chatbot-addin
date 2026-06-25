import { Client, StdioClientTransport } from "@modelcontextprotocol/client";

export class ExcelMcpClient {
  constructor({ command, args, cwd }) {
    this.command = command;
    this.args = args;
    this.cwd = cwd;
    this.client = null;
    this.transport = null;
    this.toolsCache = null;
  }

  async connect() {
    if (this.client && this.transport) {
      return;
    }

    this.client = new Client({
      name: "excel-chatbot-addin",
      version: "0.1.0",
    });

    this.transport = new StdioClientTransport({
      command: this.command,
      args: this.args,
      cwd: this.cwd,
    });

    await this.client.connect(this.transport);
  }

  async ensureTools() {
    await this.connect();
    if (!this.toolsCache) {
      const result = await this.client.listTools();
      this.toolsCache = result.tools;
    }
    return this.toolsCache;
  }

  async listTools() {
    return await this.ensureTools();
  }

  async callTool(name, args = {}) {
    await this.ensureTools();
    const result = await this.client.callTool({
      name,
      arguments: args,
    });

    return {
      name,
      args,
      isError: Boolean(result.isError),
      structuredContent: result.structuredContent || null,
      text: extractTextContent(result.content || []),
      content: result.content || [],
    };
  }

  async close() {
    if (this.transport) {
      await this.transport.close();
    }
    this.transport = null;
    this.client = null;
    this.toolsCache = null;
  }
}

function extractTextContent(content) {
  const textBlocks = content
    .filter((item) => item?.type === "text" && typeof item.text === "string")
    .map((item) => item.text.trim())
    .filter(Boolean);

  if (textBlocks.length) {
    return textBlocks.join("\n\n");
  }

  return JSON.stringify(content, null, 2);
}
