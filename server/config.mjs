import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, "..");
const bundledMcpRoot = path.join(projectRoot, "vendor", "excel-mcp-server");
const bundledMcpPython = path.join(bundledMcpRoot, ".venv", "Scripts", "python.exe");
const fallbackMcpCommand = process.platform === "win32" ? "py" : "python3";

export const config = {
  port: Number(process.env.PORT || 3100),
  host: process.env.HOST || "localhost",
  protocol: process.env.EXCEL_CHATBOT_PROTOCOL || "https",
  openaiApiKey: process.env.OPENAI_API_KEY || "",
  openaiModel: process.env.OPENAI_MODEL || "gpt-4.1-mini",
  mcpCommand: process.env.EXCEL_MCP_COMMAND || (fs.existsSync(bundledMcpPython) ? bundledMcpPython : fallbackMcpCommand),
  mcpArgs: process.env.EXCEL_MCP_ARGS
    ? JSON.parse(process.env.EXCEL_MCP_ARGS)
    : ["-m", "excel_mcp", "stdio"],
  mcpCwd: process.env.EXCEL_MCP_CWD || bundledMcpRoot,
  projectRoot,
  bundledMcpRoot,
  bundledMcpPython,
  publicDir: path.join(projectRoot, "public"),
  manifestPath: path.join(projectRoot, "manifest.xml"),
};

export function getBaseUrl() {
  return `${config.protocol}://${config.host}:${config.port}`;
}
