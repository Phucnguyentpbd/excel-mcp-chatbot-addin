import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, "..");

export const config = {
  port: Number(process.env.PORT || 3100),
  host: process.env.HOST || "localhost",
  protocol: process.env.EXCEL_CHATBOT_PROTOCOL || "https",
  openaiApiKey: process.env.OPENAI_API_KEY || "",
  openaiModel: process.env.OPENAI_MODEL || "gpt-4.1-mini",
  mcpCommand:
    process.env.EXCEL_MCP_COMMAND ||
    "C:\\Users\\PHUC\\Documents\\Codex\\2026-06-24\\se\\work\\excel-mcp-server\\.venv\\Scripts\\python.exe",
  mcpArgs: process.env.EXCEL_MCP_ARGS
    ? JSON.parse(process.env.EXCEL_MCP_ARGS)
    : ["-m", "excel_mcp", "stdio"],
  mcpCwd:
    process.env.EXCEL_MCP_CWD ||
    "C:\\Users\\PHUC\\Documents\\Codex\\2026-06-24\\se\\work\\excel-mcp-server",
  projectRoot,
  publicDir: path.join(projectRoot, "public"),
  manifestPath: path.join(projectRoot, "manifest.xml"),
};

export function getBaseUrl() {
  return `${config.protocol}://${config.host}:${config.port}`;
}
