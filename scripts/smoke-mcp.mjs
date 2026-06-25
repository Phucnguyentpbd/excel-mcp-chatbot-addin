import { config } from "../server/config.mjs";
import { ExcelMcpClient } from "../server/mcp-client.mjs";

const client = new ExcelMcpClient({
  command: config.mcpCommand,
  args: config.mcpArgs,
  cwd: config.mcpCwd,
});

const tools = await client.listTools();
console.log(
  JSON.stringify(
    {
      toolCount: tools.length,
      toolNames: tools.slice(0, 10).map((tool) => tool.name),
    },
    null,
    2,
  ),
);

await client.close();
