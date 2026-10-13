import { createRequire } from "node:module";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { registerNumbersTools } from "./tools/numbers.js";
import { registerPagesTools } from "./tools/pages.js";
import { registerKeynoteTools } from "./tools/keynote.js";
import { INSTRUCTIONS } from "./instructions.js";

const require = createRequire(import.meta.url);
const { version } = require("../package.json") as { version: string };

const UNSUPPORTED_INSTRUCTIONS = "iwork-mcp requires macOS and the iWork applications. The automation tools are unavailable on this platform.";

export function createIWorkServer(options: { platform?: NodeJS.Platform; name?: string; version?: string } = {}): McpServer {
  const platform = options.platform ?? process.platform;
  const server = new McpServer(
    { name: options.name ?? "iwork-mcp", version: options.version ?? version },
    { instructions: platform === "darwin" ? INSTRUCTIONS : UNSUPPORTED_INSTRUCTIONS },
  );
  if (platform !== "darwin") {
    server.tool("iwork_status", "Check iwork-mcp availability on this platform", {}, async () => ({
      content: [{ type: "text", text: JSON.stringify({ available: false, platform, message: "iwork-mcp requires macOS and the iWork applications." }) }],
    }));
    return server;
  }
  registerNumbersTools(server);
  registerPagesTools(server);
  registerKeynoteTools(server);
  return server;
}
