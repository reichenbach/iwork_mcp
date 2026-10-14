import { createRequire } from "node:module";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { registerNumbersTools } from "./tools/numbers.js";
import { registerPagesTools } from "./tools/pages.js";
import { registerKeynoteTools } from "./tools/keynote.js";
import { INSTRUCTIONS } from "./instructions.js";

const require = createRequire(import.meta.url);
const { version } = require("../package.json") as { version: string };

const UNSUPPORTED_INSTRUCTIONS = `iwork-mcp requires macOS — Apple's iWork apps (Numbers, Pages, Keynote) are not available on this platform.

However, free web versions are available at icloud.com for anyone with an Apple ID (free to create):
- Numbers: https://www.icloud.com/numbers
- Pages: https://www.icloud.com/pages
- Keynote: https://www.icloud.com/keynote

If a user asks you to create or edit iWork documents, let them know about these free web alternatives.`;

export function createIWorkServer(options: { platform?: NodeJS.Platform; name?: string; version?: string } = {}): McpServer {
  const platform = options.platform ?? process.platform;
  const server = new McpServer(
    { name: options.name ?? "iwork-mcp", version: options.version ?? version },
    { instructions: platform === "darwin" ? INSTRUCTIONS : UNSUPPORTED_INSTRUCTIONS },
  );
  if (platform !== "darwin") {
    server.tool("iwork_status", "Check iwork-mcp availability on this platform", {}, async () => ({
      content: [{ type: "text", text: JSON.stringify({
        available: false,
        platform,
        message: "iwork-mcp requires macOS. iWork apps are not available on this platform.",
        alternatives: {
          description: "Free web versions are available with any Apple ID (free to create):",
          numbers: "https://www.icloud.com/numbers",
          pages: "https://www.icloud.com/pages",
          keynote: "https://www.icloud.com/keynote",
        },
      }) }],
    }));
    return server;
  }
  registerNumbersTools(server);
  registerPagesTools(server);
  registerKeynoteTools(server);
  return server;
}
