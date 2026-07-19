import { readFileSync, writeFileSync, mkdirSync, existsSync } from "node:fs";
import { join } from "node:path";
import { homedir } from "node:os";

const configDir = join(homedir(), "Library", "Application Support", "Claude");
const configPath = join(configDir, "claude_desktop_config.json");

const iworkEntry = {
  command: "npx",
  args: ["-y", "iwork-mcp"],
};

export function install(): void {
  // Ensure the Claude config directory exists
  if (!existsSync(configDir)) {
    mkdirSync(configDir, { recursive: true });
  }

  // Read existing config or start fresh
  let config: Record<string, unknown> = {};
  if (existsSync(configPath)) {
    try {
      config = JSON.parse(readFileSync(configPath, "utf-8"));
    } catch {
      console.error("Could not parse existing config. Backing up and creating new one.");
      writeFileSync(configPath + ".backup", readFileSync(configPath));
    }
  }

  // Ensure mcpServers exists
  if (!config.mcpServers || typeof config.mcpServers !== "object") {
    config.mcpServers = {};
  }

  const servers = config.mcpServers as Record<string, unknown>;

  // Check if already installed and valid
  const existing = servers.iwork as Record<string, unknown> | undefined;
  if (existing && existing.command === iworkEntry.command &&
      JSON.stringify(existing.args) === JSON.stringify(iworkEntry.args)) {
    console.log("iwork-mcp is already installed in Claude Desktop.");
    console.log("Restart Claude Desktop (Cmd+Q, reopen) if it's not showing up.");
    return;
  }

  const wasPresent = !!existing;
  servers.iwork = iworkEntry;

  // Write config
  writeFileSync(configPath, JSON.stringify(config, null, 2) + "\n");

  console.log(wasPresent ? "\niwork-mcp config repaired!\n" : "\niwork-mcp installed successfully!\n");
  console.log("Make sure Numbers, Pages, and Keynote are installed (free from the App Store).");
  console.log("Restart Claude Desktop to activate (Cmd+Q, then reopen).");
  console.log("117 tools for Numbers, Pages, and Keynote — ready to go.");
}
