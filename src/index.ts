#!/usr/bin/env node

import { access } from "node:fs/promises";
import { homedir } from "node:os";
import { join } from "node:path";
import { stdin, stdout } from "node:process";

const args = process.argv.slice(2);

function trimPastedLineEndings(value: string): string { return value.replace(/[\r\n]+$/, ""); }
function isLoopbackHostname(hostname: string): boolean { return ["127.0.0.1", "localhost", "[::1]"].includes(hostname.toLowerCase()); }

function valueAfter(flag: string): string | undefined {
  const index = args.indexOf(flag);
  return index >= 0 ? args[index + 1] : undefined;
}

async function hiddenPrompt(label: string): Promise<string> {
  if (!stdin.isTTY || !stdin.setRawMode) throw new Error("A terminal is required for interactive passphrase setup");
  stdout.write(label);
  stdin.setRawMode(true);
  stdin.resume();
  return new Promise((resolve, reject) => {
    let value = "";
    const cleanup = (): void => {
      stdin.off("data", onData);
      stdin.setRawMode(false);
      stdin.pause();
      stdout.write("\n");
    };
    const onData = (chunk: Buffer): void => {
      const input = chunk.toString("utf8");
      for (const character of input) {
        if (character === "\u0003") {
          cleanup();
          reject(new Error("Passphrase setup cancelled"));
          return;
        }
        if (character === "\r" || character === "\n") {
          cleanup();
          resolve(trimPastedLineEndings(value));
          return;
        }
        if (character === "\u007f") value = value.slice(0, -1);
        else if (character !== "\u0000") value += character;
      }
    };
    stdin.on("data", onData);
  });
}

if (args[0] === "install") {
  const { install } = await import("./install.js");
  install();
} else if (args.includes("--setup-passphrase")) {
  const { savePassphrase } = await import("./auth.js");
  const configDir = process.env.IWORK_MCP_CONFIG_DIR ?? join(homedir(), ".config", "iwork-mcp");
  let passphrase = process.env.IWORK_MCP_PASSPHRASE;
  if (!passphrase) {
    passphrase = trimPastedLineEndings(await hiddenPrompt("New owner passphrase (12+ characters): "));
    const confirmation = trimPastedLineEndings(await hiddenPrompt("Confirm passphrase: "));
    if (passphrase !== confirmation) throw new Error("Passphrases do not match");
  } else passphrase = trimPastedLineEndings(passphrase);
  await savePassphrase(configDir, passphrase);
  console.error("Owner passphrase saved securely.");
} else if (args.includes("--http")) {
  const host = valueAfter("--host") ?? "127.0.0.1";
  const port = Number(valueAfter("--port") ?? "8787");
  if (!Number.isInteger(port) || port < 1 || port > 65535) throw new Error("--port must be between 1 and 65535");
  const publicUrl = process.env.IWORK_MCP_PUBLIC_URL ?? `http://${host}:${port}`;
  const parsedPublicUrl = new URL(publicUrl);
  if (parsedPublicUrl.pathname !== "/" || parsedPublicUrl.search || parsedPublicUrl.hash) throw new Error("IWORK_MCP_PUBLIC_URL must be an origin without a path, query, or fragment");
  if (parsedPublicUrl.protocol !== "https:" && !(parsedPublicUrl.protocol === "http:" && isLoopbackHostname(parsedPublicUrl.hostname))) throw new Error("IWORK_MCP_PUBLIC_URL must use HTTPS unless it is a loopback URL");
  const configDir = process.env.IWORK_MCP_CONFIG_DIR ?? join(homedir(), ".config", "iwork-mcp");
  const dataDir = process.env.IWORK_MCP_DATA_DIR ?? join(homedir(), ".local", "share", "iwork-mcp");
  if (!process.env.IWORK_MCP_PASSPHRASE) {
    try {
      await access(join(configDir, "passphrase.json"));
    } catch {
      throw new Error("No owner passphrase is configured. Run iwork-mcp --setup-passphrase first.");
    }
  }
  const { listenHttp } = await import("./http.js");
  const redirectHostAllowlist = process.env.IWORK_MCP_REDIRECT_HOST_ALLOWLIST?.split(",").map(value => value.trim()).filter(Boolean);
  await listenHttp({ host, port, publicUrl, configDir, dataDir, environmentPassphrase: process.env.IWORK_MCP_PASSPHRASE ? trimPastedLineEndings(process.env.IWORK_MCP_PASSPHRASE) : undefined, redirectHostAllowlist });
  console.error(`iwork-mcp HTTP server listening on ${host}:${port}; public URL ${publicUrl}`);
} else {
  const { StdioServerTransport } = await import("@modelcontextprotocol/sdk/server/stdio.js");
  const { createIWorkServer } = await import("./server.js");
  const server = createIWorkServer();
  await server.connect(new StdioServerTransport());
  console.error("iwork-mcp stdio server running");
}
