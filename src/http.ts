import { randomBytes, randomUUID } from "node:crypto";
import { createServer, type IncomingMessage, type Server, type ServerResponse } from "node:http";
import { join } from "node:path";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { OAuthStore, verifyPassphrase } from "./auth.js";
import { createIWorkServer } from "./server.js";

interface PendingAuthorization { clientId: string; redirectUri: string; state?: string; codeChallenge: string; expiresAt: number }
export interface HttpServerOptions { host: string; port: number; publicUrl: string; configDir: string; dataDir: string; environmentPassphrase?: string; platform?: NodeJS.Platform }

function json(res: ServerResponse, status: number, value: unknown, headers: Record<string, string> = {}): void {
  res.writeHead(status, { "content-type": "application/json", "cache-control": "no-store", ...headers }); res.end(JSON.stringify(value));
}
function html(res: ServerResponse, status: number, value: string): void {
  res.writeHead(status, { "content-type": "text/html; charset=utf-8", "cache-control": "no-store", "x-frame-options": "DENY", "content-security-policy": "default-src 'none'; style-src 'unsafe-inline'; form-action 'self'; base-uri 'none'" }); res.end(value);
}
async function body(req: IncomingMessage): Promise<string> {
  const chunks: Buffer[] = []; let size = 0;
  for await (const chunk of req) { const bytes = Buffer.from(chunk); size += bytes.length; if (size > 65_536) throw new Error("request_too_large"); chunks.push(bytes); }
  return Buffer.concat(chunks).toString("utf8");
}
function escaped(value: string): string { return value.replace(/[&<>"']/g, character => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[character]!); }
function bearer(req: IncomingMessage): string | undefined { return /^Bearer ([A-Za-z0-9_-]+)$/.exec(req.headers.authorization ?? "")?.[1]; }
function isSafeRedirect(value: string): boolean {
  try { const url = new URL(value); return url.protocol === "https:" || (url.protocol === "http:" && ["127.0.0.1", "localhost", "[::1]"].includes(url.hostname)); }
  catch { return false; }
}

export function createHttpServer(options: HttpServerOptions): Server {
  const issuer = new URL(options.publicUrl).toString().replace(/\/$/, "");
  const resource = `${issuer}/mcp`;
  const store = new OAuthStore(join(options.dataDir, "oauth-state.json"));
  const pending = new Map<string, PendingAuthorization>();
  const attempts = new Map<string, { count: number; resetAt: number }>();
  const transports = new Map<string, { transport: StreamableHTTPServerTransport; server: McpServer }>();

  return createServer(async (req, res) => {
    try {
      const url = new URL(req.url ?? "/", issuer);
      if (req.method === "GET" && url.pathname === "/.well-known/oauth-protected-resource") return json(res, 200, { resource, authorization_servers: [issuer], bearer_methods_supported: ["header"], resource_name: "iwork-mcp" });
      if (req.method === "GET" && url.pathname === "/.well-known/oauth-authorization-server") return json(res, 200, { issuer, authorization_endpoint: `${issuer}/authorize`, token_endpoint: `${issuer}/token`, registration_endpoint: `${issuer}/register`, response_types_supported: ["code"], grant_types_supported: ["authorization_code", "refresh_token"], token_endpoint_auth_methods_supported: ["none"], code_challenge_methods_supported: ["S256"], scopes_supported: ["mcp"] });
      if (req.method === "POST" && url.pathname === "/register") {
        const input = JSON.parse(await body(req)) as Record<string, unknown>;
        if (!Array.isArray(input.redirect_uris) || input.redirect_uris.length === 0 || input.redirect_uris.some(item => typeof item !== "string" || !isSafeRedirect(item))) return json(res, 400, { error: "invalid_redirect_uri" });
        if (input.token_endpoint_auth_method !== undefined && input.token_endpoint_auth_method !== "none") return json(res, 400, { error: "invalid_client_metadata" });
        return json(res, 201, await store.register({ redirect_uris: input.redirect_uris as string[], client_name: typeof input.client_name === "string" ? input.client_name : undefined }));
      }
      if (req.method === "GET" && url.pathname === "/authorize") {
        const clientId = url.searchParams.get("client_id") ?? ""; const redirectUri = url.searchParams.get("redirect_uri") ?? ""; const challenge = url.searchParams.get("code_challenge") ?? ""; const client = await store.client(clientId);
        if (!client || url.searchParams.get("response_type") !== "code" || url.searchParams.get("code_challenge_method") !== "S256" || challenge.length < 43 || !client.redirect_uris.includes(redirectUri)) return json(res, 400, { error: "invalid_request" });
        const requestId = randomBytes(24).toString("base64url"); pending.set(requestId, { clientId, redirectUri, state: url.searchParams.get("state") ?? undefined, codeChallenge: challenge, expiresAt: Date.now() + 300_000 });
        return html(res, 200, `<!doctype html><html><head><meta name="viewport" content="width=device-width"><title>Authorize iwork-mcp</title><style>body{font:16px system-ui;max-width:32rem;margin:10vh auto;padding:1rem}input,button{box-sizing:border-box;font:inherit;padding:.7rem;width:100%;margin:.4rem 0}</style></head><body><h1>Authorize iwork-mcp</h1><p>Enter the owner passphrase to allow <strong>${escaped(client.client_name ?? "this client")}</strong> to use the server.</p><form method="post" action="/authorize"><input type="hidden" name="request_id" value="${requestId}"><label>Passphrase<input type="password" name="passphrase" required autofocus></label><button type="submit">Allow access</button></form></body></html>`);
      }
      if (req.method === "POST" && url.pathname === "/authorize") {
        const key = req.socket.remoteAddress ?? "unknown"; const rate = attempts.get(key);
        if (rate && rate.resetAt > Date.now() && rate.count >= 5) return json(res, 429, { error: "too_many_attempts" }, { "retry-after": String(Math.ceil((rate.resetAt - Date.now()) / 1000)) });
        const form = new URLSearchParams(await body(req)); const requestId = form.get("request_id") ?? ""; const authorization = pending.get(requestId);
        if (!authorization || authorization.expiresAt <= Date.now()) return json(res, 400, { error: "invalid_request" });
        if (!await verifyPassphrase(options.configDir, form.get("passphrase") ?? "", options.environmentPassphrase)) {
          const current = rate && rate.resetAt > Date.now() ? rate : { count: 0, resetAt: Date.now() + 600_000 }; current.count += 1; attempts.set(key, current);
          return html(res, 403, "<!doctype html><html><body><h1>Access denied</h1><p>The passphrase was not accepted.</p></body></html>");
        }
        attempts.delete(key); pending.delete(requestId); const code = await store.issueCode(authorization.clientId, authorization.redirectUri, authorization.codeChallenge); const redirect = new URL(authorization.redirectUri); redirect.searchParams.set("code", code); if (authorization.state !== undefined) redirect.searchParams.set("state", authorization.state);
        res.writeHead(302, { location: redirect.toString(), "cache-control": "no-store" }); return res.end();
      }
      if (req.method === "POST" && url.pathname === "/token") {
        const form = new URLSearchParams(await body(req)); const grant = form.get("grant_type"); const clientId = form.get("client_id") ?? "";
        if (!await store.client(clientId)) return json(res, 400, { error: "invalid_client" });
        const tokens = grant === "authorization_code" ? await store.exchangeCode(form.get("code") ?? "", clientId, form.get("redirect_uri") ?? "", form.get("code_verifier") ?? "") : grant === "refresh_token" ? await store.refresh(form.get("refresh_token") ?? "", clientId) : undefined;
        if (grant !== "authorization_code" && grant !== "refresh_token") return json(res, 400, { error: "unsupported_grant_type" });
        if (!tokens) return json(res, 400, { error: "invalid_grant" });
        return json(res, 200, { ...tokens, token_type: "Bearer", expires_in: 900, scope: "mcp" });
      }
      if (url.pathname === "/mcp") {
        const token = bearer(req);
        if (!token || !await store.validAccessToken(token)) return json(res, 401, { error: "invalid_token" }, { "www-authenticate": `Bearer resource_metadata="${issuer}/.well-known/oauth-protected-resource", scope="mcp"` });
        const sessionId = req.headers["mcp-session-id"] as string | undefined; let current = sessionId ? transports.get(sessionId) : undefined;
        if (!current) {
          if (sessionId) return json(res, 404, { error: "unknown_session" });
          let mcpServer: McpServer;
          let transport: StreamableHTTPServerTransport;
          transport = new StreamableHTTPServerTransport({ sessionIdGenerator: randomUUID, onsessioninitialized: (id): void => { transports.set(id, { transport, server: mcpServer }); } });
          mcpServer = createIWorkServer({ platform: options.platform }); transport.onclose = () => { if (transport.sessionId) transports.delete(transport.sessionId); }; await mcpServer.connect(transport); current = { transport, server: mcpServer };
        }
        return current.transport.handleRequest(req, res);
      }
      return json(res, 404, { error: "not_found" });
    } catch (error) { return json(res, 400, { error: error instanceof Error ? error.message : "bad_request" }); }
  });
}

export async function listenHttp(options: HttpServerOptions): Promise<Server> {
  const server = createHttpServer(options);
  await new Promise<void>((resolve, reject) => { server.once("error", reject); server.listen(options.port, options.host, resolve); });
  return server;
}
