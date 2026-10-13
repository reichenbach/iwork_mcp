import { randomBytes, randomUUID } from "node:crypto";
import { createServer, type IncomingMessage, type Server, type ServerResponse } from "node:http";
import { join } from "node:path";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { OAuthStore, type OAuthStoreOptions, verifyPassphrase } from "./auth.js";
import { createIWorkServer } from "./server.js";

interface PendingAuthorization { clientId: string; redirectUri: string; state?: string; codeChallenge: string; expiresAt: number }
interface ConsentAttempt { attempts: number; failures: number; expiresAt: number }
interface RegistrationAttempt { count: number; resetAt: number }
interface McpSession { transport: StreamableHTTPServerTransport; server: McpServer; clientId: string; tokenHash: string; lastUsedAt: number }

export interface HttpServerOptions {
  host: string;
  port: number;
  publicUrl: string;
  configDir: string;
  dataDir: string;
  environmentPassphrase?: string;
  platform?: NodeJS.Platform;
  redirectHostAllowlist?: string[];
  authorizationTtlMs?: number;
  sessionIdleTtlMs?: number;
  sweepIntervalMs?: number;
  passphraseBackoffBaseMs?: number;
  passphraseBackoffMaxMs?: number;
  passphraseBackoffDecayMs?: number;
  maxConsentFailures?: number;
  registrationLimit?: number;
  registrationWindowMs?: number;
  maxPendingAuthorizations?: number;
  maxSessions?: number;
  oauthStoreOptions?: OAuthStoreOptions;
}

function json(res: ServerResponse, status: number, value: unknown, headers: Record<string, string> = {}): void {
  res.writeHead(status, { "content-type": "application/json", "cache-control": "no-store", ...headers });
  res.end(JSON.stringify(value));
}

function html(res: ServerResponse, status: number, value: string, headers: Record<string, string> = {}): void {
  res.writeHead(status, {
    "content-type": "text/html; charset=utf-8",
    "cache-control": "no-store",
    "x-frame-options": "DENY",
    "content-security-policy": "default-src 'none'; style-src 'unsafe-inline'; base-uri 'none'",
    ...headers,
  });
  res.end(value);
}

async function body(req: IncomingMessage): Promise<string> {
  const chunks: Buffer[] = [];
  let size = 0;
  for await (const chunk of req) {
    const bytes = Buffer.from(chunk);
    size += bytes.length;
    if (size > 65_536) throw new Error("request_too_large");
    chunks.push(bytes);
  }
  return Buffer.concat(chunks).toString("utf8");
}

function escaped(value: string): string {
  return value.replace(/[&<>"']/g, character => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[character]!);
}

function bearer(req: IncomingMessage): string | undefined { return /^Bearer ([A-Za-z0-9_-]+)$/.exec(req.headers.authorization ?? "")?.[1]; }
function isLoopbackHostname(hostname: string): boolean { return ["127.0.0.1", "localhost", "[::1]"].includes(hostname.toLowerCase()); }

function isSafeRedirect(value: string, allowlist: Set<string>): boolean {
  try {
    const url = new URL(value);
    if (url.username || url.password || url.hash) return false;
    if (url.protocol !== "https:" && !(url.protocol === "http:" && isLoopbackHostname(url.hostname))) return false;
    if (allowlist.size === 0) return true;
    return allowlist.has(url.host.toLowerCase()) || allowlist.has(url.hostname.toLowerCase()) || allowlist.has(url.origin.toLowerCase());
  } catch { return false; }
}

function normalizeRedirectAllowlist(values: string[] = []): Set<string> {
  const entries = new Set<string>();
  for (const value of values) {
    const item = value.trim().toLowerCase();
    if (!item) continue;
    if (!/^https?:\/\//.test(item)) { entries.add(item); continue; }
    try { const url = new URL(item); entries.add(url.origin); entries.add(url.host); }
    catch { /* Invalid allowlist entries never match. */ }
  }
  return entries;
}

function requestHostAllowed(hostHeader: string | undefined, issuer: URL): boolean {
  if (!hostHeader) return false;
  if (hostHeader.toLowerCase() === issuer.host.toLowerCase()) return true;
  try { return isLoopbackHostname(new URL(`http://${hostHeader}`).hostname); }
  catch { return false; }
}

function requestOriginAllowed(originHeader: string | undefined, issuer: URL): boolean {
  if (!originHeader) return true;
  try {
    const origin = new URL(originHeader);
    return origin.origin === issuer.origin || isLoopbackHostname(origin.hostname);
  } catch { return false; }
}

export function createHttpServer(options: HttpServerOptions): Server {
  const issuerUrl = new URL(options.publicUrl);
  const issuer = issuerUrl.toString().replace(/\/$/, "");
  const resource = `${issuer}/mcp`;
  const store = new OAuthStore(join(options.dataDir, "oauth-state.json"), options.oauthStoreOptions);
  const pending = new Map<string, PendingAuthorization>();
  const attempts = new Map<string, ConsentAttempt>();
  const registrationAttempts = new Map<string, RegistrationAttempt>();
  const transports = new Map<string, McpSession>();
  const redirectAllowlist = normalizeRedirectAllowlist(options.redirectHostAllowlist);
  const authorizationTtlMs = options.authorizationTtlMs ?? 300_000;
  const sessionIdleTtlMs = options.sessionIdleTtlMs ?? 1_800_000;
  const sweepIntervalMs = options.sweepIntervalMs ?? 60_000;
  const backoffBaseMs = options.passphraseBackoffBaseMs ?? 500;
  const backoffMaxMs = options.passphraseBackoffMaxMs ?? 30_000;
  const backoffDecayMs = options.passphraseBackoffDecayMs ?? 60_000;
  const maxConsentFailures = options.maxConsentFailures ?? 5;
  const registrationLimit = options.registrationLimit ?? 20;
  const registrationWindowMs = options.registrationWindowMs ?? 60_000;
  const maxPendingAuthorizations = options.maxPendingAuthorizations ?? 1_000;
  const maxSessions = options.maxSessions ?? 100;
  const globalBackoff = { failures: 0, lastFailureAt: 0, blockedUntil: 0 };

  function decayBackoff(now: number): void {
    if (globalBackoff.failures === 0 || globalBackoff.lastFailureAt === 0) return;
    const decay = Math.floor((now - globalBackoff.lastFailureAt) / backoffDecayMs);
    if (decay > 0) {
      globalBackoff.failures = Math.max(0, globalBackoff.failures - decay);
      globalBackoff.lastFailureAt += decay * backoffDecayMs;
      if (globalBackoff.failures === 0) globalBackoff.blockedUntil = Math.min(globalBackoff.blockedUntil, now);
    }
  }

  function recordFailedPassphrase(now: number): number {
    decayBackoff(now);
    globalBackoff.failures = Math.min(8, globalBackoff.failures + 1);
    globalBackoff.lastFailureAt = now;
    const delay = Math.min(backoffMaxMs, backoffBaseMs * (2 ** (globalBackoff.failures - 1)));
    globalBackoff.blockedUntil = Math.max(globalBackoff.blockedUntil, now + delay);
    return delay;
  }

  async function closeSession(sessionId: string, session: McpSession): Promise<void> {
    transports.delete(sessionId);
    try { await session.server.close(); }
    catch { try { await session.transport.close(); } catch { /* already closed */ } }
  }

  const sweepTimer = setInterval(() => {
    const now = Date.now();
    for (const [requestId, authorization] of pending) if (authorization.expiresAt <= now) pending.delete(requestId);
    for (const [requestId, attempt] of attempts) if (attempt.expiresAt <= now || !pending.has(requestId)) attempts.delete(requestId);
    for (const [key, attempt] of registrationAttempts) if (attempt.resetAt <= now) registrationAttempts.delete(key);
    for (const [sessionId, session] of transports) if (session.lastUsedAt + sessionIdleTtlMs <= now) void closeSession(sessionId, session);
    decayBackoff(now);
    void store.sweep();
  }, sweepIntervalMs);
  sweepTimer.unref();

  const server = createServer(async (req, res) => {
    try {
      if (!requestHostAllowed(req.headers.host, issuerUrl)) return json(res, 421, { error: "invalid_host" });
      if (!requestOriginAllowed(typeof req.headers.origin === "string" ? req.headers.origin : undefined, issuerUrl)) return json(res, 403, { error: "invalid_origin" });
      const url = new URL(req.url ?? "/", issuer);

      if (req.method === "GET" && url.pathname === "/.well-known/oauth-protected-resource") {
        return json(res, 200, { resource, authorization_servers: [issuer], bearer_methods_supported: ["header"], resource_name: "iwork-mcp" });
      }
      if (req.method === "GET" && url.pathname === "/.well-known/oauth-authorization-server") {
        return json(res, 200, {
          issuer,
          authorization_endpoint: `${issuer}/authorize`,
          token_endpoint: `${issuer}/token`,
          revocation_endpoint: `${issuer}/revoke`,
          registration_endpoint: `${issuer}/register`,
          response_types_supported: ["code"],
          grant_types_supported: ["authorization_code", "refresh_token"],
          token_endpoint_auth_methods_supported: ["none"],
          revocation_endpoint_auth_methods_supported: ["none"],
          code_challenge_methods_supported: ["S256"],
          scopes_supported: ["mcp"],
        });
      }
      if (req.method === "POST" && url.pathname === "/register") {
        const now = Date.now();
        const key = req.socket.remoteAddress ?? "unknown";
        const existing = registrationAttempts.get(key);
        const rate = existing && existing.resetAt > now ? existing : { count: 0, resetAt: now + registrationWindowMs };
        rate.count += 1;
        registrationAttempts.set(key, rate);
        if (rate.count > registrationLimit) return json(res, 429, { error: "too_many_requests" }, { "retry-after": String(Math.max(1, Math.ceil((rate.resetAt - now) / 1000))) });
        const input = JSON.parse(await body(req)) as Record<string, unknown>;
        if (!Array.isArray(input.redirect_uris) || input.redirect_uris.length === 0 || input.redirect_uris.length > 10 || input.redirect_uris.some(item => typeof item !== "string" || !isSafeRedirect(item, redirectAllowlist))) return json(res, 400, { error: "invalid_redirect_uri" });
        if (input.client_name !== undefined && (typeof input.client_name !== "string" || input.client_name.length > 200)) return json(res, 400, { error: "invalid_client_metadata" });
        if (input.token_endpoint_auth_method !== undefined && input.token_endpoint_auth_method !== "none") return json(res, 400, { error: "invalid_client_metadata" });
        const client = await store.register({ redirect_uris: input.redirect_uris as string[], client_name: input.client_name as string | undefined });
        if (!client) return json(res, 429, { error: "registration_limit_reached" }, { "retry-after": "3600" });
        return json(res, 201, client);
      }
      if (req.method === "GET" && url.pathname === "/authorize") {
        const clientId = url.searchParams.get("client_id") ?? "";
        const redirectUri = url.searchParams.get("redirect_uri") ?? "";
        const challenge = url.searchParams.get("code_challenge") ?? "";
        const client = await store.client(clientId);
        if (!client || url.searchParams.get("response_type") !== "code" || url.searchParams.get("code_challenge_method") !== "S256" || !/^[A-Za-z0-9_-]{43}$/.test(challenge) || !client.redirect_uris.includes(redirectUri)) return json(res, 400, { error: "invalid_request" });
        if (pending.size >= maxPendingAuthorizations) return json(res, 429, { error: "too_many_authorization_requests" }, { "retry-after": "60" });
        const requestId = randomBytes(24).toString("base64url");
        pending.set(requestId, { clientId, redirectUri, state: url.searchParams.get("state") ?? undefined, codeChallenge: challenge, expiresAt: Date.now() + authorizationTtlMs });
        const redirectHost = new URL(redirectUri).host;
        return html(res, 200, `<!doctype html><html><head><meta name="viewport" content="width=device-width"><title>Authorize iwork-mcp</title><style>body{font:16px system-ui;max-width:32rem;margin:10vh auto;padding:1rem}.destination{border:2px solid;padding:1rem;border-radius:.5rem}input,button{box-sizing:border-box;font:inherit;padding:.7rem;width:100%;margin:.4rem 0}</style></head><body><h1>Authorize iwork-mcp</h1><p>Enter the owner passphrase to allow <strong>${escaped(client.client_name ?? "this client")}</strong> to use the server.</p><p class="destination">After approval, your browser will return to:<br><strong>${escaped(redirectHost)}</strong></p><form method="post" action="/authorize"><input type="hidden" name="request_id" value="${requestId}"><label>Passphrase<input type="password" name="passphrase" required autofocus></label><button type="submit">Allow access</button></form></body></html>`);
      }
      if (req.method === "POST" && url.pathname === "/authorize") {
        const form = new URLSearchParams(await body(req));
        const requestId = form.get("request_id") ?? "";
        const authorization = pending.get(requestId);
        const now = Date.now();
        if (!authorization || authorization.expiresAt <= now) {
          pending.delete(requestId);
          attempts.delete(requestId);
          return json(res, 400, { error: "invalid_request" });
        }
        decayBackoff(now);
        if (globalBackoff.blockedUntil > now) return json(res, 429, { error: "too_many_attempts" }, { "retry-after": String(Math.max(1, Math.ceil((globalBackoff.blockedUntil - now) / 1000))) });

        const attempt = attempts.get(requestId) ?? { attempts: 0, failures: 0, expiresAt: authorization.expiresAt };
        attempt.attempts += 1;
        attempts.set(requestId, attempt);
        // Reserve the global verification slot before scrypt starts, so concurrent
        // submissions cannot all enter the expensive verification path.
        globalBackoff.blockedUntil = now + backoffBaseMs;
        const accepted = await verifyPassphrase(options.configDir, form.get("passphrase") ?? "", options.environmentPassphrase);
        if (!accepted) {
          attempt.failures += 1;
          const delay = recordFailedPassphrase(Date.now());
          if (attempt.failures >= maxConsentFailures) { pending.delete(requestId); attempts.delete(requestId); }
          return html(res, 403, "<!doctype html><html><body><h1>Access denied</h1><p>The passphrase was not accepted.</p></body></html>", { "retry-after": String(Math.max(1, Math.ceil(delay / 1000))) });
        }

        globalBackoff.failures = Math.max(0, globalBackoff.failures - 2);
        globalBackoff.blockedUntil = Date.now();
        pending.delete(requestId);
        attempts.delete(requestId);
        const code = await store.issueCode(authorization.clientId, authorization.redirectUri, authorization.codeChallenge);
        const redirect = new URL(authorization.redirectUri);
        redirect.searchParams.set("code", code);
        if (authorization.state !== undefined) redirect.searchParams.set("state", authorization.state);
        res.writeHead(302, { location: redirect.toString(), "cache-control": "no-store" });
        return res.end();
      }
      if (req.method === "POST" && url.pathname === "/token") {
        const form = new URLSearchParams(await body(req));
        const grant = form.get("grant_type");
        const clientId = form.get("client_id") ?? "";
        if (!await store.client(clientId)) return json(res, 400, { error: "invalid_client" });
        const tokens = grant === "authorization_code"
          ? await store.exchangeCode(form.get("code") ?? "", clientId, form.get("redirect_uri") ?? "", form.get("code_verifier") ?? "")
          : grant === "refresh_token"
            ? await store.refresh(form.get("refresh_token") ?? "", clientId)
            : undefined;
        if (grant !== "authorization_code" && grant !== "refresh_token") return json(res, 400, { error: "unsupported_grant_type" });
        if (!tokens) return json(res, 400, { error: "invalid_grant" });
        return json(res, 200, { ...tokens, token_type: "Bearer", expires_in: Math.floor((options.oauthStoreOptions?.accessTokenTtlMs ?? 900_000) / 1000), scope: "mcp" });
      }
      if (req.method === "POST" && url.pathname === "/revoke") {
        const form = new URLSearchParams(await body(req));
        const clientId = form.get("client_id") ?? "";
        if (await store.client(clientId)) await store.revoke(form.get("token") ?? "", clientId);
        return json(res, 200, {});
      }
      if (url.pathname === "/mcp") {
        const token = bearer(req);
        const identity = token ? await store.accessToken(token) : undefined;
        if (!identity) return json(res, 401, { error: "invalid_token" }, { "www-authenticate": `Bearer resource_metadata="${issuer}/.well-known/oauth-protected-resource", scope="mcp"` });
        const sessionId = req.headers["mcp-session-id"] as string | undefined;
        let current = sessionId ? transports.get(sessionId) : undefined;
        if (current && current.lastUsedAt + sessionIdleTtlMs <= Date.now()) {
          await closeSession(sessionId!, current);
          current = undefined;
        }
        if (sessionId && !current) return json(res, 404, { error: "unknown_session" });
        if (current && (current.clientId !== identity.clientId || current.tokenHash !== identity.tokenHash)) return json(res, 403, { error: "session_token_mismatch" });
        if (!current) {
          if (transports.size >= maxSessions) return json(res, 429, { error: "too_many_sessions" }, { "retry-after": "60" });
          let mcpServer: McpServer;
          let transport: StreamableHTTPServerTransport;
          const allowedHosts = [...new Set([issuerUrl.host, "localhost", "127.0.0.1", "[::1]", req.headers.host!])];
          transport = new StreamableHTTPServerTransport({
            sessionIdGenerator: randomUUID,
            allowedHosts,
            enableDnsRebindingProtection: true,
            onsessioninitialized: (id): void => { transports.set(id, { transport, server: mcpServer, clientId: identity.clientId, tokenHash: identity.tokenHash, lastUsedAt: Date.now() }); },
            onsessionclosed: (id): void => { transports.delete(id); },
          });
          mcpServer = createIWorkServer({ platform: options.platform });
          transport.onclose = () => { if (transport.sessionId) transports.delete(transport.sessionId); };
          await mcpServer.connect(transport);
          current = { transport, server: mcpServer, clientId: identity.clientId, tokenHash: identity.tokenHash, lastUsedAt: Date.now() };
        }
        current.lastUsedAt = Date.now();
        return current.transport.handleRequest(req, res);
      }
      return json(res, 404, { error: "not_found" });
    } catch (error) {
      if (!res.headersSent) return json(res, 400, { error: error instanceof Error ? error.message : "bad_request" });
      if (!res.writableEnded) res.destroy(error instanceof Error ? error : undefined);
    }
  });

  server.on("close", () => {
    clearInterval(sweepTimer);
    for (const [sessionId, session] of transports) void closeSession(sessionId, session);
    void store.close();
  });
  return server;
}

export async function listenHttp(options: HttpServerOptions): Promise<Server> {
  const server = createHttpServer(options);
  await new Promise<void>((resolve, reject) => { server.once("error", reject); server.listen(options.port, options.host, resolve); });
  return server;
}
