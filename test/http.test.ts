import { createHash } from "node:crypto";
import { mkdtemp, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { after, before, describe, it, type TestContext } from "node:test";
import assert from "node:assert/strict";
import type { AddressInfo } from "node:net";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js";
import { createHttpServer } from "../src/http.js";

describe("remote HTTP and OAuth", () => {
  let directory: string;
  let baseUrl: string;
  let server: ReturnType<typeof createHttpServer>;
  let clientId: string;
  let accessToken: string;
  let refreshToken: string;
  let socketUnavailable = false;
  const redirectUri = "http://127.0.0.1/callback";
  const passphrase = "a-test-passphrase-only";
  const verifier = "test-verifier-that-is-long-enough-for-pkce-0123456789";
  const challenge = createHash("sha256").update(verifier).digest("base64url");

  before(async () => {
    directory = await mkdtemp(join(tmpdir(), "iwork-mcp-test-"));
    server = createHttpServer({ host: "127.0.0.1", port: 0, publicUrl: "http://127.0.0.1", configDir: join(directory, "config"), dataDir: join(directory, "data"), environmentPassphrase: passphrase, platform: "darwin" });
    try {
      await new Promise<void>((resolve, reject) => { server.once("error", reject); server.listen(0, "127.0.0.1", resolve); });
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === "EPERM") { socketUnavailable = true; return; }
      throw error;
    }
    baseUrl = `http://127.0.0.1:${(server.address() as AddressInfo).port}`;
  });

  after(async () => {
    if (!socketUnavailable) await new Promise<void>((resolve, reject) => server.close(error => error ? reject(error) : resolve()));
    await rm(directory, { recursive: true });
  });

  function requireSocket(context: TestContext): boolean {
    if (socketUnavailable) context.skip("socket binding is unavailable in this sandbox");
    return !socketUnavailable;
  }

  it("serves protected-resource and authorization-server metadata", async context => {
    if (!requireSocket(context)) return;
    const resource = await fetch(`${baseUrl}/.well-known/oauth-protected-resource`).then(response => response.json()) as Record<string, unknown>;
    const authorization = await fetch(`${baseUrl}/.well-known/oauth-authorization-server`).then(response => response.json()) as Record<string, unknown>;
    assert.equal(resource.resource, "http://127.0.0.1/mcp");
    assert.deepEqual(authorization.code_challenge_methods_supported, ["S256"]);
    assert.deepEqual(authorization.grant_types_supported, ["authorization_code", "refresh_token"]);
  });

  it("dynamically registers a public client and rejects unsafe redirects", async context => {
    if (!requireSocket(context)) return;
    const bad = await fetch(`${baseUrl}/register`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ redirect_uris: ["http://remote.invalid/callback"] }) });
    assert.equal(bad.status, 400);
    const response = await fetch(`${baseUrl}/register`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ redirect_uris: [redirectUri], client_name: "Test Agent" }) });
    assert.equal(response.status, 201);
    const registration = await response.json() as { client_id: string; token_endpoint_auth_method: string };
    clientId = registration.client_id;
    assert.equal(registration.token_endpoint_auth_method, "none");
    assert.equal((await stat(join(directory, "data", "oauth-state.json"))).mode & 0o777, 0o600);
  });

  it("requires exact redirects, S256 PKCE, and the correct passphrase", async context => {
    if (!requireSocket(context)) return;
    const badRedirect = new URL(`${baseUrl}/authorize`);
    badRedirect.search = new URLSearchParams({ response_type: "code", client_id: clientId, redirect_uri: `${redirectUri}/wrong`, code_challenge: challenge, code_challenge_method: "S256" }).toString();
    assert.equal((await fetch(badRedirect)).status, 400);

    const authorize = new URL(`${baseUrl}/authorize`);
    authorize.search = new URLSearchParams({ response_type: "code", client_id: clientId, redirect_uri: redirectUri, code_challenge: challenge, code_challenge_method: "S256", state: "kept" }).toString();
    const page = await fetch(authorize);
    assert.equal(page.status, 200);
    const requestId = (await page.text()).match(/name="request_id" value="([^"]+)"/)?.[1];
    assert.ok(requestId);
    const wrong = await fetch(`${baseUrl}/authorize`, { method: "POST", headers: { "content-type": "application/x-www-form-urlencoded" }, body: new URLSearchParams({ request_id: requestId, passphrase: "wrong-passphrase" }) });
    assert.equal(wrong.status, 403);
    const consent = await fetch(`${baseUrl}/authorize`, { method: "POST", headers: { "content-type": "application/x-www-form-urlencoded" }, body: new URLSearchParams({ request_id: requestId, passphrase }), redirect: "manual" });
    assert.equal(consent.status, 302);
    const redirected = new URL(consent.headers.get("location")!);
    assert.equal(redirected.searchParams.get("state"), "kept");
    const code = redirected.searchParams.get("code")!;
    const badToken = await fetch(`${baseUrl}/token`, { method: "POST", body: new URLSearchParams({ grant_type: "authorization_code", client_id: clientId, redirect_uri: redirectUri, code, code_verifier: `${verifier}bad` }) });
    assert.equal(badToken.status, 400);
    const tokenResponse = await fetch(`${baseUrl}/token`, { method: "POST", body: new URLSearchParams({ grant_type: "authorization_code", client_id: clientId, redirect_uri: redirectUri, code, code_verifier: verifier }) });
    assert.equal(tokenResponse.status, 200);
    const tokens = await tokenResponse.json() as { access_token: string; refresh_token: string };
    accessToken = tokens.access_token; refreshToken = tokens.refresh_token;
  });

  it("rejects missing and invalid bearer tokens with a challenge", async context => {
    if (!requireSocket(context)) return;
    for (const headers of [{}, { authorization: "Bearer invalid" }]) {
      const response = await fetch(`${baseUrl}/mcp`, { method: "POST", headers: { "content-type": "application/json", ...headers }, body: "{}" });
      assert.equal(response.status, 401);
      assert.match(response.headers.get("www-authenticate") ?? "", /resource_metadata=/);
    }
  });

  it("refreshes tokens and rotates the refresh token", async context => {
    if (!requireSocket(context)) return;
    const response = await fetch(`${baseUrl}/token`, { method: "POST", body: new URLSearchParams({ grant_type: "refresh_token", client_id: clientId, refresh_token: refreshToken }) });
    assert.equal(response.status, 200);
    const tokens = await response.json() as { access_token: string; refresh_token: string };
    assert.notEqual(tokens.access_token, accessToken); assert.notEqual(tokens.refresh_token, refreshToken); accessToken = tokens.access_token;
    const reuse = await fetch(`${baseUrl}/token`, { method: "POST", body: new URLSearchParams({ grant_type: "refresh_token", client_id: clientId, refresh_token: refreshToken }) });
    assert.equal(reuse.status, 400);
  });

  it("initializes MCP and lists the same 117 tools over HTTP", async context => {
    if (!requireSocket(context)) return;
    const client = new Client({ name: "http-test", version: "0.0.0" });
    const transport = new StreamableHTTPClientTransport(new URL(`${baseUrl}/mcp`), { requestInit: { headers: { authorization: `Bearer ${accessToken}` } } });
    await client.connect(transport);
    const { tools } = await client.listTools();
    assert.equal(tools.length, 117);
    await client.close();
  });
});
