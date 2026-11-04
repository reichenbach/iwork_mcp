import { createHash, randomUUID } from "node:crypto";
import { mkdtemp, readFile, rm, stat } from "node:fs/promises";
import { request as httpRequest } from "node:http";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { after, before, describe, it, type TestContext } from "node:test";
import assert from "node:assert/strict";
import type { AddressInfo } from "node:net";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js";
import { OAuthStore } from "../src/auth.js";
import { createHttpServer, registrationClientKey } from "../src/http.js";

const sleep = (milliseconds: number): Promise<void> => new Promise(resolve => setTimeout(resolve, milliseconds));
const verifier = "test-verifier-that-is-long-enough-for-pkce-0123456789";
const challenge = createHash("sha256").update(verifier).digest("base64url");

describe("OAuth token store", () => {
  let directory: string;

  before(async () => { directory = await mkdtemp(join(tmpdir(), "iwork-mcp-auth-test-")); });
  after(async () => { await rm(directory, { recursive: true }); });

  it("hashes codes and tokens at rest and rejects code reuse", async () => {
    const statePath = join(directory, "hashed-state.json");
    const store = new OAuthStore(statePath, { persistDebounceMs: 5 });
    const client = (await store.register({ redirect_uris: ["http://127.0.0.1/callback"] }))!;
    const code = await store.issueCode(client.client_id, client.redirect_uris[0], challenge);
    const tokens = await store.exchangeCode(code, client.client_id, client.redirect_uris[0], verifier);
    assert.ok(tokens);
    assert.equal(await store.exchangeCode(code, client.client_id, client.redirect_uris[0], verifier), undefined);
    await store.flush();
    const persisted = await readFile(statePath, "utf8");
    assert.equal(persisted.includes(code), false);
    assert.equal(persisted.includes(tokens.access_token), false);
    assert.equal(persisted.includes(tokens.refresh_token), false);
    assert.equal((await stat(statePath)).mode & 0o777, 0o600);
    await store.close();
  });

  it("rejects expired codes plus missing and plain PKCE", async () => {
    const store = new OAuthStore(join(directory, "expiry-state.json"), { codeTtlMs: 5 });
    const client = (await store.register({ redirect_uris: ["http://127.0.0.1/callback"] }))!;
    const expired = await store.issueCode(client.client_id, client.redirect_uris[0], challenge);
    await sleep(10);
    assert.equal(await store.exchangeCode(expired, client.client_id, client.redirect_uris[0], verifier), undefined);
    const missing = await store.issueCode(client.client_id, client.redirect_uris[0], challenge);
    assert.equal(await store.exchangeCode(missing, client.client_id, client.redirect_uris[0], ""), undefined);
    const plain = await store.issueCode(client.client_id, client.redirect_uris[0], verifier);
    assert.equal(await store.exchangeCode(plain, client.client_id, client.redirect_uris[0], verifier), undefined);
    await store.close();
  });

  it("rejects refresh from another client and revokes a family on reuse", async () => {
    const store = new OAuthStore(join(directory, "rotation-state.json"));
    const first = (await store.register({ redirect_uris: ["http://127.0.0.1/one"] }))!;
    const second = (await store.register({ redirect_uris: ["http://127.0.0.1/two"] }))!;
    const code = await store.issueCode(first.client_id, first.redirect_uris[0], challenge);
    const initial = (await store.exchangeCode(code, first.client_id, first.redirect_uris[0], verifier))!;
    const initialIdentity = (await store.accessToken(initial.access_token))!;
    assert.equal(await store.refresh(initial.refresh_token, second.client_id), undefined);
    const rotated = await store.refresh(initial.refresh_token, first.client_id);
    assert.ok(rotated);
    const rotatedIdentity = (await store.accessToken(rotated.access_token))!;
    assert.equal(rotatedIdentity.familyId, initialIdentity.familyId);
    assert.notEqual(rotatedIdentity.tokenHash, initialIdentity.tokenHash);
    assert.equal(await store.validAccessToken(rotated.access_token), true);
    assert.equal(await store.refresh(initial.refresh_token, first.client_id), undefined);
    assert.equal(await store.validAccessToken(initial.access_token), false);
    assert.equal(await store.validAccessToken(rotated.access_token), false);
    await store.close();
  });

  it("does not keep an unissued client alive when it is looked up", async () => {
    const store = new OAuthStore(join(directory, "unused-state.json"), { clientUnusedTtlMs: 50 });
    const client = (await store.register({ redirect_uris: ["http://127.0.0.1/one"] }))!;
    await sleep(40);
    assert.ok(await store.client(client.client_id));
    await sleep(20);
    await store.sweep();
    assert.equal(await store.client(client.client_id), undefined);
    await store.close();
  });

  it("evicts the oldest client without tokens at the cap and rejects when all have tokens", async () => {
    const store = new OAuthStore(join(directory, "bounded-state.json"), { maxClients: 2, clientUnusedTtlMs: 60_000 });
    const oldest = (await store.register({ redirect_uris: ["http://127.0.0.1/oldest"] }))!;
    await sleep(5);
    const retained = (await store.register({ redirect_uris: ["http://127.0.0.1/retained"] }))!;
    const replacement = (await store.register({ redirect_uris: ["http://127.0.0.1/replacement"] }))!;
    assert.equal(await store.client(oldest.client_id), undefined);
    assert.ok(await store.client(retained.client_id));
    assert.ok(await store.client(replacement.client_id));

    await sleep(5);
    const retainedCode = await store.issueCode(retained.client_id, retained.redirect_uris[0], challenge);
    const newest = (await store.register({ redirect_uris: ["http://127.0.0.1/newest"] }))!;
    assert.ok(await store.client(retained.client_id));
    assert.equal(await store.client(replacement.client_id), undefined);

    assert.ok(await store.exchangeCode(retainedCode, retained.client_id, retained.redirect_uris[0], verifier));
    const newestCode = await store.issueCode(newest.client_id, newest.redirect_uris[0], challenge);
    assert.ok(await store.exchangeCode(newestCode, newest.client_id, newest.redirect_uris[0], verifier));
    assert.equal(await store.register({ redirect_uris: ["http://127.0.0.1/rejected"] }), undefined);
    await store.close();
  });
});

describe("registration client keys", () => {
  it("trusts CF-Connecting-IP only from loopback socket peers", () => {
    for (const peer of ["127.0.0.1", "127.42.0.2", "::1", "::ffff:127.0.0.1", "::ffff:127.42.0.2", "::ffff:7f00:1"]) {
      assert.equal(registrationClientKey(peer, "198.51.100.1"), "198.51.100.1");
    }
    for (const [peer, key] of [["192.0.2.1", "192.0.2.1"], ["128.0.0.1", "128.0.0.1"], ["::ffff:192.0.2.1", "192.0.2.1"], ["2001:db8::1", "2001:db8::/64"], [undefined, "unknown"]]) {
      assert.equal(registrationClientKey(peer, "198.51.100.1"), key);
      assert.equal(registrationClientKey(peer, "198.51.100.2"), key);
    }
  });

  it("uses only the first valid IP value and falls back for invalid headers", () => {
    assert.equal(registrationClientKey("::1", " 198.51.100.1, 198.51.100.2"), "198.51.100.1");
    assert.equal(registrationClientKey("::1", ["198.51.100.1", "198.51.100.2"]), "198.51.100.1");
    assert.equal(registrationClientKey("::1", "2001:0DB8:0:0:0:0:0:1"), "2001:db8::/64");
    for (const value of [undefined, [], "", "garbage", "999.0.0.1", "198.51.100.1:80", "[::1]", "fe80::1%lo0", ", 198.51.100.1", "bad, 198.51.100.1", ["bad", "198.51.100.1"]]) {
      assert.equal(registrationClientKey("127.0.0.1", value), "127.0.0.1");
    }
  });

  it("groups IPv6 addresses by /64 for headers and socket peers", () => {
    for (const address of ["2001:db8:abcd:1234::1", "2001:0DB8:abcd:1234:ffff:ffff:ffff:ffff", "2001:db8:abcd:1234::"]) {
      assert.equal(registrationClientKey("127.0.0.1", address), "2001:db8:abcd:1234::/64");
      assert.equal(registrationClientKey(address, undefined), "2001:db8:abcd:1234::/64");
    }
    assert.equal(registrationClientKey("::1", "2001:db8:abcd:1235::1"), "2001:db8:abcd:1235::/64");
    assert.equal(registrationClientKey("::1", "2001:db8::1"), "2001:db8::/64");
    assert.equal(registrationClientKey("::1", undefined), "::/64");
    assert.equal(registrationClientKey("::1", "::"), "::/64");
  });

  it("normalizes IPv4-mapped IPv6 addresses to IPv4 for headers and socket peers", () => {
    for (const address of ["192.0.2.1", "::ffff:192.0.2.1", "::FFFF:c000:201", "0:0:0:0:0:ffff:c000:0201"]) {
      assert.equal(registrationClientKey("127.0.0.1", address), "192.0.2.1");
      assert.equal(registrationClientKey(address, undefined), "192.0.2.1");
    }
    assert.equal(registrationClientKey("::1", "::ffff:192.0.2.2"), "192.0.2.2");
  });
});

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

  before(async () => {
    directory = await mkdtemp(join(tmpdir(), "iwork-mcp-test-"));
    server = createHttpServer({ host: "127.0.0.1", port: 0, publicUrl: "http://127.0.0.1", configDir: join(directory, "config"), dataDir: join(directory, "data"), environmentPassphrase: passphrase, platform: "darwin", redirectHostAllowlist: ["127.0.0.1"], passphraseBackoffBaseMs: 5, passphraseBackoffMaxMs: 20, passphraseBackoffDecayMs: 20 });
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

  async function authorizeTokens(targetClientId = clientId): Promise<{ access_token: string; refresh_token: string }> {
    const authorize = new URL(`${baseUrl}/authorize`);
    authorize.search = new URLSearchParams({ response_type: "code", client_id: targetClientId, redirect_uri: redirectUri, code_challenge: challenge, code_challenge_method: "S256" }).toString();
    const page = await fetch(authorize);
    const requestId = (await page.text()).match(/name="request_id" value="([^"]+)"/)?.[1];
    assert.ok(requestId);
    const consent = await fetch(`${baseUrl}/authorize`, { method: "POST", headers: { "content-type": "application/x-www-form-urlencoded" }, body: new URLSearchParams({ request_id: requestId, passphrase }), redirect: "manual" });
    assert.equal(consent.status, 302);
    const code = new URL(consent.headers.get("location")!).searchParams.get("code")!;
    const response = await fetch(`${baseUrl}/token`, { method: "POST", body: new URLSearchParams({ grant_type: "authorization_code", client_id: targetClientId, redirect_uri: redirectUri, code, code_verifier: verifier }) });
    assert.equal(response.status, 200);
    return response.json() as Promise<{ access_token: string; refresh_token: string }>;
  }

  async function startSupplementalServer(name: string, overrides: Partial<Parameters<typeof createHttpServer>[0]> = {}): Promise<{ supplemental: ReturnType<typeof createHttpServer>; url: string }> {
    const supplemental = createHttpServer({
      host: "127.0.0.1",
      port: 0,
      publicUrl: "http://127.0.0.1",
      configDir: join(directory, `${name}-config`),
      dataDir: join(directory, `${name}-data`),
      environmentPassphrase: passphrase,
      platform: "darwin",
      ...overrides,
    });
    await new Promise<void>((resolve, reject) => { supplemental.once("error", reject); supplemental.listen(0, "127.0.0.1", resolve); });
    return { supplemental, url: `http://127.0.0.1:${(supplemental.address() as AddressInfo).port}` };
  }

  async function closeServer(target: ReturnType<typeof createHttpServer>): Promise<void> {
    await new Promise<void>((resolve, reject) => target.close(error => error ? reject(error) : resolve()));
  }

  async function registerViaHttp(url: string, connectingIp?: string | string[]): Promise<{ status: number; retryAfter: string | undefined; body: unknown }> {
    return new Promise((resolve, reject) => {
      const request = httpRequest(`${url}/register`, {
        method: "POST",
        headers: { "content-type": "application/json", ...(connectingIp === undefined ? {} : { "CF-Connecting-IP": connectingIp }) },
      }, response => {
        const chunks: Buffer[] = [];
        response.on("data", chunk => chunks.push(Buffer.from(chunk)));
        response.once("error", reject);
        response.once("end", () => {
          try { resolve({ status: response.statusCode ?? 0, retryAfter: response.headers["retry-after"], body: JSON.parse(Buffer.concat(chunks).toString("utf8")) }); }
          catch (error) { reject(error); }
        });
      });
      request.once("error", reject);
      request.end(JSON.stringify({ redirect_uris: [redirectUri] }));
    });
  }

  async function registerAndAuthorize(url: string): Promise<{ clientId: string; accessToken: string }> {
    const registrationResponse = await fetch(`${url}/register`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ redirect_uris: [redirectUri], client_name: "Short-lived Agent" }) });
    assert.equal(registrationResponse.status, 201);
    const registration = await registrationResponse.json() as { client_id: string };
    const authorize = new URL(`${url}/authorize`);
    authorize.search = new URLSearchParams({ response_type: "code", client_id: registration.client_id, redirect_uri: redirectUri, code_challenge: challenge, code_challenge_method: "S256" }).toString();
    const requestId = (await (await fetch(authorize)).text()).match(/name="request_id" value="([^"]+)"/)?.[1];
    assert.ok(requestId);
    const approval = await fetch(`${url}/authorize`, { method: "POST", body: new URLSearchParams({ request_id: requestId, passphrase }), redirect: "manual" });
    assert.equal(approval.status, 302);
    const code = new URL(approval.headers.get("location")!).searchParams.get("code")!;
    const tokenResponse = await fetch(`${url}/token`, { method: "POST", body: new URLSearchParams({ grant_type: "authorization_code", client_id: registration.client_id, redirect_uri: redirectUri, code, code_verifier: verifier }) });
    assert.equal(tokenResponse.status, 200);
    const tokens = await tokenResponse.json() as { access_token: string };
    return { clientId: registration.client_id, accessToken: tokens.access_token };
  }

  it("serves protected-resource and authorization-server metadata", async context => {
    if (!requireSocket(context)) return;
    const resource = await fetch(`${baseUrl}/.well-known/oauth-protected-resource`).then(response => response.json()) as Record<string, unknown>;
    const authorization = await fetch(`${baseUrl}/.well-known/oauth-authorization-server`).then(response => response.json()) as Record<string, unknown>;
    assert.equal(resource.resource, "http://127.0.0.1/mcp");
    assert.deepEqual(authorization.code_challenge_methods_supported, ["S256"]);
    assert.deepEqual(authorization.grant_types_supported, ["authorization_code", "refresh_token"]);
    assert.equal(authorization.revocation_endpoint, "http://127.0.0.1/revoke");
  });

  it("gives loopback-proxied clients separate registration budgets", async context => {
    if (!requireSocket(context)) return;
    const { supplemental, url } = await startSupplementalServer("registration-clients", { registrationLimit: 1 });
    try {
      assert.equal((await registerViaHttp(url, "198.51.100.1")).status, 201);
      const limited = await registerViaHttp(url, "198.51.100.1, 198.51.100.99");
      assert.equal(limited.status, 429);
      assert.deepEqual(limited.body, { error: "too_many_requests" });
      assert.ok(Number(limited.retryAfter) > 0);
      assert.equal((await registerViaHttp(url, ["198.51.100.2", "198.51.100.1"])).status, 201);
      assert.equal((await registerViaHttp(url, "198.51.100.2")).status, 429);
      assert.equal((await registerViaHttp(url)).status, 201);
      assert.equal((await registerViaHttp(url, "invalid, 198.51.100.3")).status, 429);
    } finally { await closeServer(supplemental); }
  });

  it("does not charge client-limited rejections against the global registration budget", async context => {
    if (!requireSocket(context)) return;
    const { supplemental, url } = await startSupplementalServer("registration-rejections", { registrationLimit: 1, globalRegistrationLimit: 2 });
    try {
      assert.equal((await registerViaHttp(url, "198.51.100.1")).status, 201);
      for (let count = 0; count < 3; count += 1) {
        const limited = await registerViaHttp(url, "198.51.100.1");
        assert.equal(limited.status, 429);
        assert.deepEqual(limited.body, { error: "too_many_requests" });
        assert.ok(Number(limited.retryAfter) > 0);
      }
      assert.equal((await registerViaHttp(url, "198.51.100.2")).status, 201);
      assert.equal((await registerViaHttp(url, "198.51.100.3")).status, 429);
    } finally { await closeServer(supplemental); }
  });

  it("shares registration budgets across IPv6 /64 addresses and mapped IPv4 spellings", async context => {
    if (!requireSocket(context)) return;
    const { supplemental, url } = await startSupplementalServer("registration-normalization", { registrationLimit: 1 });
    try {
      assert.equal((await registerViaHttp(url, "2001:db8:abcd:1234::1")).status, 201);
      assert.equal((await registerViaHttp(url, "2001:0DB8:abcd:1234:ffff:ffff:ffff:ffff")).status, 429);
      assert.equal((await registerViaHttp(url, "2001:db8:abcd:1235::1")).status, 201);
      assert.equal((await registerViaHttp(url, "192.0.2.1")).status, 201);
      assert.equal((await registerViaHttp(url, "::ffff:192.0.2.1")).status, 429);
      assert.equal((await registerViaHttp(url, "::ffff:c000:201")).status, 429);
    } finally { await closeServer(supplemental); }
  });

  it("caps registration attempts globally across client keys and bucket eviction", async context => {
    if (!requireSocket(context)) return;
    const { supplemental, url } = await startSupplementalServer("registration-global", { registrationLimit: 2, globalRegistrationLimit: 3, maxRegistrationBuckets: 1 });
    try {
      for (const ip of ["198.51.100.1", "198.51.100.2", "198.51.100.3"]) assert.equal((await registerViaHttp(url, ip)).status, 201);
      for (const ip of ["198.51.100.4", "198.51.100.1"]) {
        const limited = await registerViaHttp(url, ip);
        assert.equal(limited.status, 429);
        assert.deepEqual(limited.body, { error: "too_many_requests" });
        assert.ok(Number(limited.retryAfter) > 0);
      }
    } finally { await closeServer(supplemental); }
  });

  it("evicts the oldest registration bucket at the memory cap", async context => {
    if (!requireSocket(context)) return;
    const { supplemental, url } = await startSupplementalServer("registration-buckets", { registrationLimit: 1, maxRegistrationBuckets: 2 });
    try {
      assert.equal((await registerViaHttp(url, "198.51.100.1")).status, 201);
      assert.equal((await registerViaHttp(url, "198.51.100.2")).status, 201);
      assert.equal((await registerViaHttp(url, "198.51.100.1")).status, 429);
      assert.equal((await registerViaHttp(url, "198.51.100.3")).status, 201);
      assert.equal((await registerViaHttp(url, "198.51.100.2")).status, 429);
      assert.equal((await registerViaHttp(url, "198.51.100.1")).status, 201);
    } finally { await closeServer(supplemental); }
  });

  it("renews expired client and global registration budgets before a sweep", async context => {
    if (!requireSocket(context)) return;
    const { supplemental, url } = await startSupplementalServer("registration-expiry", { registrationLimit: 1, globalRegistrationLimit: 2, registrationWindowMs: 100, sweepIntervalMs: 60_000 });
    try {
      assert.equal((await registerViaHttp(url, "198.51.100.1")).status, 201);
      assert.equal((await registerViaHttp(url, "198.51.100.1")).status, 429);
      assert.equal((await registerViaHttp(url, "198.51.100.2")).status, 201);
      assert.equal((await registerViaHttp(url, "198.51.100.3")).status, 429);
      await sleep(150);
      assert.equal((await registerViaHttp(url, "198.51.100.1")).status, 201);
      assert.equal((await registerViaHttp(url, "198.51.100.2")).status, 201);
    } finally { await closeServer(supplemental); }
  });

  it("dynamically registers a public client and rejects unsafe redirects", async context => {
    if (!requireSocket(context)) return;
    for (const invalidRedirect of ["http://remote.invalid/callback", "https://not-allowed.invalid/callback"]) {
      const bad = await fetch(`${baseUrl}/register`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ redirect_uris: [invalidRedirect] }) });
      assert.equal(bad.status, 400);
    }
    const response = await fetch(`${baseUrl}/register`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ redirect_uris: [redirectUri], client_name: "Test <Agent>" }) });
    assert.equal(response.status, 201);
    const registration = await response.json() as { client_id: string; token_endpoint_auth_method: string };
    clientId = registration.client_id;
    assert.equal(registration.token_endpoint_auth_method, "none");
    await sleep(75);
    assert.equal((await stat(join(directory, "data", "oauth-state.json"))).mode & 0o777, 0o600);
  });

  it("requires exact redirects, S256 PKCE, and the correct passphrase", async context => {
    if (!requireSocket(context)) return;
    const badRedirect = new URL(`${baseUrl}/authorize`);
    badRedirect.search = new URLSearchParams({ response_type: "code", client_id: clientId, redirect_uri: `${redirectUri}/wrong`, code_challenge: challenge, code_challenge_method: "S256" }).toString();
    assert.equal((await fetch(badRedirect)).status, 400);
    for (const parameters of [
      { response_type: "code", client_id: clientId, redirect_uri: redirectUri },
      { response_type: "code", client_id: clientId, redirect_uri: redirectUri, code_challenge: verifier, code_challenge_method: "plain" },
    ]) {
      const invalidPkce = new URL(`${baseUrl}/authorize`);
      invalidPkce.search = new URLSearchParams(parameters).toString();
      assert.equal((await fetch(invalidPkce)).status, 400);
    }

    const authorize = new URL(`${baseUrl}/authorize`);
    authorize.search = new URLSearchParams({ response_type: "code", client_id: clientId, redirect_uri: redirectUri, code_challenge: challenge, code_challenge_method: "S256", state: "kept" }).toString();
    const page = await fetch(authorize);
    assert.equal(page.status, 200);
    assert.doesNotMatch(page.headers.get("content-security-policy") ?? "", /form-action/);
    const pageBody = await page.text();
    assert.match(pageBody, /<strong>127\.0\.0\.1<\/strong>/);
    assert.match(pageBody, /<strong>Test &lt;Agent&gt;<\/strong>/);
    const requestId = pageBody.match(/name="request_id" value="([^"]+)"/)?.[1];
    assert.ok(requestId);
    const wrong = await fetch(`${baseUrl}/authorize`, { method: "POST", headers: { "content-type": "application/x-www-form-urlencoded" }, body: new URLSearchParams({ request_id: requestId, passphrase: "wrong-passphrase" }) });
    assert.equal(wrong.status, 403);
    await sleep(25);
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
    const codeReuse = await fetch(`${baseUrl}/token`, { method: "POST", body: new URLSearchParams({ grant_type: "authorization_code", client_id: clientId, redirect_uri: redirectUri, code, code_verifier: verifier }) });
    assert.equal(codeReuse.status, 400);
  });

  it("rejects DNS-rebinding Host and cross-origin requests", async context => {
    if (!requireSocket(context)) return;
    const badHostStatus = await new Promise<number>((resolve, reject) => {
      const request = httpRequest(`${baseUrl}/.well-known/oauth-authorization-server`, { headers: { host: "attacker.invalid" } }, response => {
        response.resume();
        response.once("end", () => resolve(response.statusCode ?? 0));
      });
      request.once("error", reject);
      request.end();
    });
    assert.equal(badHostStatus, 421);
    const badOrigin = await fetch(`${baseUrl}/.well-known/oauth-authorization-server`, { headers: { origin: "https://attacker.invalid" } });
    assert.equal(badOrigin.status, 403);
  });

  it("rejects missing and invalid bearer tokens with a challenge", async context => {
    if (!requireSocket(context)) return;
    for (const headers of [{}, { authorization: "Bearer invalid" }]) {
      const response = await fetch(`${baseUrl}/mcp`, { method: "POST", headers: { "content-type": "application/json", ...headers }, body: "{}" });
      assert.equal(response.status, 401);
      assert.match(response.headers.get("www-authenticate") ?? "", /resource_metadata=/);
    }
  });

  it("rejects the wrong refresh client and revokes the family on reuse", async context => {
    if (!requireSocket(context)) return;
    const otherRegistration = await fetch(`${baseUrl}/register`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ redirect_uris: [redirectUri], client_name: "Other Agent" }) });
    const otherClient = await otherRegistration.json() as { client_id: string };
    const wrongClient = await fetch(`${baseUrl}/token`, { method: "POST", body: new URLSearchParams({ grant_type: "refresh_token", client_id: otherClient.client_id, refresh_token: refreshToken }) });
    assert.equal(wrongClient.status, 400);
    const response = await fetch(`${baseUrl}/token`, { method: "POST", body: new URLSearchParams({ grant_type: "refresh_token", client_id: clientId, refresh_token: refreshToken }) });
    assert.equal(response.status, 200);
    const tokens = await response.json() as { access_token: string; refresh_token: string };
    assert.notEqual(tokens.access_token, accessToken); assert.notEqual(tokens.refresh_token, refreshToken);
    const reuse = await fetch(`${baseUrl}/token`, { method: "POST", body: new URLSearchParams({ grant_type: "refresh_token", client_id: clientId, refresh_token: refreshToken }) });
    assert.equal(reuse.status, 400);
    const familyRevoked = await fetch(`${baseUrl}/mcp`, { method: "POST", headers: { authorization: `Bearer ${tokens.access_token}`, "content-type": "application/json" }, body: "{}" });
    assert.equal(familyRevoked.status, 401);
    ({ access_token: accessToken, refresh_token: refreshToken } = await authorizeTokens());
  });

  it("revokes tokens through the RFC 7009 endpoint", async context => {
    if (!requireSocket(context)) return;
    const revoked = await fetch(`${baseUrl}/revoke`, { method: "POST", body: new URLSearchParams({ client_id: clientId, token: refreshToken, token_type_hint: "refresh_token" }) });
    assert.equal(revoked.status, 200);
    const denied = await fetch(`${baseUrl}/mcp`, { method: "POST", headers: { authorization: `Bearer ${accessToken}`, "content-type": "application/json" }, body: "{}" });
    assert.equal(denied.status, 401);
    ({ access_token: accessToken, refresh_token: refreshToken } = await authorizeTokens());
  });

  it("allows refreshed tokens from the session family and hides sessions from other families", async context => {
    if (!requireSocket(context)) return;
    const commonHeaders = { authorization: `Bearer ${accessToken}`, "content-type": "application/json", accept: "application/json, text/event-stream" };
    const unknown = await fetch(`${baseUrl}/mcp`, { method: "POST", headers: { ...commonHeaders, "mcp-session-id": randomUUID() }, body: "{}" });
    assert.equal(unknown.status, 404);
    const initialize = await fetch(`${baseUrl}/mcp`, {
      method: "POST",
      headers: commonHeaders,
      body: JSON.stringify({ jsonrpc: "2.0", id: 1, method: "initialize", params: { protocolVersion: "2025-03-26", capabilities: {}, clientInfo: { name: "binding-test", version: "1.0.0" } } }),
    });
    assert.equal(initialize.status, 200);
    const sessionId = initialize.headers.get("mcp-session-id");
    assert.ok(sessionId);
    const refresh = await fetch(`${baseUrl}/token`, { method: "POST", body: new URLSearchParams({ grant_type: "refresh_token", client_id: clientId, refresh_token: refreshToken }) });
    assert.equal(refresh.status, 200);
    ({ access_token: accessToken, refresh_token: refreshToken } = await refresh.json() as { access_token: string; refresh_token: string });
    const resumed = await fetch(`${baseUrl}/mcp`, {
      method: "POST",
      headers: { ...commonHeaders, authorization: `Bearer ${accessToken}`, "mcp-session-id": sessionId },
      body: JSON.stringify({ jsonrpc: "2.0", method: "notifications/initialized" }),
    });
    assert.equal(resumed.status, 202);

    const otherFamily = await authorizeTokens();
    const mismatch = await fetch(`${baseUrl}/mcp`, { method: "POST", headers: { ...commonHeaders, authorization: `Bearer ${otherFamily.access_token}`, "mcp-session-id": sessionId }, body: "{}" });
    assert.equal(mismatch.status, 404);
  });

  it("rate-limits concurrent wrong passphrases before verification", async context => {
    if (!requireSocket(context)) return;
    const { supplemental, url } = await startSupplementalServer("rate-limit", { passphraseBackoffBaseMs: 50, passphraseBackoffMaxMs: 50 });
    try {
      const registration = await fetch(`${url}/register`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ redirect_uris: [redirectUri] }) }).then(response => response.json()) as { client_id: string };
      const authorize = new URL(`${url}/authorize`);
      authorize.search = new URLSearchParams({ response_type: "code", client_id: registration.client_id, redirect_uri: redirectUri, code_challenge: challenge, code_challenge_method: "S256" }).toString();
      const requestId = (await (await fetch(authorize)).text()).match(/name="request_id" value="([^"]+)"/)?.[1];
      assert.ok(requestId);
      const statuses = await Promise.all(Array.from({ length: 6 }, async () => (await fetch(`${url}/authorize`, { method: "POST", body: new URLSearchParams({ request_id: requestId, passphrase: "wrong" }) })).status));
      assert.equal(statuses.filter(status => status === 403).length, 1);
      assert.equal(statuses.filter(status => status === 429).length, 5);
    } finally { await closeServer(supplemental); }
  });

  it("invalidates a consent request after repeated passphrase failures", async context => {
    if (!requireSocket(context)) return;
    const { supplemental, url } = await startSupplementalServer("consent-failures", { passphraseBackoffBaseMs: 5, passphraseBackoffMaxMs: 5, maxConsentFailures: 2 });
    try {
      const registration = await fetch(`${url}/register`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ redirect_uris: [redirectUri] }) }).then(response => response.json()) as { client_id: string };
      const authorize = new URL(`${url}/authorize`);
      authorize.search = new URLSearchParams({ response_type: "code", client_id: registration.client_id, redirect_uri: redirectUri, code_challenge: challenge, code_challenge_method: "S256" }).toString();
      const requestId = (await (await fetch(authorize)).text()).match(/name="request_id" value="([^"]+)"/)?.[1];
      assert.ok(requestId);
      for (let count = 0; count < 2; count += 1) {
        const denied = await fetch(`${url}/authorize`, { method: "POST", body: new URLSearchParams({ request_id: requestId, passphrase: "wrong" }) });
        assert.equal(denied.status, 403);
        await sleep(10);
      }
      const invalidated = await fetch(`${url}/authorize`, { method: "POST", body: new URLSearchParams({ request_id: requestId, passphrase }) });
      assert.equal(invalidated.status, 400);
    } finally { await closeServer(supplemental); }
  });

  it("returns 401 for an expired access token", async context => {
    if (!requireSocket(context)) return;
    const { supplemental, url } = await startSupplementalServer("expired-token", { oauthStoreOptions: { accessTokenTtlMs: 5 } });
    try {
      const issued = await registerAndAuthorize(url);
      await sleep(15);
      const response = await fetch(`${url}/mcp`, { method: "POST", headers: { authorization: `Bearer ${issued.accessToken}`, "content-type": "application/json" }, body: "{}" });
      assert.equal(response.status, 401);
    } finally { await closeServer(supplemental); }
  });

  it("expires idle MCP sessions", async context => {
    if (!requireSocket(context)) return;
    const { supplemental, url } = await startSupplementalServer("idle-session", { sessionIdleTtlMs: 30, sweepIntervalMs: 5 });
    try {
      const issued = await registerAndAuthorize(url);
      const headers = { authorization: `Bearer ${issued.accessToken}`, "content-type": "application/json", accept: "application/json, text/event-stream" };
      const initialize = await fetch(`${url}/mcp`, { method: "POST", headers, body: JSON.stringify({ jsonrpc: "2.0", id: 1, method: "initialize", params: { protocolVersion: "2025-03-26", capabilities: {}, clientInfo: { name: "idle-test", version: "1.0.0" } } }) });
      assert.equal(initialize.status, 200);
      const sessionId = initialize.headers.get("mcp-session-id");
      assert.ok(sessionId);
      await sleep(60);
      const expired = await fetch(`${url}/mcp`, { method: "POST", headers: { ...headers, "mcp-session-id": sessionId }, body: "{}" });
      assert.equal(expired.status, 404);
    } finally { await closeServer(supplemental); }
  });

  it("initializes MCP and lists the same 125 tools over HTTP", async context => {
    if (!requireSocket(context)) return;
    const client = new Client({ name: "http-test", version: "0.0.0" });
    const transport = new StreamableHTTPClientTransport(new URL(`${baseUrl}/mcp`), { requestInit: { headers: { authorization: `Bearer ${accessToken}` } } });
    await client.connect(transport);
    const { tools } = await client.listTools();
    assert.equal(tools.length, 125);
    await client.close();
  });
});
