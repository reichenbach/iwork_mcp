import { createHash, randomBytes, scrypt as scryptCallback, timingSafeEqual } from "node:crypto";
import { mkdir, readFile, rename, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { promisify } from "node:util";

const scrypt = promisify(scryptCallback);
const SHA256_KEY = /^[a-f0-9]{64}$/;
const PKCE_VERIFIER = /^[A-Za-z0-9._~-]{43,128}$/;

export interface ClientRecord {
  client_id: string;
  redirect_uris: string[];
  client_name?: string;
  token_endpoint_auth_method: "none";
  grant_types: string[];
  response_types: string[];
  client_id_issued_at: number;
}

interface StoredClient extends ClientRecord { lastUsedAt: number; hasIssuedTokens: boolean }
interface AuthorizationRecord { clientId: string; redirectUri: string; codeChallenge: string; expiresAt: number; used: boolean }
interface TokenRecord { clientId: string; kind: "access" | "refresh"; expiresAt: number; familyId: string; status: "active" | "used" | "revoked" }
interface PersistedState { version: 3; clients: Record<string, StoredClient>; authorizationCodes: Record<string, AuthorizationRecord>; tokens: Record<string, TokenRecord> }
interface PassphraseHash { algorithm: "scrypt"; salt: string; hash: string }

export interface OAuthStoreOptions {
  codeTtlMs?: number;
  accessTokenTtlMs?: number;
  refreshTokenTtlMs?: number;
  clientUnusedTtlMs?: number;
  clientInactiveTtlMs?: number;
  maxClients?: number;
  persistDebounceMs?: number;
}

export interface AccessTokenIdentity { clientId: string; tokenHash: string; familyId: string; expiresAt: number }

const DEFAULTS = {
  codeTtlMs: 300_000,
  accessTokenTtlMs: 900_000,
  refreshTokenTtlMs: 2_592_000_000,
  clientUnusedTtlMs: 3_600_000,
  clientInactiveTtlMs: 2_592_000_000,
  maxClients: 100,
  persistDebounceMs: 50,
};

function emptyState(): PersistedState { return { version: 3, clients: {}, authorizationCodes: {}, tokens: {} }; }
function hashSecret(value: string): string { return createHash("sha256").update(value).digest("hex"); }

async function writePrivateJson(path: string, value: unknown): Promise<void> {
  await mkdir(dirname(path), { recursive: true, mode: 0o700 });
  const temp = `${path}.${process.pid}.${randomBytes(6).toString("hex")}.tmp`;
  await writeFile(temp, `${JSON.stringify(value, null, 2)}\n`, { mode: 0o600 });
  await rename(temp, path);
}

async function readJson<T>(path: string): Promise<T | undefined> {
  try { return JSON.parse(await readFile(path, "utf8")) as T; }
  catch (error) { if ((error as NodeJS.ErrnoException).code === "ENOENT") return undefined; throw error; }
}

export async function hashPassphrase(passphrase: string): Promise<PassphraseHash> {
  const salt = randomBytes(16);
  const derived = await scrypt(passphrase, salt, 32) as Buffer;
  return { algorithm: "scrypt", salt: salt.toString("base64url"), hash: derived.toString("base64url") };
}

export async function savePassphrase(configDir: string, passphrase: string): Promise<void> {
  if (passphrase.length < 12) throw new Error("Passphrase must contain at least 12 characters");
  await writePrivateJson(join(configDir, "passphrase.json"), await hashPassphrase(passphrase));
}

export async function verifyPassphrase(configDir: string, candidate: string, environmentPassphrase?: string): Promise<boolean> {
  if (environmentPassphrase !== undefined) return safeEqual(environmentPassphrase, candidate);
  const stored = await readJson<PassphraseHash>(join(configDir, "passphrase.json"));
  if (!stored || stored.algorithm !== "scrypt") return false;
  const expected = Buffer.from(stored.hash, "base64url");
  const actual = await scrypt(candidate, Buffer.from(stored.salt, "base64url"), expected.length) as Buffer;
  return timingSafeEqual(expected, actual);
}

export class OAuthStore {
  private state: PersistedState = emptyState();
  private loadPromise?: Promise<void>;
  private dirty = false;
  private persistTimer?: NodeJS.Timeout;
  private writeQueue: Promise<void> = Promise.resolve();
  private readonly settings: typeof DEFAULTS;

  constructor(private readonly statePath: string, options: OAuthStoreOptions = {}) {
    this.settings = { ...DEFAULTS, ...options };
  }

  async load(): Promise<void> {
    if (!this.loadPromise) this.loadPromise = this.loadState();
    await this.loadPromise;
  }

  private async loadState(): Promise<void> {
    const raw = await readJson<Partial<PersistedState> & { version?: number }>(this.statePath);
    if (!raw) return;
    const now = Date.now();
    const authorizationCodes: Record<string, AuthorizationRecord> = {};
    for (const [key, record] of Object.entries(raw.authorizationCodes ?? {})) authorizationCodes[SHA256_KEY.test(key) ? key : hashSecret(key)] = record;
    const tokens: Record<string, TokenRecord> = {};
    for (const [key, value] of Object.entries(raw.tokens ?? {})) {
      const record = value as Partial<TokenRecord> & Pick<TokenRecord, "clientId" | "kind" | "expiresAt">;
      tokens[SHA256_KEY.test(key) ? key : hashSecret(key)] = {
        clientId: record.clientId,
        kind: record.kind,
        expiresAt: record.expiresAt,
        familyId: record.familyId ?? randomBytes(18).toString("base64url"),
        status: record.status ?? "active",
      };
    }
    const clientsWithTokens = new Set(Object.values(tokens).map(record => record.clientId));
    const clients: Record<string, StoredClient> = {};
    for (const [id, client] of Object.entries(raw.clients ?? {})) {
      clients[id] = {
        ...client,
        lastUsedAt: (client as StoredClient).lastUsedAt ?? client.client_id_issued_at * 1000,
        hasIssuedTokens: (client as StoredClient).hasIssuedTokens ?? clientsWithTokens.has(id),
      };
    }
    this.state = { version: 3, clients, authorizationCodes, tokens };
    const changed = raw.version !== 3 || this.prune(now);
    if (changed) this.markDirty();
  }

  private prune(now = Date.now()): boolean {
    let changed = false;
    for (const [codeHash, record] of Object.entries(this.state.authorizationCodes)) {
      if (record.expiresAt <= now) { delete this.state.authorizationCodes[codeHash]; changed = true; }
    }
    for (const [tokenHash, record] of Object.entries(this.state.tokens)) {
      if (record.expiresAt <= now) { delete this.state.tokens[tokenHash]; changed = true; }
    }
    const referencedClients = new Set([
      ...Object.values(this.state.authorizationCodes).map(record => record.clientId),
      ...Object.values(this.state.tokens).map(record => record.clientId),
    ]);
    for (const [clientId, client] of Object.entries(this.state.clients)) {
      const ttl = client.hasIssuedTokens ? this.settings.clientInactiveTtlMs : this.settings.clientUnusedTtlMs;
      if (client.lastUsedAt + ttl > now) continue;
      if (client.hasIssuedTokens && referencedClients.has(clientId)) continue;
      this.deleteClient(clientId);
      changed = true;
    }
    return changed;
  }

  async sweep(): Promise<void> {
    await this.load();
    if (this.prune()) this.markDirty();
  }

  private markDirty(): void {
    this.dirty = true;
    if (this.persistTimer) return;
    this.persistTimer = setTimeout(() => { this.persistTimer = undefined; void this.flush(); }, this.settings.persistDebounceMs);
    this.persistTimer.unref();
  }

  async flush(): Promise<void> {
    await this.load();
    if (this.persistTimer) { clearTimeout(this.persistTimer); this.persistTimer = undefined; }
    if (!this.dirty) return this.writeQueue;
    this.dirty = false;
    this.prune();
    const snapshot = structuredClone(this.state);
    this.writeQueue = this.writeQueue.then(() => writePrivateJson(this.statePath, snapshot));
    await this.writeQueue;
    if (this.dirty) this.markDirty();
  }

  async close(): Promise<void> { await this.flush(); }

  async register(metadata: { redirect_uris: string[]; client_name?: string }): Promise<ClientRecord | undefined> {
    await this.load();
    if (this.prune()) this.markDirty();
    while (Object.keys(this.state.clients).length >= this.settings.maxClients) {
      const oldestUnused = Object.values(this.state.clients)
        .filter(client => !client.hasIssuedTokens)
        .sort((left, right) => left.lastUsedAt - right.lastUsedAt || left.client_id_issued_at - right.client_id_issued_at)[0];
      if (!oldestUnused) return undefined;
      this.deleteClient(oldestUnused.client_id);
      this.markDirty();
    }
    const now = Date.now();
    const client: StoredClient = {
      client_id: randomBytes(24).toString("base64url"),
      redirect_uris: metadata.redirect_uris,
      client_name: metadata.client_name,
      token_endpoint_auth_method: "none",
      grant_types: ["authorization_code", "refresh_token"],
      response_types: ["code"],
      client_id_issued_at: Math.floor(now / 1000),
      lastUsedAt: now,
      hasIssuedTokens: false,
    };
    this.state.clients[client.client_id] = client;
    this.markDirty();
    return this.publicClient(client);
  }

  async client(clientId: string): Promise<ClientRecord | undefined> {
    await this.load();
    const client = this.state.clients[clientId];
    if (!client) return undefined;
    return this.publicClient(client);
  }

  private publicClient(client: StoredClient): ClientRecord {
    const { lastUsedAt: _lastUsedAt, hasIssuedTokens: _hasIssuedTokens, ...publicRecord } = client;
    return publicRecord;
  }

  private deleteClient(clientId: string): void {
    delete this.state.clients[clientId];
    for (const [codeHash, record] of Object.entries(this.state.authorizationCodes)) {
      if (record.clientId === clientId) delete this.state.authorizationCodes[codeHash];
    }
    for (const [tokenHash, record] of Object.entries(this.state.tokens)) {
      if (record.clientId === clientId) delete this.state.tokens[tokenHash];
    }
  }

  async issueCode(clientId: string, redirectUri: string, codeChallenge: string): Promise<string> {
    await this.load();
    const client = this.state.clients[clientId];
    if (client) client.lastUsedAt = Date.now();
    const code = randomBytes(32).toString("base64url");
    this.state.authorizationCodes[hashSecret(code)] = { clientId, redirectUri, codeChallenge, expiresAt: Date.now() + this.settings.codeTtlMs, used: false };
    this.markDirty();
    return code;
  }

  async exchangeCode(code: string, clientId: string, redirectUri: string, verifier: string): Promise<{ access_token: string; refresh_token: string } | undefined> {
    await this.load();
    if (!PKCE_VERIFIER.test(verifier)) return undefined;
    const record = this.state.authorizationCodes[hashSecret(code)];
    const challenge = createHash("sha256").update(verifier).digest("base64url");
    if (!record || record.used || record.expiresAt <= Date.now() || record.clientId !== clientId || record.redirectUri !== redirectUri || !safeEqual(record.codeChallenge, challenge)) return undefined;
    record.used = true;
    return this.issueTokens(clientId);
  }

  async refresh(refreshToken: string, clientId: string): Promise<{ access_token: string; refresh_token: string } | undefined> {
    await this.load();
    const record = this.state.tokens[hashSecret(refreshToken)];
    if (!record || record.kind !== "refresh" || record.expiresAt <= Date.now() || record.clientId !== clientId) return undefined;
    if (record.status !== "active") {
      this.revokeFamily(record.familyId);
      this.markDirty();
      return undefined;
    }
    record.status = "used";
    return this.issueTokens(clientId, record.familyId);
  }

  private async issueTokens(clientId: string, familyId = randomBytes(18).toString("base64url")): Promise<{ access_token: string; refresh_token: string }> {
    const client = this.state.clients[clientId];
    if (client) {
      client.lastUsedAt = Date.now();
      client.hasIssuedTokens = true;
    }
    const access_token = randomBytes(32).toString("base64url");
    const refresh_token = randomBytes(32).toString("base64url");
    this.state.tokens[hashSecret(access_token)] = { clientId, kind: "access", expiresAt: Date.now() + this.settings.accessTokenTtlMs, familyId, status: "active" };
    this.state.tokens[hashSecret(refresh_token)] = { clientId, kind: "refresh", expiresAt: Date.now() + this.settings.refreshTokenTtlMs, familyId, status: "active" };
    this.markDirty();
    return { access_token, refresh_token };
  }

  private revokeFamily(familyId: string): void {
    for (const record of Object.values(this.state.tokens)) if (record.familyId === familyId) record.status = "revoked";
  }

  async revoke(token: string, clientId: string): Promise<void> {
    await this.load();
    const record = this.state.tokens[hashSecret(token)];
    if (!record || record.clientId !== clientId) return;
    this.revokeFamily(record.familyId);
    this.markDirty();
  }

  async accessToken(token: string): Promise<AccessTokenIdentity | undefined> {
    await this.load();
    const tokenHash = hashSecret(token);
    const record = this.state.tokens[tokenHash];
    if (!record || record.kind !== "access" || record.status !== "active" || record.expiresAt <= Date.now()) return undefined;
    return { clientId: record.clientId, tokenHash, familyId: record.familyId, expiresAt: record.expiresAt };
  }

  async validAccessToken(token: string): Promise<boolean> { return Boolean(await this.accessToken(token)); }
}

export function safeEqual(a: string, b: string): boolean {
  const left = Buffer.from(a); const right = Buffer.from(b);
  return left.length === right.length && timingSafeEqual(left, right);
}
