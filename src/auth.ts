import { createHash, randomBytes, scrypt as scryptCallback, timingSafeEqual } from "node:crypto";
import { mkdir, readFile, rename, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { promisify } from "node:util";

const scrypt = promisify(scryptCallback);

export interface ClientRecord {
  client_id: string;
  redirect_uris: string[];
  client_name?: string;
  token_endpoint_auth_method: "none";
  grant_types: string[];
  response_types: string[];
  client_id_issued_at: number;
}
interface AuthorizationRecord { clientId: string; redirectUri: string; codeChallenge: string; expiresAt: number; used: boolean }
interface TokenRecord { clientId: string; kind: "access" | "refresh"; expiresAt: number }
interface PersistedState { clients: Record<string, ClientRecord>; authorizationCodes: Record<string, AuthorizationRecord>; tokens: Record<string, TokenRecord> }
interface PassphraseHash { algorithm: "scrypt"; salt: string; hash: string }

function emptyState(): PersistedState { return { clients: {}, authorizationCodes: {}, tokens: {} }; }

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
  private loaded = false;
  constructor(private readonly statePath: string) {}
  async load(): Promise<void> { if (!this.loaded) { this.state = await readJson<PersistedState>(this.statePath) ?? emptyState(); this.loaded = true; this.prune(); } }
  private prune(): void {
    const now = Date.now();
    for (const [code, record] of Object.entries(this.state.authorizationCodes)) if (record.expiresAt <= now || record.used) delete this.state.authorizationCodes[code];
    for (const [token, record] of Object.entries(this.state.tokens)) if (record.expiresAt <= now) delete this.state.tokens[token];
  }
  private async persist(): Promise<void> { this.prune(); await writePrivateJson(this.statePath, this.state); }
  async register(metadata: { redirect_uris: string[]; client_name?: string }): Promise<ClientRecord> {
    await this.load();
    const client: ClientRecord = { client_id: randomBytes(24).toString("base64url"), redirect_uris: metadata.redirect_uris, client_name: metadata.client_name, token_endpoint_auth_method: "none", grant_types: ["authorization_code", "refresh_token"], response_types: ["code"], client_id_issued_at: Math.floor(Date.now() / 1000) };
    this.state.clients[client.client_id] = client; await this.persist(); return client;
  }
  async client(clientId: string): Promise<ClientRecord | undefined> { await this.load(); return this.state.clients[clientId]; }
  async issueCode(clientId: string, redirectUri: string, codeChallenge: string): Promise<string> {
    await this.load(); const code = randomBytes(32).toString("base64url");
    this.state.authorizationCodes[code] = { clientId, redirectUri, codeChallenge, expiresAt: Date.now() + 300_000, used: false }; await this.persist(); return code;
  }
  async exchangeCode(code: string, clientId: string, redirectUri: string, verifier: string): Promise<{ access_token: string; refresh_token: string } | undefined> {
    await this.load(); const record = this.state.authorizationCodes[code]; const challenge = createHash("sha256").update(verifier).digest("base64url");
    if (!record || record.used || record.expiresAt <= Date.now() || record.clientId !== clientId || record.redirectUri !== redirectUri || !safeEqual(record.codeChallenge, challenge)) return undefined;
    record.used = true; return this.issueTokens(clientId);
  }
  async refresh(refreshToken: string, clientId: string): Promise<{ access_token: string; refresh_token: string } | undefined> {
    await this.load(); const record = this.state.tokens[refreshToken];
    if (!record || record.kind !== "refresh" || record.expiresAt <= Date.now() || record.clientId !== clientId) return undefined;
    delete this.state.tokens[refreshToken]; return this.issueTokens(clientId);
  }
  private async issueTokens(clientId: string): Promise<{ access_token: string; refresh_token: string }> {
    const access_token = randomBytes(32).toString("base64url"); const refresh_token = randomBytes(32).toString("base64url");
    this.state.tokens[access_token] = { clientId, kind: "access", expiresAt: Date.now() + 900_000 };
    this.state.tokens[refresh_token] = { clientId, kind: "refresh", expiresAt: Date.now() + 2_592_000_000 };
    await this.persist(); return { access_token, refresh_token };
  }
  async validAccessToken(token: string): Promise<boolean> {
    await this.load(); const found = Object.entries(this.state.tokens).find(([candidate]) => safeEqual(candidate, token));
    return Boolean(found && found[1].kind === "access" && found[1].expiresAt > Date.now());
  }
}

export function safeEqual(a: string, b: string): boolean {
  const left = Buffer.from(a); const right = Buffer.from(b);
  return left.length === right.length && timingSafeEqual(left, right);
}
