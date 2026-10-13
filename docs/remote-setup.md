# Remote connector setup

This guide exposes the authenticated Streamable HTTP endpoint through a free named Cloudflare Tunnel. Replace every uppercase placeholder before running commands. The MCP server remains bound to `127.0.0.1`; only the tunnel accepts public traffic.

## 1. Set the owner passphrase

```bash
iwork-mcp --setup-passphrase
```

The passphrase is scrypt-hashed in the current user's configuration directory. OAuth clients and tokens are stored in the current user's data directory with owner-only permissions. Alternatively, provide `IWORK_MCP_PASSPHRASE` in the process environment; avoid shell history and source-control files.

## 2. Create a named tunnel

Install `cloudflared`, authenticate it to your own account, and create the tunnel:

```bash
cloudflared tunnel login
cloudflared tunnel create YOUR-TUNNEL-NAME
cloudflared tunnel route dns YOUR-TUNNEL-NAME YOUR-TUNNEL-HOST
```

Create the tunnel configuration in cloudflared's user configuration directory:

```yaml
tunnel: YOUR-TUNNEL-ID
credentials-file: YOUR-CLOUDFLARED-CONFIG-DIRECTORY/YOUR-TUNNEL-ID.json
ingress:
  - hostname: YOUR-TUNNEL-HOST
    service: http://127.0.0.1:8787
  - service: http_status:404
```

## 3. Start both processes

```bash
IWORK_MCP_PUBLIC_URL=https://YOUR-TUNNEL-HOST iwork-mcp --http --port 8787 --host 127.0.0.1
cloudflared tunnel run YOUR-TUNNEL-NAME
```

The public MCP URL is `https://YOUR-TUNNEL-HOST/mcp`. The issuer must be the externally visible origin, so `IWORK_MCP_PUBLIC_URL` must exactly match it and must not include `/mcp`.

To restrict dynamic clients to known callback hosts, set a comma-separated allowlist. Entries may be hostnames, `host:port` pairs, or origins:

```sh
IWORK_MCP_REDIRECT_HOST_ALLOWLIST=127.0.0.1,localhost,client.example.com
```

For launch-at-login, copy [the plist template](../packaging/com.example.iwork-mcp.plist), replace its placeholders, and place the result in the current user's LaunchAgents directory. Keep secrets out of the plist; use the stored passphrase created above.

## Security notes

- Do not bind the MCP process directly to a public interface.
- Use an HTTPS hostname you control.
- Protect the local account and tunnel credentials.
- Re-run `--setup-passphrase` to replace the owner passphrase. Existing access and refresh tokens remain valid until they expire; remove the OAuth state file from the per-user data directory to invalidate all sessions.
