# iwork-mcp

An agent-neutral MCP server for automating Numbers, Pages, and Keynote on macOS. It exposes 117 tools over local stdio or authenticated Streamable HTTP.

![Numbers dashboard](screenshots/numbers-dashboard.jpg)

## Capabilities

- Numbers: create spreadsheets, read and write cells and ranges, use formulas, format tables, add charts, and export to PDF, spreadsheet, or CSV formats.
- Pages: create and format documents, edit paragraphs, insert images and tables, and export to PDF, document, EPUB, or text formats.
- Keynote: create and edit presentations, lay out slide items, add presenter notes and transitions, and export to PDF, presentation, HTML, or image formats.

Requirements: macOS 13 or newer, current iWork applications, and Node.js 18 or newer. macOS asks for Automation permission on first use.

## Local stdio setup

stdio is the default transport. Each client starts its own server process.

### Grok Bot

Add an MCP command server named `iwork` with command `npx` and arguments `-y`, `iwork-mcp`.

### Cursor

Use the included plugin listing, or add this to the client's MCP configuration:

```json
{
  "mcpServers": {
    "iwork": {
      "command": "npx",
      "args": ["-y", "iwork-mcp"]
    }
  }
}
```

### Claude

For Claude Desktop:

```bash
npx iwork-mcp install
```

For Claude Code:

```bash
claude mcp add iwork -- npx -y iwork-mcp
```

### Codex

```bash
codex mcp add iwork -- npx -y iwork-mcp
```

## Remote HTTP setup

> **Runs on your own Mac.** iWork (Pages, Numbers, Keynote) is macOS-only, so this server must run on the Mac that has iWork installed. The Cloudflare Tunnel only exposes that Mac's server to remote clients; it cannot run in the cloud.

First configure an owner passphrase and start the server. It binds only to loopback by default:

```bash
iwork-mcp --setup-passphrase
IWORK_MCP_PUBLIC_URL=https://YOUR-TUNNEL-HOST iwork-mcp --http --port 8787
```

The endpoint is `https://YOUR-TUNNEL-HOST/mcp`. It uses OAuth 2.1 dynamic client registration, authorization code with required PKCE S256, rotating refresh tokens, and short-lived opaque access tokens. Every MCP request requires a bearer token.

Optionally set `IWORK_MCP_REDIRECT_HOST_ALLOWLIST` to a comma-separated list of permitted OAuth callback hosts, `host:port` pairs, or origins.

For a free named Cloudflare Tunnel and the optional launch-at-login template, follow [Remote connector setup](docs/remote-setup.md).

### Grok Bot, Cursor, Claude, and Codex

For any client that accepts a remote MCP URL, add:

```json
{
  "mcpServers": {
    "iwork": {
      "url": "https://YOUR-TUNNEL-HOST/mcp"
    }
  }
}
```

The client discovers the OAuth metadata, registers itself, opens the local consent page, and obtains tokens. If a client uses a different configuration shape, enter the same URL in its remote connector UI.

## Example requests

> Build a monthly budget in Numbers with planned and actual income, categorized expenses, variance formulas, and a net savings row.

> Create a concise Keynote review deck with six slides, a consistent layout, and presenter notes.

> Draft a formatted Pages project brief with objectives, milestones, risks, and next steps.

## How it works

For local use, an MCP client communicates with the server over stdio. For remote use, the official SDK's Streamable HTTP transport serves `/mcp` behind OAuth. Tool registration is shared, so both transports expose the same tool set. Tool calls run JavaScript for Automation through the system scripting process and control the installed iWork applications.

## Development

```bash
npm install
npm run build
npm test
```

`npm test` runs unit and remote OAuth/HTTP tests without opening iWork applications. The HTTP suite skips with a clear reason when a sandbox prohibits socket binding. Integration tests require the applications and skip when they are unavailable:

```bash
npm run test:integration
npm run test:all
```

### Reviewer HTTP and OAuth test

Use temporary per-user directories or remove them after testing. Start the server:

```bash
IWORK_MCP_PASSPHRASE='YOUR-LONG-TEST-PASSPHRASE' \
IWORK_MCP_PUBLIC_URL=http://127.0.0.1:8787 \
iwork-mcp --http --host 127.0.0.1 --port 8787
```

In a second terminal, inspect metadata and register a loopback client:

```bash
curl http://127.0.0.1:8787/.well-known/oauth-protected-resource
curl http://127.0.0.1:8787/.well-known/oauth-authorization-server
curl -sS -X POST http://127.0.0.1:8787/register \
  -H 'content-type: application/json' \
  -d '{"redirect_uris":["http://127.0.0.1/callback"],"client_name":"Review Client","token_endpoint_auth_method":"none"}'
```

Create a PKCE verifier and challenge, replacing `YOUR-CLIENT-ID` with the registration result:

```bash
VERIFIER='review-verifier-0123456789012345678901234567890123456789'
CHALLENGE=$(printf %s "$VERIFIER" | openssl dgst -sha256 -binary | openssl base64 -A | tr '+/' '-_' | tr -d '=')
open "http://127.0.0.1:8787/authorize?response_type=code&client_id=YOUR-CLIENT-ID&redirect_uri=http%3A%2F%2F127.0.0.1%2Fcallback&code_challenge=$CHALLENGE&code_challenge_method=S256&state=review"
```

Enter the owner passphrase. The browser redirects to the loopback callback; copy the `code` query value from its address. Exchange it and then initialize MCP:

```bash
curl -sS -X POST http://127.0.0.1:8787/token \
  -d grant_type=authorization_code \
  -d client_id=YOUR-CLIENT-ID \
  -d redirect_uri=http://127.0.0.1/callback \
  -d code=YOUR-AUTHORIZATION-CODE \
  -d code_verifier="$VERIFIER"

curl -i -X POST http://127.0.0.1:8787/mcp \
  -H 'authorization: Bearer YOUR-ACCESS-TOKEN' \
  -H 'content-type: application/json' \
  -H 'accept: application/json, text/event-stream' \
  -d '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-06-18","capabilities":{},"clientInfo":{"name":"review","version":"1"}}}'
```

Copy the `mcp-session-id` response header, send an `notifications/initialized` notification, then call `tools/list` with the same bearer token and session header. The automated test performs this complete sequence and verifies that HTTP returns the same 117 tools as stdio.

## Security and limitations

- Known limitation: `/register` is rate-limited per remote address. Behind `cloudflared` every request arrives from loopback, so all clients share one bucket; a flood of registrations can temporarily block new client setup (existing tokens are unaffected).

The HTTP server defaults to `127.0.0.1`. Keep it on loopback behind an HTTPS tunnel, set `IWORK_MCP_PUBLIC_URL` to the exact external origin, use a strong owner passphrase, and never store secrets in the repository. OAuth state is stored in the current user's data directory with mode `0600`; the passphrase hash is stored separately in the current user's configuration directory.

iWork's scripting dictionaries determine what can be automated. Some application features are unavailable or require macOS Accessibility permission.
