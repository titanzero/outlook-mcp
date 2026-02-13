<div align="center">

<a href="https://glama.ai/mcp/servers/@titanzero/outlook-mcp">
  <img width="380" height="200" src="https://glama.ai/mcp/servers/@titanzero/outlook-mcp/badge" alt="Outlook MCP on Glama" />
</a>

# Outlook MCP Server

**Let AI manage your Outlook inbox, calendar, and rules — through natural language.**

Built on [Model Context Protocol](https://modelcontextprotocol.io) · Powered by [Microsoft Graph API](https://learn.microsoft.com/en-us/graph/overview)

[![Node.js](https://img.shields.io/badge/Node.js-%3E%3D14-339933?logo=nodedotjs&logoColor=white)](https://nodejs.org)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![MCP SDK](https://img.shields.io/badge/MCP_SDK-1.1-blueviolet)](https://www.npmjs.com/package/@modelcontextprotocol/sdk)

</div>

---

## What is this?

This MCP server turns Claude into a full-featured **Outlook assistant**. Instead of clicking through the Outlook UI, just ask Claude:

> *"Show me unread emails from this week"*
> *"Schedule a meeting with Alice tomorrow at 3pm"*
> *"Create a rule to move all GitHub notifications to a folder"*

Claude handles authentication, API calls, pagination, filtering — everything. You just talk.

---

## Capabilities

| Area | What Claude can do |
|---|---|
| **Email** | List, search, read (preview or full body), send, mark read/unread |
| **Calendar** | List upcoming events, create, accept, decline, cancel, delete |
| **Folders** | List folder hierarchy, create folders, move emails between folders |
| **Rules** | List inbox rules, create new rules, change rule execution order |
| **Auth** | OAuth 2.0 with automatic token refresh — authenticate once, use forever |

---

## Quick Start

```bash
# 1. Clone and install
git clone https://github.com/titanzero/outlook-mcp.git
cd outlook-mcp
npm install

# 2. Configure (see Azure Setup below)
cp .env.example .env
# Edit .env with your Azure credentials

# 3. Start the OAuth server and authenticate
npm run auth-server

# 4. Add to Claude Desktop config and start using!
```

---

## Azure App Setup

You need an Azure app registration to connect to Microsoft Graph.

<details>
<summary><strong>1. Register the App</strong></summary>

1. Open [Azure Portal → App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)
2. Click **New registration**
3. Name: `Outlook MCP Server` (or anything you like)
4. Account type: **Accounts in any organizational directory and personal Microsoft accounts**
5. Redirect URI: **Web** → `http://localhost:3333/auth/callback`
6. Click **Register**
7. Copy the **Application (client) ID** → this is your `OUTLOOK_CLIENT_ID`

</details>

<details>
<summary><strong>2. Set API Permissions</strong></summary>

Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**, then add:

- `offline_access`
- `User.Read`
- `Mail.Read`
- `Mail.Send`
- `Calendars.Read`
- `Calendars.ReadWrite`
- `Contacts.Read`

</details>

<details>
<summary><strong>3. Generate Client Secret</strong></summary>

1. Go to **Certificates & secrets** → **Client secrets** → **New client secret**
2. Set description and longest expiration
3. **Copy the secret VALUE** (not the Secret ID!)
4. This is your `OUTLOOK_CLIENT_SECRET`

</details>

---

## Configuration

### Environment Variables

Create `.env` in the project root:

```bash
OUTLOOK_CLIENT_ID=your-application-client-id
OUTLOOK_CLIENT_SECRET=your-client-secret-VALUE
```

> **Important:** Always use the secret **VALUE** from Azure, not the Secret ID.

### Claude Desktop

Add to your Claude Desktop config (`claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "outlook-assistant": {
      "command": "node",
      "args": ["/absolute/path/to/outlook-mcp/index.js"],
      "env": {
        "OUTLOOK_CLIENT_ID": "your-client-id",
        "OUTLOOK_CLIENT_SECRET": "your-client-secret"
      }
    }
  }
}
```

### Advanced

Edit `config.js` to customize server name, timezone, pagination limits, field selections, and response format (`toon` or `text`).

---

## Authentication Flow

```
You ──ask Claude──▸ "authenticate"
                        │
Claude returns URL ◂────┘
                        │
You open URL in browser ▸ Microsoft login ▸ Grant permissions
                                                    │
                        ┌───────────────────────────┘
                        ▼
              OAuth callback on localhost:3333
              Tokens saved to ~/.outlook-mcp-tokens.json
              ✔ Auto-refresh — no re-auth needed
```

**Step 1** — Start the auth server (must be running before authenticating):

```bash
npm run auth-server
```

**Step 2** — Ask Claude to `authenticate`, open the URL, sign in, done.

Tokens persist in `~/.outlook-mcp-tokens.json` and refresh automatically.

---

## Project Structure

```
index.js                  ── MCP entry point
config.js                 ── centralized constants & settings
outlook-auth-server.js    ── standalone OAuth server

auth/                     ── authentication & token management
email/                    ── list, search, read, send, mark-as-read
calendar/                 ── list, create, accept, decline, cancel, delete
folder/                   ── list, create, move
rules/                    ── list, create, edit-rule-sequence

utils/
├── graph-client.js       ── Graph SDK wrapper with pagination
├── response-formatter.js ── TOON / plain-text output toggle
└── response-helpers.js   ── error detection & MCP response builders

scripts/                  ── CLI utilities & debug helpers
```

---

## Available Commands

| Command | Description |
|---|---|
| `npm install` | Install dependencies |
| `npm start` | Start the MCP server (stdio) |
| `npm run auth-server` | Start OAuth server on port 3333 |
| `npm run inspect` | Launch MCP Inspector for interactive testing |
| `npm test` | Run Jest test suite |
| `npm run debug` | Print env vars and start server |
| `npx kill-port 3333` | Free port 3333 if occupied |

---

## Troubleshooting

<details>
<summary><strong>Cannot find module '@modelcontextprotocol/sdk'</strong></summary>

Run `npm install` first.

</details>

<details>
<summary><strong>EADDRINUSE: port 3333 already in use</strong></summary>

```bash
npx kill-port 3333
npm run auth-server
```

</details>

<details>
<summary><strong>Invalid client secret (AADSTS7000215)</strong></summary>

You're using the Secret **ID** instead of the Secret **Value**. Go to Azure Portal → Certificates & secrets → copy the **Value** column.

</details>

<details>
<summary><strong>Auth URL doesn't load / "site can't be reached"</strong></summary>

The auth server isn't running. Start it first with `npm run auth-server`, then retry.

</details>

<details>
<summary><strong>"Authentication required" after setup</strong></summary>

Token may be expired or corrupted. Delete `~/.outlook-mcp-tokens.json` and re-authenticate.

</details>

<details>
<summary><strong>Server doesn't start in Claude Desktop</strong></summary>

1. Verify the absolute path to `index.js` in your Claude Desktop config
2. Ensure `OUTLOOK_CLIENT_ID` and `OUTLOOK_CLIENT_SECRET` are set
3. Restart Claude Desktop after config changes

</details>

---

## Extending the Server

Adding a new tool is straightforward:

1. Create a handler file in the appropriate module directory
2. Export `{ name, description, inputSchema, handler }`
3. Add it to the module's `index.js` exports
4. It's automatically registered via `index.js` at the root

See `.cursor/rules/new-tool.mdc` for the full checklist.

---

## License

[MIT](LICENSE)
