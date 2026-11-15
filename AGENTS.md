# AGENTS.md

## Commands

- `npm install` — **always run first**
- `npm start` — start MCP server (stdio transport)
- `npm run auth-server` — OAuth server on port 3333 (must be running before auth)
- `npm run inspect` — MCP Inspector for interactive testing
- `npm test` — Jest tests
- `npm run debug` — debug: print env vars and start MCP server
- `npm run find-folders` — find folder IDs via Graph API
- `npm run move-github` — move GitHub emails to subfolder
- `npm run create-rules` — create inbox rules for GitHub notifications
- `npx kill-port 3333` — free port 3333 if occupied

## Project Structure

```
index.js                  — MCP entry point; registers all tools, resources, prompts
config.js                 — all constants: field selections, pagination, auth, timezone, response format
outlook-auth-server.js    — standalone Express OAuth server

auth/
  index.js                — exports `registerAuthTools`, `ensureAuthenticated`, `tokenManager`
  tools.js                — auth tool definitions (`about`, `authenticate`, `check-auth-status`)
  token-manager.js        — token load/save/refresh logic
  oauth-server.js         — OAuth callback/exchange helpers

email/                    — list, search, read, send, reply, forward, delete,
                          — mark-as-read, attachments, drafts, flag, categorize
calendar/                 — list calendars & events, create/update/delete calendars,
                          — create/update/delete events, accept/decline, free-busy
folder/                   — list, create, rename, delete, move emails
contacts/                 — list, create (basic/full), update, delete
mailbox/                  — settings, out-of-office, user profile
tasks/                    — task lists and tasks (Microsoft To Do)
rules/                    — list, create, update, delete, edit-sequence

prompts/                  — prompt templates for common workflows
completion/               — autocompletion handlers

utils/
  graph-client.js         — Graph SDK wrapper: `getGraphClient()`, `graphGetPaginated()`
  response-formatter.js   — TOON / plain-text toggle (reads `config.RESPONSE_FORMAT`)
  response-helpers.js     — `isAuthError()`, `makeErrorResponse()`, `makeResponse()`

scripts/                  — CLI & debug one-off scripts (not imported by app modules)
  debug-env.js            — debug: print env vars and start MCP server
  test-config.js          — debug: verify config and token path
  find-folder-ids.js      — CLI: find folder IDs via Graph API
  move-github-emails.js   — CLI: move GitHub emails to subfolder
  create-notifications-rule.js — CLI: create inbox rules for GitHub
  backup-logs.sh          — shell: backup MCP client logs
  test-direct.sh          — shell: test server directly via nc
  test-modular-server.sh  — shell: test server via MCP Inspector
```

## Key Patterns

- **CommonJS** (`require` / `module.exports`) — no ESM
- **Module pattern**: each module `index.js` exports `register{Module}Tools(server)` function plus handlers
- **Tool shape**: `{ name, description, inputSchema, handler }` with Zod schemas and annotations
- **Resource pattern**: `server.resource(name, uriOrTemplate, metadata, handler)` with ResourceTemplate for completions
- **Prompt pattern**: `server.prompt(name, description, argsSchema, handler)`
- **Constants**: keep field selections, defaults, limits, and URLs centralized in `config.js`
- **Logging**: use `console.error()` for diagnostics (stdout is reserved for MCP protocol traffic)

## Graph and Response Flow

- `getGraphClient()` authenticates internally via `ensureAuthenticated()`; handlers should not run a separate auth check.
- Use `graphGetPaginated()` for list endpoints that may return `@odata.nextLink`.
- Data handlers should build a structured payload plus text fallback, then call `formatResponse(structured, textFallback)`.
- Return `makeResponse(text)` for success; in `catch`, use `isAuthError(error)` and `makeErrorResponse(...)`.

## Tool Annotations

All tools include annotations for client optimization:

- `readOnlyHint: true` — tool doesn't modify state
- `destructiveHint: true` — tool deletes/destroys data
- `idempotentHint: true` — safe to call multiple times
- `openWorldHint: true` — interacts with external systems

## Resource Templates with Completions

| Template | URI | Completion |
|---|---|---|
| `outlook-folder` | `outlook://folders/{folderName}` | Folder names/paths |
| `outlook-contact` | `outlook://contacts/{contactId}` | Contact emails/names |
| `outlook-calendar` | `outlook://calendars/{calendarId}` | Calendar names |
| `outlook-event` | `outlook://events/{eventId}` | Event subjects |
| `outlook-task-list` | `outlook://tasks/lists/{listId}` | Task list names |
| `outlook-rule` | `outlook://rules/{ruleId}` | Rule names |
| `outlook-email` | `outlook://emails/{emailId}` | Email subjects |

## Prompts

| Prompt | Description |
|---|---|
| `summarize-inbox` | Summarize recent emails with themes, action items |
| `schedule-meeting` | Check availability and draft meeting invite |
| `triage-email` | Analyze email and suggest actions |
| `daily-briefing` | Calendar, tasks, important emails for the day |
| `weekly-review` | Weekly summary of meetings, emails, tasks |
| `follow-up-draft` | Create follow-up email draft |
| `enrich-contact` | Enrich contact from email/calendar history |

## Config Values (config.js)

| Constant | Value |
|---|---|
| `DEFAULT_PAGE_SIZE` | 50 |
| `MAX_RESULT_COUNT` | 1000 |
| `DEFAULT_TIMEZONE` | "Central European Standard Time" |
| `RESPONSE_FORMAT` | `process.env.OUTLOOK_RESPONSE_FORMAT \|\| 'toon'` |
| `EMAIL_SELECT_FIELDS` | `id,subject,from,receivedDateTime,isRead` |
| `EMAIL_DETAIL_FIELDS` | bodyPreview-based (no full body) |
| `EMAIL_FULL_BODY_FIELDS` | includes `body` field (used when `fullBody=true`) |

## Environment Variables

- `.env` file: `OUTLOOK_CLIENT_ID`, `OUTLOOK_CLIENT_SECRET`
- MCP client config: `OUTLOOK_CLIENT_ID`, `OUTLOOK_CLIENT_SECRET` (used by MCP server)
- `OUTLOOK_RESPONSE_FORMAT` — `'toon'` (default) or `'text'`
- **Important**: use Azure secret **VALUE**, not the Secret ID
- Token storage: `~/.outlook-mcp-tokens.json` (via `config.getTokenStorePath()`)

## Auth Flow

1. Register Azure app with permissions: Mail.Read, Mail.ReadWrite, Mail.Send, Calendars.Read, Calendars.ReadWrite, etc.
2. `npm run auth-server` → opens port 3333
3. Use `authenticate` tool → returns OAuth URL
4. Complete browser login → tokens saved to `~/.outlook-mcp-tokens.json`
5. Tokens auto-refresh via `ensureAuthenticated()`

## Testing

- Framework: Jest
- Tests in `test/` mirroring source structure (`test/email/list.test.js`)
- Mock `utils/graph-client` and module dependencies with `jest.mock()`
- Set `process.env.OUTLOOK_RESPONSE_FORMAT = 'text'` at top of test files to get plain-text output for assertions

## MCP Client Configuration

### Generic (any MCP-compatible client)

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

### Claude Desktop

See `CLAUDE.md` for Claude-specific configuration.

### VS Code (with MCP extension)

See `VSCODE.md` for VS Code-specific configuration.

## Extending the Server

Adding a new tool is straightforward:

1. Create a handler file in the appropriate module directory
2. Export `{ name, description, inputSchema, handler }` with Zod schema
3. Add it to the module's `index.js` exports via `server.tool(...)`
4. It's automatically registered via `index.js` at the root

### Adding a Resource

1. Create a resource handler in the module's `resources.js`
2. Use `server.resource(name, uri, metadata, handler)` for static resources
3. Use `ResourceTemplate` with `complete` callbacks for completions
4. Export `register{Module}Resources(server)` from module's `index.js`

### Adding a Prompt

1. Add prompt to `prompts/index.js` using `server.prompt(name, description, argsSchema, handler)`
2. Prompts are automatically registered via `index.js` at the root

### Adding Completions

1. Add completion function to `completion/index.js`
2. Reference it in a `ResourceTemplate`'s `complete` property