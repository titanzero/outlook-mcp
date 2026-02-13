# CLAUDE.md

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

## Cursor Rules Map

- `.cursor/rules/general.mdc` — always-apply core conventions (CommonJS, config centralization, baseline response/error/logging patterns)
- `.cursor/rules/graph-api.mdc` — Graph SDK client usage, fluent API, pagination, and field selection rules
- `.cursor/rules/toon-responses.mdc` — TOON data-shaping and formatter usage
- `.cursor/rules/new-tool.mdc` — workflow and checklist for adding/registering new tools
- `.cursor/rules/testing.mdc` — Jest conventions, mocks, and test structure

## Project Structure

```
index.js              — MCP entry point; imports `authTools`, `calendarTools`, `emailTools`, `folderTools`, `rulesTools` and combines them into `TOOLS`
config.js             — all constants: field selections, pagination, auth, timezone, response format
outlook-auth-server.js — standalone Express OAuth server

auth/
  index.js            — exports `authTools`, `ensureAuthenticated`, `tokenManager`
  tools.js            — auth tool definitions (`about`, `authenticate`, `check-auth-status`)
  token-manager.js    — token load/save/refresh logic
  oauth-server.js     — OAuth callback/exchange helpers

email/                — list, search, read, send, mark-as-read + folder-utils.js
calendar/             — list, create, accept, decline, cancel, delete
folder/               — list, create, move
rules/                — list, create, edit-rule-sequence

utils/
  graph-client.js     — Graph SDK wrapper: `getGraphClient()`, `graphGetPaginated()`
  response-formatter.js — TOON / plain-text toggle (reads `config.RESPONSE_FORMAT`)
  response-helpers.js — `isAuthError()`, `makeErrorResponse()`, `makeResponse()`

scripts/              — CLI & debug one-off scripts (not imported by app modules)
  debug-env.js        — debug: print env vars and start MCP server
  test-config.js      — debug: verify config and token path
  find-folder-ids.js  — CLI: find folder IDs via Graph API
  move-github-emails.js — CLI: move GitHub emails to subfolder
  create-notifications-rule.js — CLI: create inbox rules for GitHub
  backup-logs.sh      — shell: backup Claude Desktop logs
  test-direct.sh      — shell: test server directly via nc
  test-modular-server.sh — shell: test server via MCP Inspector
```

## Key Patterns

- **CommonJS** (`require` / `module.exports`) — no ESM
- **Module pattern**: each module `index.js` exports a `{module}Tools` array plus handlers
- **Tool shape**: `{ name, description, inputSchema, handler }`
- **Constants**: keep field selections, defaults, limits, and URLs centralized in `config.js`
- **Logging**: use `console.error()` for diagnostics (stdout is reserved for MCP protocol traffic)

## Graph and Response Flow

- `getGraphClient()` authenticates internally via `ensureAuthenticated()`; handlers should not run a separate auth check.
- Use `graphGetPaginated()` for list endpoints that may return `@odata.nextLink`.
- Data handlers should build a structured payload plus text fallback, then call `formatResponse(structured, textFallback)`.
- Return `makeResponse(text)` for success; in `catch`, use `isAuthError(error)` and `makeErrorResponse(...)`.

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
- Claude Desktop config: `OUTLOOK_CLIENT_ID`, `OUTLOOK_CLIENT_SECRET` (used by MCP server)
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
