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
index.js              — MCP entry point; imports `authTools`, `calendarTools`, `emailTools`, `folderTools`, `rulesTools` and combines them into `TOOLS`
config.js             — all constants: field selections, pagination, auth, timezone, response format
outlook-auth-server.js — standalone Express OAuth server

auth/
  index.js            — exports `authTools`, `ensureAuthenticated`, `tokenManager`
  tools.js            — auth tool definitions (`about`, `authenticate`, `check-auth-status`)
  token-manager.js    — token load/save/refresh logic
  oauth-server.js     — OAuth callback/exchange helpers

email/                — list, search, read, send, mark-as-read + folder-utils.js, shared-mailbox.js, groups.js, conversations.js, recall.js, kql-search.js
calendar/             — list, create, accept, decline, cancel, delete, sharing.js, rooms.js, groups.js, view.js, reminders.js, online-meetings.js
folder/               — list, create, move, bulk-operations.js, permissions.js, statistics.js, hierarchy.js
rules/                — list, create, edit-rule-sequence
contacts/             — list, create, update, delete, folders.js, photos.js, lists.js, search.js
mailbox/              — settings, out-of-office
tasks/                — list, create, update, complete, delete, steps.js, recurrence.js, attachments.js, linked-resources.js, categories.js

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
  backup-logs.sh      — shell: backup MCP client logs
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

| Constant                 | Value                                             |
| ------------------------ | ------------------------------------------------- |
| `DEFAULT_PAGE_SIZE`      | 50                                                |
| `MAX_RESULT_COUNT`       | 1000                                              |
| `DEFAULT_TIMEZONE`       | "Central European Standard Time"                  |
| `RESPONSE_FORMAT`        | `process.env.OUTLOOK_RESPONSE_FORMAT \|\| 'toon'` |
| `EMAIL_SELECT_FIELDS`    | `id,subject,from,receivedDateTime,isRead`         |
| `EMAIL_DETAIL_FIELDS`    | bodyPreview-based (no full body)                  |
| `EMAIL_FULL_BODY_FIELDS` | includes `body` field (used when `fullBody=true`) |

## Environment Variables

- `.env` file: `OUTLOOK_CLIENT_ID`, `OUTLOOK_CLIENT_SECRET`
- MCP client config: `OUTLOOK_CLIENT_ID`, `OUTLOOK_CLIENT_SECRET` (used by MCP server)
- `OUTLOOK_RESPONSE_FORMAT` — `'toon'` (default) or `'text'`
- **Important**: use Azure secret **VALUE**, not the Secret ID
- Token storage: `~/.outlook-mcp-tokens.json` (via `config.getTokenStorePath()`)

## Auth Flow

1. Register Azure app with permissions: 
   - Mail.Read, Mail.ReadWrite, Mail.Send, **Mail.Read.Shared**
   - Calendars.Read, Calendars.ReadWrite, **Calendars.Read.Shared**
   - Contacts.Read, **Contacts.ReadWrite**
   - MailboxSettings.ReadWrite, MailboxSettings.Read
   - MailboxFolder.ReadWrite, MailboxFolder.Read
   - **Tasks.ReadWrite**
   - **Group.Read.All**
   - User.Read, offline_access, email, openid
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
2. Export `{ name, description, inputSchema, handler }`
3. Add it to the module's `index.js` exports
4. It's automatically registered via `index.js` at the root

## Git Conventions

- **Conventional Commits**: Use `type(scope): description` format
  - `feat:` — new feature
  - `fix:` — bug fix
  - `chore:` — maintenance, tooling, deps
  - `docs:` — documentation only
  - `refactor:` — code restructure
  - `test:` — test changes
  - `perf:` — performance improvement
- **Branch names**: `feat/short-name`, `fix/short-name`, `chore/short-name`
- **Workflow**: Squash feature branches to single commit before merge to main
