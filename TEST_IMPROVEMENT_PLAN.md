# Test Improvement Plan

## Current State (as of commit 4da1919)

| Metric | Value |
|--------|-------|
| Total Tests | 92 passing |
| Coverage (Statements) | 67.21% |
| Coverage (Branches) | 63.32% |
| Coverage (Functions) | 72.46% |
| Coverage (Lines) | 67.15% |

### Coverage by Module

| Module | Stmts | Branches | Funcs | Lines | Status |
|--------|-------|----------|-------|-------|--------|
| **auth/token-manager.js** | 73.22% | 68.47% | 75% | 73.22% | Partial - refresh logic untested |
| **auth/oauth-server.js** | 97.67% | 90% | 100% | 97.67% | ✅ Good |
| **auth/index.js** | 8.33% | 0% | 0% | 8.33% | ❌ Nearly untested |
| **auth/tools.js** | 17.24% | 0% | 0% | 17.24% | ❌ Nearly untested |
| **email/list.js** | 100% | 100% | 100% | 100% | ✅ Excellent |
| **email/folder-utils.js** | 87.28% | 85.96% | 91.66% | 87.17% | ✅ Good |
| **calendar/create.js** | 93.75% | 100% | 50% | 93.75% | ✅ Good |
| **utils/graph-client.js** | 8.77% | 0% | 0% | 8.77% | ❌ Nearly untested |
| **utils/response-formatter.js** | 75% | 50% | 66.66% | 75% | ⚠️ Partial |
| **utils/response-helpers.js** | 85.71% | 83.33% | 100% | 85.71% | ✅ Good |
| **All other modules** | 0% | 0% | 0% | 0% | ❌ Untested |

---

## Priority 1: Critical Gaps (High Impact, Medium Effort)

### 1.1 auth/index.js (8.33% coverage)
**Missing tests for:**
- `ensureAuthenticated()` - main auth gate
- `tokenManager` exports
- Error propagation

**Test file needed:** `test/auth/index.test.js`

### 1.2 auth/tools.js (17.24% coverage)
**Missing tests for:**
- `about` tool
- `authenticate` tool (OAuth URL generation)
- `check-auth-status` tool

**Test file needed:** `test/auth/tools.test.js`

### 1.3 utils/graph-client.js (8.77% coverage)
**Missing tests for:**
- `getGraphClient()` - client creation, auth integration
- `graphGetPaginated()` - pagination logic, @odata.nextLink handling
- Error handling (auth errors, network errors)

**Test file needed:** `test/utils/graph-client.test.js`

---

## Priority 2: Core Module Coverage (High Impact, Higher Effort)

### 2.1 email module (partial)
**Missing test files:**
- `test/email/read.test.js` - read emails (preview/full body)
- `test/email/send.test.js` - send, reply, forward
- `test/email/search.test.js` - search functionality
- `test/email/draft.test.js` - draft management
- `test/email/mark-as-read.test.js` - mark read/unread
- `test/email/attachments.test.js` - attachment handling
- `test/email/flag.test.js` - flag/unflag

### 2.2 calendar module (partial)
**Missing test files:**
- `test/calendar/list.test.js` - list events/calendars
- `test/calendar/update.test.js` - update events
- `test/calendar/accept.test.js` - accept/decline/tentative
- `test/calendar/cancel.test.js` - cancel events
- `test/calendar/delete.test.js` - delete events
- `test/calendar/free-busy.test.js` - free/busy lookup

### 2.3 folder module (0%)
**Missing test files:**
- `test/folder/list.test.js`
- `test/folder/create.test.js`
- `test/folder/move.test.js`
- `test/folder/rename.test.js`
- `test/folder/delete.test.js`

### 2.4 contacts module (0%)
**Missing test files:**
- `test/contacts/list.test.js`
- `test/contacts/create.test.js` (basic + full)
- `test/contacts/update.test.js`
- `test/contacts/delete.test.js`

### 2.5 tasks module (0%)
**Missing test files:**
- `test/tasks/list.test.js` (lists + tasks)
- `test/tasks/create.test.js`
- `test/tasks/update.test.js`
- `test/tasks/complete.test.js`
- `test/tasks/delete.test.js`

### 2.6 rules module (0%)
**Missing test files:**
- `test/rules/list.test.js`
- `test/rules/create.test.js`
- `test/rules/update.test.js`
- `test/rules/delete.test.js`
- `test/rules/edit-sequence.test.js`

### 2.7 mailbox module (0%)
**Missing test files:**
- `test/mailbox/settings.test.js`
- `test/mailbox/out-of-office.test.js`

---

## Priority 3: New MCP Features (Medium Impact, Medium Effort)

### 3.1 Resources (0% coverage)
**Missing test files:**
- `test/resources/email.test.js`
- `test/resources/calendar.test.js`
- `test/resources/contacts.test.js`
- `test/resources/folder.test.js`
- `test/resources/mailbox.test.js`
- `test/resources/tasks.test.js`
- `test/resources/rules.test.js`

### 3.2 Prompts (0% coverage)
**Missing test files:**
- `test/prompts/index.test.js` - all 7 prompts

### 3.3 Completions (0% coverage)
**Missing test files:**
- `test/completion/index.test.js` - all 9 completion handlers

### 3.4 Main entry point (0% coverage)
**Missing test file:**
- `test/index.test.js` - McpServer initialization, tool/resource/prompt/completion registration

---

## Priority 4: Edge Cases & Integration (Lower Priority)

### 4.1 Error handling improvements
- Network timeout scenarios
- Rate limiting (429) handling
- Expired token during request
- Partial failures in batch operations

### 4.2 Response format testing
- Test both `toon` and `text` response formats
- Verify structured payload + text fallback

### 4.3 Token manager edge cases (gaps in current tests)
Lines uncovered in token-manager.js:
- 169-170: refresh token from disk fallback
- 176-179: refresh token preservation
- 191: no refresh token error
- 200: refresh token missing error
- 208-303: refreshAccessToken() network flow
- 330: refresh failure in getAccessToken()
- 417: code exchange error handling
- 460: code exchange network error
- 505: token-manager.test.js unused import warning

---

## Recommended Implementation Order

### Phase A: Foundation (Week 1)
1. `test/utils/graph-client.test.js` - core utility
2. `test/auth/index.test.js` - auth gate
3. `test/auth/tools.test.js` - auth tools
4. Fix token-manager.js uncovered lines

**Target: 80%+ coverage on auth + utils**

### Phase B: Email & Calendar (Week 2)
5. `test/email/read.test.js`
6. `test/email/send.test.js`
7. `test/email/search.test.js`
8. `test/calendar/list.test.js`
9. `test/calendar/update.test.js`

**Target: 85%+ coverage on email + calendar**

### Phase C: Remaining Core Modules (Week 3)
10-18. folder, contacts, tasks, rules, mailbox test files

**Target: 75%+ coverage on all core modules**

### Phase D: MCP Features (Week 4)
19-25. Resources, prompts, completions, main index

**Target: 80%+ coverage on new features**

---

## Test Patterns to Follow

### Existing Good Patterns
```javascript
// 1. Mock setup at top
jest.mock('../../utils/graph-client');
jest.mock('../../auth', () => ({ ensureAuthenticated: jest.fn().mockResolvedValue(true) }));

// 2. Force text mode for assertions
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

// 3. Helper for mock Graph client
function createMockClient(responses) { ... }

// 4. Clear mocks + console spy in beforeEach
beforeEach(() => {
  jest.clearAllMocks();
  jest.spyOn(console, 'error').mockImplementation(() => {});
});
```

### New Pattern for McpServer Tests
```javascript
// Test resource registration
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import index from '../../index';

test('should register all 55 tools', () => {
  const server = new McpServer({ name: 'test', version: '1.0.0' });
  // call registration function
  expect(server._tools.size).toBe(55);
});
```

---

## Coverage Targets

| Milestone | Overall Coverage | Key Modules ≥ 80% |
|-----------|------------------|-------------------|
| Current | 67% | 3/20 |
| Phase A | 75% | 6/20 |
| Phase B | 82% | 10/20 |
| Phase C | 88% | 16/20 |
| Phase D | 92% | 20/20 |

---

## Quick Wins (Can do immediately)

1. **Add token-manager edge case tests** - cover lines 169-170, 176-179, 191, 200
2. **Add utils/response-formatter tests** - cover lines 29, 41
3. **Add utils/response-helpers tests** - cover line 17
4. **Fix test/auth/token-manager.test.js** - remove unused `fs` import (lint warning)

---

## Notes

- Keep using CommonJS `require` / `module.exports` (no ESM in tests)
- Mock `utils/graph-client` and `auth` modules consistently
- Use `process.env.OUTLOOK_RESPONSE_FORMAT = 'text'` for plain-text assertions
- Target: 0 lint errors, maintain 90 warnings as acceptable
- All new tests must pass: `npm test`