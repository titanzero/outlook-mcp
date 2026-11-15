# CLAUDE.md

This project uses a shared **AGENTS.md** for agent-agnostic instructions.
See [AGENTS.md](./AGENTS.md) for the complete reference.

---

## Claude Desktop Configuration

Add to your `claude_desktop_config.json`:

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

**Important:** Use the Azure secret **VALUE**, not the Secret ID.

Restart Claude Desktop after config changes.

---

## Quick Reference

| Task               | Command               |
| ------------------ | --------------------- |
| Install deps       | `npm install`         |
| Start server       | `npm start`           |
| Start auth server  | `npm run auth-server` |
| Open MCP Inspector | `npm run inspect`     |
| Run tests          | `npm test`            |
| Debug env          | `npm run debug`       |
| Free port 3333     | `npx kill-port 3333`  |
