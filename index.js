#!/usr/bin/env node
const { McpServer } = require("@modelcontextprotocol/sdk/server/mcp.js");
const { StdioServerTransport } = require("@modelcontextprotocol/sdk/server/stdio.js");
const config = require('./config');

const { registerAuthTools } = require('./auth');
const { registerCalendarTools, registerCalendarResources } = require('./calendar');
const { registerEmailTools, registerEmailResources } = require('./email');
const { registerFolderTools, registerFolderResources } = require('./folder');
const { registerRulesTools, registerRulesResources } = require('./rules');
const { registerContactsTools, registerContactsResources } = require('./contacts');
const { registerMailboxTools, registerMailboxResources } = require('./mailbox');
const { registerTasksTools, registerTasksResources } = require('./tasks');

console.error(`STARTING ${config.SERVER_NAME.toUpperCase()} MCP SERVER`);

const server = new McpServer(
  { name: config.SERVER_NAME, version: config.SERVER_VERSION },
  {
    capabilities: { 
      tools: { listChanged: true },
      resources: { listChanged: true },
      prompts: { listChanged: true }
    },
    instructions:
      "Outlook data in tool responses is provided in TOON (Token-Oriented Object Notation) format, " +
      "a compact encoding optimized for LLM token efficiency. Parse TOON responses as structured data. " +
      "When sending data back (e.g. creating emails or events), use standard JSON arguments as defined " +
      "in each tool's input schema.",
  }
);

// Register all tool modules
registerAuthTools(server);
registerCalendarTools(server);
registerEmailTools(server);
registerFolderTools(server);
registerRulesTools(server);
registerContactsTools(server);
registerMailboxTools(server);
registerTasksTools(server);

// Register all resource modules
registerFolderResources(server);
registerContactsResources(server);
registerCalendarResources(server);
registerTasksResources(server);
registerMailboxResources(server);
registerEmailResources(server);
registerRulesResources(server);

process.on('SIGTERM', () => {
  console.error('SIGTERM received, shutting down');
  process.exit(0);
});

const transport = new StdioServerTransport();
server.connect(transport)
  .then(() => console.error(`${config.SERVER_NAME} connected and listening`))
  .catch(error => {
    console.error(`Connection error: ${error.message}`);
    process.exit(1);
  });