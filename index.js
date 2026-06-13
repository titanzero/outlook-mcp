#!/usr/bin/env node
const { Server } = require("@modelcontextprotocol/sdk/server/index.js");
const { StdioServerTransport } = require("@modelcontextprotocol/sdk/server/stdio.js");
const {
  ListToolsRequestSchema,
  CallToolRequestSchema,
  McpError,
  ErrorCode,
} = require("@modelcontextprotocol/sdk/types.js");
const config = require('./config');

const { authTools } = require('./auth');
const { calendarTools } = require('./calendar');
const { emailTools } = require('./email');
const { folderTools } = require('./folder');
const { rulesTools } = require('./rules');
const { contactsTools } = require('./contacts');
const { mailboxTools } = require('./mailbox');
const { tasksTools } = require('./tasks');

console.error(`STARTING ${config.SERVER_NAME.toUpperCase()} MCP SERVER`);

const TOOLS = [
  ...authTools,
  ...calendarTools,
  ...emailTools,
  ...folderTools,
  ...rulesTools,
  ...contactsTools,
  ...mailboxTools,
  ...tasksTools,
];

const server = new Server(
  { name: config.SERVER_NAME, version: config.SERVER_VERSION },
  {
    capabilities: { tools: {} },
    instructions:
      "Outlook data in tool responses is provided in TOON (Token-Oriented Object Notation) format, " +
      "a compact encoding optimized for LLM token efficiency. Parse TOON responses as structured data. " +
      "When sending data back (e.g. creating emails or events), use standard JSON arguments as defined " +
      "in each tool's input schema.",
  }
);

server.setRequestHandler(ListToolsRequestSchema, async () => {
  console.error(`TOOLS LIST REQUEST (${TOOLS.length} tools)`);
  return {
    tools: TOOLS.map(({ name, description, inputSchema }) => ({
      name,
      description,
      inputSchema,
    })),
  };
});

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args = {} } = request.params;
  console.error(`TOOL CALL: ${name}`);

  const tool = TOOLS.find(t => t.name === name);
  if (!tool?.handler) {
    throw new McpError(ErrorCode.MethodNotFound, `Tool not found: ${name}`);
  }

  return await tool.handler(args);
});

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
