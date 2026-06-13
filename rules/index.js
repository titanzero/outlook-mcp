const { handleListRules, getInboxRules } = require('./list');
const handleCreateRule = require('./create');
const { getGraphClient } = require('../utils/graph-client');
const config = require('../config');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

async function handleEditRuleSequence(args) {
  const { ruleName, sequence } = args;

  if (!ruleName) return makeErrorResponse('Rule name is required.');
  if (!sequence || isNaN(sequence) || sequence < 1) {
    return makeErrorResponse('A positive sequence number is required. Lower numbers run first (higher priority).');
  }

  try {
    const client = await getGraphClient();
    const rules = await getInboxRules(client);
    const rule = rules.find(r => r.displayName === ruleName);
    if (!rule) return makeErrorResponse(`Rule "${ruleName}" not found.`);

    await client.api(`me/mailFolders/inbox/messageRules/${rule.id}`).patch({ sequence });

    return makeResponse(`Successfully updated the sequence of rule "${ruleName}" to ${sequence}.`);
  } catch (error) {
    if (isAuthError(error)) return makeErrorResponse(error.message);
    return makeErrorResponse(`Error updating rule sequence: ${error.message}`);
  }
}

async function handleDeleteRule(args) {
  const { ruleName } = args;

  if (!ruleName) return makeErrorResponse('Rule name is required.');

  try {
    const client = await getGraphClient();
    const rules = await getInboxRules(client);
    const rule = rules.find(r => r.displayName === ruleName);
    if (!rule) return makeErrorResponse(`Rule "${ruleName}" not found.`);

    await client.api(`me/mailFolders/inbox/messageRules/${rule.id}`).delete();

    return makeResponse(`Rule "${ruleName}" deleted successfully.`);
  } catch (error) {
    if (isAuthError(error)) return makeErrorResponse(error.message);
    return makeErrorResponse(`Error deleting rule: ${error.message}`);
  }
}

async function handleUpdateRule(args) {
  const { ruleName, isEnabled, fromAddresses, containsSubject, moveToFolder, markAsRead, sequence } = args;

  if (!ruleName) return makeErrorResponse('Rule name is required.');

  try {
    const client = await getGraphClient();
    const rules = await getInboxRules(client);
    const rule = rules.find(r => r.displayName === ruleName);
    if (!rule) return makeErrorResponse(`Rule "${ruleName}" not found.`);

    const patch = {};
    if (isEnabled !== undefined) patch.isEnabled = isEnabled;
    if (sequence !== undefined) patch.sequence = sequence;

    if (fromAddresses !== undefined || containsSubject !== undefined) {
      patch.conditions = { ...rule.conditions };
      if (fromAddresses !== undefined) {
        patch.conditions.fromAddresses = fromAddresses.split(',').map(a => ({
          emailAddress: { address: a.trim(), name: a.trim() }
        }));
      }
      if (containsSubject !== undefined) {
        patch.conditions.subjectContains = [containsSubject];
      }
    }

    if (moveToFolder !== undefined || markAsRead !== undefined) {
      patch.actions = { ...rule.actions };
      if (moveToFolder !== undefined) patch.actions.moveToFolder = moveToFolder;
      if (markAsRead !== undefined) patch.actions.markAsRead = markAsRead;
    }

    await client.api(`me/mailFolders/inbox/messageRules/${rule.id}`).patch(patch);

    return makeResponse(`Rule "${ruleName}" updated successfully.`);
  } catch (error) {
    if (isAuthError(error)) return makeErrorResponse(error.message);
    return makeErrorResponse(`Error updating rule: ${error.message}`);
  }
}

const rulesTools = [
  {
    name: "list-rules",
    description: "Lists inbox rules in your Outlook account",
    inputSchema: {
      type: "object",
      properties: {
        includeDetails: { type: "boolean", description: "Include detailed rule conditions and actions" }
      },
      required: []
    },
    handler: handleListRules
  },
  {
    name: "create-rule",
    description: "Creates a new inbox rule",
    inputSchema: {
      type: "object",
      properties: {
        name: { type: "string", description: "Name of the rule to create" },
        fromAddresses: { type: "string", description: "Comma-separated list of sender email addresses for the rule" },
        containsSubject: { type: "string", description: "Subject text the email must contain" },
        hasAttachments: { type: "boolean", description: "Whether the rule applies to emails with attachments" },
        moveToFolder: { type: "string", description: "Name of the folder to move matching emails to" },
        markAsRead: { type: "boolean", description: "Whether to mark matching emails as read" },
        isEnabled: { type: "boolean", description: "Whether the rule should be enabled after creation (default: true)" },
        sequence: { type: "number", description: `Order in which the rule is executed (lower numbers run first, default: ${config.DEFAULT_RULE_SEQUENCE})` }
      },
      required: ["name"]
    },
    handler: handleCreateRule
  },
  {
    name: "update-rule",
    description: "Updates an existing inbox rule's conditions, actions, or status",
    inputSchema: {
      type: "object",
      properties: {
        ruleName: { type: "string", description: "Exact name of the rule to update" },
        isEnabled: { type: "boolean", description: "Enable or disable the rule" },
        sequence: { type: "number", description: "New execution order (lower runs first)" },
        fromAddresses: { type: "string", description: "Comma-separated sender addresses (replaces existing)" },
        containsSubject: { type: "string", description: "New subject filter text" },
        moveToFolder: { type: "string", description: "New destination folder name" },
        markAsRead: { type: "boolean", description: "Whether to mark matching emails as read" }
      },
      required: ["ruleName"]
    },
    handler: handleUpdateRule
  },
  {
    name: "edit-rule-sequence",
    description: "Changes the execution order of an existing inbox rule",
    inputSchema: {
      type: "object",
      properties: {
        ruleName: { type: "string", description: "Name of the rule to modify" },
        sequence: { type: "number", description: "New sequence value for the rule (lower numbers run first)" }
      },
      required: ["ruleName", "sequence"]
    },
    handler: handleEditRuleSequence
  },
  {
    name: "delete-rule",
    description: "Deletes an existing inbox rule",
    inputSchema: {
      type: "object",
      properties: {
        ruleName: { type: "string", description: "Exact name of the rule to delete" }
      },
      required: ["ruleName"]
    },
    handler: handleDeleteRule
  }
];

module.exports = {
  rulesTools,
  handleListRules,
  handleCreateRule,
  handleEditRuleSequence,
  handleUpdateRule,
  handleDeleteRule,
};
