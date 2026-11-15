const { handleListRules, getInboxRules } = require('./list');
const handleCreateRule = require('./create');
const { getGraphClient } = require('../utils/graph-client');
const config = require('../config');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { registerRulesResources } = require('./resources');
const { z } = require('zod');

async function handleEditRuleSequence(args) {
  const { ruleName, sequence } = args;

  if (!ruleName) {
    return makeErrorResponse('Rule name is required.');
  }
  if (!sequence || isNaN(sequence) || sequence < 1) {
    return makeErrorResponse(
      'A positive sequence number is required. Lower numbers run first (higher priority).'
    );
  }

  try {
    const client = await getGraphClient();
    const rules = await getInboxRules(client);
    const rule = rules.find(r => r.displayName === ruleName);
    if (!rule) {
      return makeErrorResponse(`Rule "${ruleName}" not found.`);
    }

    await client.api(`me/mailFolders/inbox/messageRules/${rule.id}`).patch({ sequence });

    return makeResponse(`Successfully updated the sequence of rule "${ruleName}" to ${sequence}.`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error updating rule sequence: ${error.message}`);
  }
}

async function handleDeleteRule(args) {
  const { ruleName } = args;

  if (!ruleName) {
    return makeErrorResponse('Rule name is required.');
  }

  try {
    const client = await getGraphClient();
    const rules = await getInboxRules(client);
    const rule = rules.find(r => r.displayName === ruleName);
    if (!rule) {
      return makeErrorResponse(`Rule "${ruleName}" not found.`);
    }

    await client.api(`me/mailFolders/inbox/messageRules/${rule.id}`).delete();

    return makeResponse(`Rule "${ruleName}" deleted successfully.`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error deleting rule: ${error.message}`);
  }
}

async function handleUpdateRule(args) {
  const {
    ruleName,
    isEnabled,
    fromAddresses,
    containsSubject,
    moveToFolder,
    markAsRead,
    sequence,
  } = args;

  if (!ruleName) {
    return makeErrorResponse('Rule name is required.');
  }

  try {
    const client = await getGraphClient();
    const rules = await getInboxRules(client);
    const rule = rules.find(r => r.displayName === ruleName);
    if (!rule) {
      return makeErrorResponse(`Rule "${ruleName}" not found.`);
    }

    const patch = {};
    if (isEnabled !== undefined) {
      patch.isEnabled = isEnabled;
    }
    if (sequence !== undefined) {
      patch.sequence = sequence;
    }

    if (fromAddresses !== undefined || containsSubject !== undefined) {
      patch.conditions = { ...rule.conditions };
      if (fromAddresses !== undefined) {
        patch.conditions.fromAddresses = fromAddresses.split(',').map(a => ({
          emailAddress: { address: a.trim(), name: a.trim() },
        }));
      }
      if (containsSubject !== undefined) {
        patch.conditions.subjectContains = [containsSubject];
      }
    }

    if (moveToFolder !== undefined || markAsRead !== undefined) {
      patch.actions = { ...rule.actions };
      if (moveToFolder !== undefined) {
        patch.actions.moveToFolder = moveToFolder;
      }
      if (markAsRead !== undefined) {
        patch.actions.markAsRead = markAsRead;
      }
    }

    await client.api(`me/mailFolders/inbox/messageRules/${rule.id}`).patch(patch);

    return makeResponse(`Rule "${ruleName}" updated successfully.`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error updating rule: ${error.message}`);
  }
}

/**
 * Register rules tools with the MCP server
 * @param {McpServer} server - The MCP server instance
 */
function registerRulesTools(server) {
  server.tool(
    'list-rules',
    'Lists inbox rules in your Outlook account',
    {
      includeDetails: z
        .boolean()
        .optional()
        .describe('Include detailed rule conditions and actions'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListRules
  );

  server.tool(
    'create-rule',
    'Creates a new inbox rule',
    {
      name: z.string().describe('Name of the rule to create'),
      fromAddresses: z
        .string()
        .optional()
        .describe('Comma-separated list of sender email addresses for the rule'),
      containsSubject: z.string().optional().describe('Subject text the email must contain'),
      hasAttachments: z
        .boolean()
        .optional()
        .describe('Whether the rule applies to emails with attachments'),
      moveToFolder: z.string().optional().describe('Name of the folder to move matching emails to'),
      markAsRead: z.boolean().optional().describe('Whether to mark matching emails as read'),
      isEnabled: z
        .boolean()
        .optional()
        .describe('Whether the rule should be enabled after creation (default: true)'),
      sequence: z
        .number()
        .int()
        .positive()
        .optional()
        .describe(
          `Order in which the rule is executed (lower numbers run first, default: ${config.DEFAULT_RULE_SEQUENCE})`
        ),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    handleCreateRule
  );

  server.tool(
    'update-rule',
    "Updates an existing inbox rule's conditions, actions, or status",
    {
      ruleName: z.string().describe('Exact name of the rule to update'),
      isEnabled: z.boolean().optional().describe('Enable or disable the rule'),
      sequence: z
        .number()
        .int()
        .positive()
        .optional()
        .describe('New execution order (lower runs first)'),
      fromAddresses: z
        .string()
        .optional()
        .describe('Comma-separated sender addresses (replaces existing)'),
      containsSubject: z.string().optional().describe('New subject filter text'),
      moveToFolder: z.string().optional().describe('New destination folder name'),
      markAsRead: z.boolean().optional().describe('Whether to mark matching emails as read'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleUpdateRule
  );

  server.tool(
    'edit-rule-sequence',
    'Changes the execution order of an existing inbox rule',
    {
      ruleName: z.string().describe('Name of the rule to modify'),
      sequence: z
        .number()
        .int()
        .positive()
        .describe('New sequence value for the rule (lower numbers run first)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleEditRuleSequence
  );

  server.tool(
    'delete-rule',
    'Deletes an existing inbox rule',
    {
      ruleName: z.string().describe('Exact name of the rule to delete'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleDeleteRule
  );
}

module.exports = {
  registerRulesTools,
  registerRulesResources,
  handleListRules,
  handleCreateRule,
  handleEditRuleSequence,
  handleUpdateRule,
  handleDeleteRule,
};
