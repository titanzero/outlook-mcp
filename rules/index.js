const { handleListRules, getInboxRules } = require('./list');
const handleCreateRule = require('./create');
const { getGraphClient } = require('../utils/graph-client');
const config = require('../config');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { registerRulesResources } = require('./resources');
const { handleCreateAdvancedRule, handleUpdateAdvancedRule, handleGetRule, handleTestRule, handleBulkEnableRules } = require('./advanced');
const { handleListRuleTemplates, handleCreateRuleFromTemplate, handleExportRules, handleImportRules } = require('./templates');
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

  // Advanced Rules (extended conditions and actions)
  server.tool(
    'create-advanced-rule',
    'Creates a new inbox rule with full conditions and actions support',
    {
      name: z.string().describe('Name of the rule to create'),
      isEnabled: z.boolean().optional().describe('Whether the rule should be enabled (default: true)'),
      sequence: z.number().int().positive().optional().describe(`Execution order (lower runs first, default: ${config.DEFAULT_RULE_SEQUENCE})`),
      // Conditions
      fromAddresses: z.string().optional().describe('Comma-separated sender emails (exact match)'),
      senderAddressIncludes: z.string().optional().describe('Comma-separated sender patterns (contains match)'),
      recipientAddressIncludes: z.string().optional().describe('Comma-separated recipient patterns (contains match)'),
      containsSubject: z.string().optional().describe('Subject text to match'),
      containsBody: z.string().optional().describe('Body text to match'),
      containsBodyOrSubject: z.string().optional().describe('Text to match in body or subject'),
      importance: z.enum(['low', 'normal', 'high']).optional().describe('Importance level to match'),
      hasAttachments: z.boolean().optional().describe('Match emails with attachments'),
      sentToMe: z.boolean().optional().describe('Match emails sent directly to me'),
      ccMe: z.boolean().optional().describe('Match emails where I am CC\'d'),
      messageSizeOver: z.string().optional().describe('Match emails larger than bytes'),
      messageSizeUnder: z.string().optional().describe('Match emails smaller than bytes'),
      categories: z.string().optional().describe('Comma-separated categories to match'),
      // Actions
      moveToFolder: z.string().optional().describe('Move matching emails to folder'),
      copyToFolder: z.string().optional().describe('Copy matching emails to folder'),
      delete: z.boolean().optional().describe('Permanently delete matching emails'),
      markAsRead: z.boolean().optional().describe('Mark matching emails as read'),
      markImportance: z.enum(['low', 'normal', 'high']).optional().describe('Set importance on matching emails'),
      forwardTo: z.string().optional().describe('Comma-separated emails to forward to'),
      redirectTo: z.string().optional().describe('Comma-separated emails to redirect to'),
      assignCategories: z.string().optional().describe('Comma-separated categories to assign'),
      stopProcessingRules: z.boolean().optional().describe('Stop processing further rules'),
      flagForAction: z.string().optional().describe('Flag for follow-up with custom text'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    handleCreateAdvancedRule
  );

  server.tool(
    'update-advanced-rule',
    'Updates an inbox rule with full conditions and actions support',
    {
      ruleName: z.string().describe('Exact name of the rule to update'),
      isEnabled: z.boolean().optional().describe('Enable or disable the rule'),
      sequence: z.number().int().positive().optional().describe('New execution order'),
      // Conditions (set to empty string to clear)
      fromAddresses: z.string().optional().describe('Comma-separated sender emails (empty to clear)'),
      senderAddressIncludes: z.string().optional().describe('Comma-separated sender patterns (empty to clear)'),
      recipientAddressIncludes: z.string().optional().describe('Comma-separated recipient patterns (empty to clear)'),
      containsSubject: z.string().optional().describe('Subject text to match (empty to clear)'),
      containsBody: z.string().optional().describe('Body text to match (empty to clear)'),
      containsBodyOrSubject: z.string().optional().describe('Body or subject text (empty to clear)'),
      importance: z.enum(['low', 'normal', 'high']).optional().describe('Importance level (empty to clear)'),
      hasAttachments: z.boolean().optional().describe('Match attachments (empty to clear)'),
      sentToMe: z.boolean().optional().describe('Sent to me (empty to clear)'),
      ccMe: z.boolean().optional().describe('CC me (empty to clear)'),
      messageSizeOver: z.string().optional().describe('Size over bytes (empty to clear)'),
      messageSizeUnder: z.string().optional().describe('Size under bytes (empty to clear)'),
      categories: z.string().optional().describe('Categories to match (empty to clear)'),
      clearConditions: z.boolean().optional().describe('Clear all conditions'),
      // Actions
      moveToFolder: z.string().optional().describe('Move to folder (empty to clear)'),
      copyToFolder: z.string().optional().describe('Copy to folder (empty to clear)'),
      delete: z.boolean().optional().describe('Delete action'),
      markAsRead: z.boolean().optional().describe('Mark as read'),
      markImportance: z.enum(['low', 'normal', 'high']).optional().describe('Set importance (empty to clear)'),
      forwardTo: z.string().optional().describe('Forward to emails (empty to clear)'),
      redirectTo: z.string().optional().describe('Redirect to emails (empty to clear)'),
      assignCategories: z.string().optional().describe('Assign categories (empty to clear)'),
      stopProcessingRules: z.boolean().optional().describe('Stop processing rules'),
      flagForAction: z.string().optional().describe('Flag for action (empty to clear)'),
      clearActions: z.boolean().optional().describe('Clear all actions'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleUpdateAdvancedRule
  );

  server.tool(
    'get-rule',
    'Gets detailed rule with all conditions and actions',
    {
      ruleName: z.string().describe('Exact name of the rule'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleGetRule
  );

  server.tool(
    'test-rule',
    'Tests a rule against a specific email to see if it would match',
    {
      ruleName: z.string().describe('Rule name to test'),
      emailId: z.string().describe('Email ID to test against (from list-emails)'),
    },
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleTestRule
  );

  server.tool(
    'bulk-enable-rules',
    'Enables or disables multiple rules at once',
    {
      ruleNames: z.array(z.string()).describe('Array of rule names'),
      isEnabled: z.boolean().describe('True to enable, false to disable'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    handleBulkEnableRules
  );

  // Rule Templates
  server.tool(
    'list-rule-templates',
    'Lists available rule templates for common scenarios',
    {},
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleListRuleTemplates
  );

  server.tool(
    'create-rule-from-template',
    'Creates a rule from a predefined template',
    {
      templateId: z.string().describe('Template ID (from list-rule-templates)'),
      parameters: z.record(z.string()).describe('Template parameters as key-value pairs'),
      ruleName: z.string().optional().describe('Custom rule name (optional)'),
      isEnabled: z.boolean().optional().describe('Enable rule after creation (default: true)'),
      sequence: z.number().int().positive().optional().describe('Execution order'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    handleCreateRuleFromTemplate
  );

  // Import/Export
  server.tool(
    'export-rules',
    'Exports all inbox rules as JSON for backup or migration',
    {},
    {
      annotations: {
        readOnlyHint: true,
        openWorldHint: false,
      },
    },
    handleExportRules
  );

  server.tool(
    'import-rules',
    'Imports rules from JSON export',
    {
      rules: z.array(z.object({
        displayName: z.string(),
        sequence: z.number().int().positive().optional(),
        isEnabled: z.boolean().optional(),
        conditions: z.any().optional(),
        actions: z.any().optional(),
      })).describe('Array of rule objects from export-rules'),
      skipExisting: z.boolean().optional().describe('Skip rules that already exist (default: true)'),
      overwriteSequence: z.boolean().optional().describe('Use sequence from import (default: false)'),
    },
    {
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    handleImportRules
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
  handleCreateAdvancedRule,
  handleUpdateAdvancedRule,
  handleGetRule,
  handleTestRule,
  handleBulkEnableRules,
  handleListRuleTemplates,
  handleCreateRuleFromTemplate,
  handleExportRules,
  handleImportRules,
};
