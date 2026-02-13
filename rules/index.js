/**
 * Email rules management module for Outlook MCP server
 * Uses the Microsoft Graph JS SDK.
 */
const { handleListRules, getInboxRules } = require('./list');
const handleCreateRule = require('./create');
const { getGraphClient } = require('../utils/graph-client');
const config = require('../config');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');

/**
 * Edit rule sequence handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleEditRuleSequence(args) {
  const { ruleName, sequence } = args;
  
  if (!ruleName) {
    return makeErrorResponse('Rule name is required. Please specify the exact name of an existing rule.');
  }
  
  if (!sequence || isNaN(sequence) || sequence < 1) {
    return makeErrorResponse('A positive sequence number is required. Lower numbers run first (higher priority).');
  }
  
  try {
    const client = await getGraphClient();
    
    // Get all rules
    const rules = await getInboxRules(client);
    
    // Find the rule by name
    const rule = rules.find(r => r.displayName === ruleName);
    if (!rule) {
      return makeErrorResponse(`Rule with name "${ruleName}" not found.`);
    }
    
    // Update the rule sequence
    await client.api(`me/mailFolders/inbox/messageRules/${rule.id}`)
      .patch({ sequence: sequence });
    
    return makeResponse(`Successfully updated the sequence of rule "${ruleName}" to ${sequence}.`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    
    return makeErrorResponse(`Error updating rule sequence: ${error.message}`);
  }
}

// Rules management tool definitions
const rulesTools = [
  {
    name: "list-rules",
    description: "Lists inbox rules in your Outlook account",
    inputSchema: {
      type: "object",
      properties: {
        includeDetails: {
          type: "boolean",
          description: "Include detailed rule conditions and actions"
        }
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
        name: {
          type: "string",
          description: "Name of the rule to create"
        },
        fromAddresses: {
          type: "string",
          description: "Comma-separated list of sender email addresses for the rule"
        },
        containsSubject: {
          type: "string",
          description: "Subject text the email must contain"
        },
        hasAttachments: {
          type: "boolean",
          description: "Whether the rule applies to emails with attachments"
        },
        moveToFolder: {
          type: "string",
          description: "Name of the folder to move matching emails to"
        },
        markAsRead: {
          type: "boolean", 
          description: "Whether to mark matching emails as read"
        },
        isEnabled: {
          type: "boolean",
          description: "Whether the rule should be enabled after creation (default: true)"
        },
        sequence: {
          type: "number",
          description: `Order in which the rule is executed (lower numbers run first, default: ${config.DEFAULT_RULE_SEQUENCE})`
        }
      },
      required: ["name"]
    },
    handler: handleCreateRule
  },
  {
    name: "edit-rule-sequence",
    description: "Changes the execution order of an existing inbox rule",
    inputSchema: {
      type: "object",
      properties: {
        ruleName: {
          type: "string",
          description: "Name of the rule to modify"
        },
        sequence: {
          type: "number",
          description: "New sequence value for the rule (lower numbers run first)"
        }
      },
      required: ["ruleName", "sequence"]
    },
    handler: handleEditRuleSequence
  }
];

module.exports = {
  rulesTools,
  handleListRules,
  handleCreateRule,
  handleEditRuleSequence
};
