/**
 * Rules resources for MCP
 * Exposes inbox rules as addressable resources
 */
const { getGraphClient, graphGetPaginated } = require('../utils/graph-client');
const { isAuthError } = require('../utils/response-helpers');
const { ResourceTemplate } = require('@modelcontextprotocol/sdk/server/mcp.js');
const { getRuleCompletions } = require('../completion');

/**
 * Register rules resources with the MCP server
 * @param {McpServer} server - The MCP server instance
 */
function registerRulesResources(server) {
  // Resource for list of inbox rules
  server.resource(
    'outlook-rules',
    'outlook://rules',
    {
      title: 'Inbox Rules',
      description: 'List all inbox rules in the account',
    },
    async uri => {
      try {
        const client = await getGraphClient();
        const response = await graphGetPaginated(client, 'me/mailFolders/inbox/messageRules', {
          $select: 'id,displayName,sequence,isEnabled,conditions,actions',
          $orderby: 'sequence asc',
        });

        const rules = response.value || [];

        const structured = rules.map(r => ({
          id: r.id,
          name: r.displayName,
          sequence: r.sequence,
          isEnabled: r.isEnabled,
          conditions: r.conditions
            ? {
              fromAddresses:
                  r.conditions.fromAddresses?.map(a => a.emailAddress?.address).filter(Boolean) ||
                  [],
              subjectContains: r.conditions.subjectContains || [],
              bodyContains: r.conditions.bodyContains || [],
              importance: r.conditions.importance,
              hasAttachments: r.conditions.hasAttachments,
              sentToMe: r.conditions.sentToMe,
              ccMe: r.conditions.ccMe,
              senderAddressIncludes:
                  r.conditions.senderAddressIncludes
                    ?.map(a => a.emailAddress?.address)
                    .filter(Boolean) || [],
              recipientAddressIncludes:
                  r.conditions.recipientAddressIncludes
                    ?.map(a => a.emailAddress?.address)
                    .filter(Boolean) || [],
              messageSizeOver: r.conditions.messageSizeOver,
              messageSizeUnder: r.conditions.messageSizeUnder,
            }
            : null,
          actions: r.actions
            ? {
              moveToFolder: r.actions.moveToFolder,
              copyToFolder: r.actions.copyToFolder,
              delete: r.actions.delete,
              markAsRead: r.actions.markAsRead,
              markImportance: r.actions.markImportance,
              forwardTo:
                  r.actions.forwardTo?.map(a => a.emailAddress?.address).filter(Boolean) || [],
              redirectTo:
                  r.actions.redirectTo?.map(a => a.emailAddress?.address).filter(Boolean) || [],
              assignCategories: r.actions.assignCategories || [],
              stopProcessingRules: r.actions.stopProcessingRules,
            }
            : null,
        }));

        const textFallback =
          rules.length === 0
            ? 'No inbox rules found.'
            : `Inbox Rules (${rules.length}):\n\n` +
              rules
                .map(
                  (r, i) =>
                    `${i + 1}. ${r.displayName} [seq: ${r.sequence}] ${r.isEnabled ? 'ENABLED' : 'DISABLED'}\n` +
                    `   Conditions: ${JSON.stringify(r.conditions || {})}\n` +
                    `   Actions: ${JSON.stringify(r.actions || {})}\n` +
                    `   ID: ${r.id}`
                )
                .join('\n\n');

        return {
          contents: [
            {
              uri: uri.toString(),
              mimeType: 'application/json',
              text: JSON.stringify({ rules: structured }, null, 2),
            },
          ],
        };
      } catch (error) {
        if (isAuthError(error)) {
          throw error;
        }
        throw new Error(`Error reading rules: ${error.message}`);
      }
    }
  );

  // Resource for individual rule with completion
  const ruleTemplate = new ResourceTemplate('outlook://rules/{ruleId}', {
    complete: {
      ruleId: getRuleCompletions,
    },
  });

  server.resource(
    'outlook-rule',
    ruleTemplate,
    {
      title: 'Inbox Rule',
      description: 'Access a specific inbox rule by ID',
    },
    async (uri, { ruleId }) => {
      try {
        const client = await getGraphClient();
        const rule = await client.api(`me/mailFolders/inbox/messageRules/${ruleId}`).get();

        const structured = {
          id: rule.id,
          name: rule.displayName,
          sequence: rule.sequence,
          isEnabled: rule.isEnabled,
          conditions: rule.conditions
            ? {
              fromAddresses:
                  rule.conditions.fromAddresses
                    ?.map(a => a.emailAddress?.address)
                    .filter(Boolean) || [],
              subjectContains: rule.conditions.subjectContains || [],
              bodyContains: rule.conditions.bodyContains || [],
              importance: rule.conditions.importance,
              hasAttachments: rule.conditions.hasAttachments,
              sentToMe: rule.conditions.sentToMe,
              ccMe: rule.conditions.ccMe,
              senderAddressIncludes:
                  rule.conditions.senderAddressIncludes
                    ?.map(a => a.emailAddress?.address)
                    .filter(Boolean) || [],
              recipientAddressIncludes:
                  rule.conditions.recipientAddressIncludes
                    ?.map(a => a.emailAddress?.address)
                    .filter(Boolean) || [],
              messageSizeOver: rule.conditions.messageSizeOver,
              messageSizeUnder: rule.conditions.messageSizeUnder,
            }
            : null,
          actions: rule.actions
            ? {
              moveToFolder: rule.actions.moveToFolder,
              copyToFolder: rule.actions.copyToFolder,
              delete: rule.actions.delete,
              markAsRead: rule.actions.markAsRead,
              markImportance: rule.actions.markImportance,
              forwardTo:
                  rule.actions.forwardTo?.map(a => a.emailAddress?.address).filter(Boolean) || [],
              redirectTo:
                  rule.actions.redirectTo?.map(a => a.emailAddress?.address).filter(Boolean) || [],
              assignCategories: rule.actions.assignCategories || [],
              stopProcessingRules: rule.actions.stopProcessingRules,
            }
            : null,
        };

        const textFallback = [
          `Rule: ${rule.displayName}`,
          `ID: ${rule.id}`,
          `Sequence: ${rule.sequence}`,
          `Enabled: ${rule.isEnabled ? 'Yes' : 'No'}`,
          '',
          'Conditions:',
          rule.conditions ? JSON.stringify(rule.conditions, null, 2) : '  (none)',
          '',
          'Actions:',
          rule.actions ? JSON.stringify(rule.actions, null, 2) : '  (none)',
        ].join('\n');

        return {
          contents: [
            {
              uri: uri.toString(),
              mimeType: 'application/json',
              text: JSON.stringify(structured, null, 2),
            },
          ],
        };
      } catch (error) {
        if (isAuthError(error)) {
          throw error;
        }
        throw new Error(`Error reading rule: ${error.message}`);
      }
    }
  );
}

module.exports = { registerRulesResources };
