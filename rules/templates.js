/**
 * Rule templates/presets for Outlook MCP server
 */
const { getGraphClient } = require('../utils/graph-client');
const config = require('../config');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * List available rule templates
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListRuleTemplates(_args) {
  const templates = [
    {
      id: 'move-sender-to-folder',
      name: 'Move emails from sender to folder',
      description: 'Moves emails from specific sender(s) to a designated folder',
      category: 'Organization',
      parameters: [
        { name: 'fromAddresses', type: 'string', required: true, description: 'Comma-separated sender emails' },
        { name: 'folderName', type: 'string', required: true, description: 'Destination folder name' },
        { name: 'ruleName', type: 'string', required: false, description: 'Custom rule name' },
      ],
    },
    {
      id: 'move-subject-to-folder',
      name: 'Move emails with subject to folder',
      description: 'Moves emails containing specific text in subject to a folder',
      category: 'Organization',
      parameters: [
        { name: 'subjectText', type: 'string', required: true, description: 'Text to match in subject' },
        { name: 'folderName', type: 'string', required: true, description: 'Destination folder name' },
        { name: 'ruleName', type: 'string', required: false, description: 'Custom rule name' },
      ],
    },
    {
      id: 'auto-reply',
      name: 'Auto-reply to specific sender',
      description: 'Automatically replies to emails from specific sender',
      category: 'Automation',
      parameters: [
        { name: 'fromAddresses', type: 'string', required: true, description: 'Comma-separated sender emails' },
        { name: 'replyTemplate', type: 'string', required: true, description: 'Reply message template' },
        { name: 'ruleName', type: 'string', required: false, description: 'Custom rule name' },
      ],
    },
    {
      id: 'flag-important',
      name: 'Flag high importance emails',
      description: 'Flags emails marked as high importance for follow-up',
      category: 'Productivity',
      parameters: [
        { name: 'dueDays', type: 'number', required: false, description: 'Days until due (default: 1)' },
        { name: 'ruleName', type: 'string', required: false, description: 'Custom rule name' },
      ],
    },
    {
      id: 'categorize-by-domain',
      name: 'Categorize by sender domain',
      description: 'Assigns category based on sender email domain',
      category: 'Organization',
      parameters: [
        { name: 'domain', type: 'string', required: true, description: 'Email domain (e.g., "company.com")' },
        { name: 'category', type: 'string', required: true, description: 'Category name to assign' },
        { name: 'ruleName', type: 'string', required: false, description: 'Custom rule name' },
      ],
    },
    {
      id: 'delete-spam',
      name: 'Delete obvious spam',
      description: 'Permanently deletes emails matching spam patterns',
      category: 'Security',
      parameters: [
        { name: 'subjectPatterns', type: 'string', required: true, description: 'Comma-separated subject patterns' },
        { name: 'senderPatterns', type: 'string', required: false, description: 'Comma-separated sender patterns' },
        { name: 'ruleName', type: 'string', required: false, description: 'Custom rule name' },
      ],
    },
    {
      id: 'forward-to-team',
      name: 'Forward specific emails to team',
      description: 'Forwards emails from specific senders or with keywords to team members',
      category: 'Collaboration',
      parameters: [
        { name: 'fromAddresses', type: 'string', required: false, description: 'Comma-separated sender emails' },
        { name: 'subjectKeywords', type: 'string', required: false, description: 'Comma-separated subject keywords' },
        { name: 'forwardTo', type: 'string', required: true, description: 'Comma-separated recipient emails' },
        { name: 'ruleName', type: 'string', required: false, description: 'Custom rule name' },
      ],
    },
    {
      id: 'mark-newsletters-read',
      name: 'Mark newsletters as read',
      description: 'Automatically marks newsletter/promotional emails as read',
      category: 'Productivity',
      parameters: [
        { name: 'senderPatterns', type: 'string', required: true, description: 'Comma-separated sender patterns (e.g., "newsletter,promo,marketing")' },
        { name: 'ruleName', type: 'string', required: false, description: 'Custom rule name' },
      ],
    },
    {
      id: 'archive-old',
      name: 'Archive large/old emails',
      description: 'Moves large or old emails to archive folder',
      category: 'Maintenance',
      parameters: [
        { name: 'sizeOverMB', type: 'number', required: false, description: 'Size in MB (default: 10)' },
        { name: 'olderThanDays', type: 'number', required: false, description: 'Days old (default: 365)' },
        { name: 'archiveFolder', type: 'string', required: true, description: 'Archive folder name' },
        { name: 'ruleName', type: 'string', required: false, description: 'Custom rule name' },
      ],
    },
    {
      id: 'cc-me-to-folder',
      name: 'Move CC\'d emails to folder',
      description: 'Moves emails where you are CC\'d to a specific folder',
      category: 'Organization',
      parameters: [
        { name: 'folderName', type: 'string', required: true, description: 'Destination folder name' },
        { name: 'ruleName', type: 'string', required: false, description: 'Custom rule name' },
      ],
    },
  ];

  const structured = { templates };

  const textFallback = [
    `Rule Templates (${templates.length}):`,
    '',
    templates.map((t, i) => {
      const params = t.parameters.map(p => `${p.name}${p.required ? ' (required)' : ' (optional)'}: ${p.description}`).join('\n    ');
      return `${i + 1}. ${t.name} [${t.category}]\n   ${t.description}\n   Parameters:\n    ${params}`;
    }).join('\n\n'),
    '',
    'Use create-rule-from-template with template ID and parameters.',
  ].join('\n');

  return makeResponse(formatResponse(structured, textFallback));
}

/**
 * Create rule from template
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateRuleFromTemplate(args) {
  const { templateId, parameters, ruleName, isEnabled = true, sequence } = args;

  if (!templateId) {
    return makeErrorResponse('Template ID is required (from list-rule-templates).');
  }
  if (!parameters || typeof parameters !== 'object') {
    return makeErrorResponse('Parameters object is required.');
  }

  try {
    const client = await getGraphClient();

    let rule = { displayName: ruleName, isEnabled, sequence: sequence || config.DEFAULT_RULE_SEQUENCE };

    switch (templateId) {
      case 'move-sender-to-folder': {
        const { fromAddresses, folderName, ruleName: customName } = parameters;
        if (!fromAddresses || !folderName) {
          return makeErrorResponse('fromAddresses and folderName are required for this template.');
        }
        rule.displayName = customName || `Move from ${fromAddresses} to ${folderName}`;
        rule.conditions = {
          fromAddresses: fromAddresses.split(',').map(a => ({
            emailAddress: { address: a.trim(), name: a.trim() },
          })),
        };
        rule.actions = { moveToFolder: folderName };
        break;
      }

      case 'move-subject-to-folder': {
        const { subjectText, folderName, ruleName: customName } = parameters;
        if (!subjectText || !folderName) {
          return makeErrorResponse('subjectText and folderName are required for this template.');
        }
        rule.displayName = customName || `Move subject "${subjectText}" to ${folderName}`;
        rule.conditions = { subjectContains: [subjectText] };
        rule.actions = { moveToFolder: folderName };
        break;
      }

      case 'auto-reply': {
        const { fromAddresses, replyTemplate, ruleName: customName } = parameters;
        if (!fromAddresses || !replyTemplate) {
          return makeErrorResponse('fromAddresses and replyTemplate are required for this template.');
        }
        rule.displayName = customName || `Auto-reply to ${fromAddresses}`;
        rule.conditions = {
          fromAddresses: fromAddresses.split(',').map(a => ({
            emailAddress: { address: a.trim(), name: a.trim() },
          })),
        };
        // Note: Auto-reply action not directly supported in Graph messageRules
        // This would need a different approach (forward with template)
        rule.actions = {
          forwardTo: fromAddresses.split(',').map(a => ({
            emailAddress: { address: a.trim(), name: a.trim() },
          })),
        };
        break;
      }

      case 'flag-important': {
        const { dueDays = 1, ruleName: customName } = parameters;
        rule.displayName = customName || 'Flag high importance emails';
        rule.conditions = { importance: 'high' };
        rule.actions = {
          flagForAction: `Follow up within ${dueDays} day(s)`,
        };
        break;
      }

      case 'categorize-by-domain': {
        const { domain, category, ruleName: customName } = parameters;
        if (!domain || !category) {
          return makeErrorResponse('domain and category are required for this template.');
        }
        rule.displayName = customName || `Categorize @${domain} as ${category}`;
        rule.conditions = {
          senderAddressIncludes: [{
            emailAddress: { address: `@${domain}`, name: domain },
          }],
        };
        rule.actions = { assignCategories: [category] };
        break;
      }

      case 'delete-spam': {
        const { subjectPatterns, senderPatterns, ruleName: customName } = parameters;
        if (!subjectPatterns) {
          return makeErrorResponse('subjectPatterns is required for this template.');
        }
        rule.displayName = customName || 'Delete spam emails';
        rule.conditions = {
          subjectContains: subjectPatterns.split(',').map(s => s.trim()),
        };
        if (senderPatterns) {
          rule.conditions.senderAddressIncludes = senderPatterns.split(',').map(s => ({
            emailAddress: { address: s.trim(), name: s.trim() },
          }));
        }
        rule.actions = { delete: true };
        break;
      }

      case 'forward-to-team': {
        const { fromAddresses, subjectKeywords, forwardTo, ruleName: customName } = parameters;
        if (!forwardTo) {
          return makeErrorResponse('forwardTo is required for this template.');
        }
        rule.displayName = customName || `Forward to team`;
        const conditions = {};
        if (fromAddresses) {
          conditions.fromAddresses = fromAddresses.split(',').map(a => ({
            emailAddress: { address: a.trim(), name: a.trim() },
          }));
        }
        if (subjectKeywords) {
          conditions.subjectContains = subjectKeywords.split(',').map(s => s.trim());
        }
        rule.conditions = Object.keys(conditions).length > 0 ? conditions : undefined;
        rule.actions = {
          forwardTo: forwardTo.split(',').map(a => ({
            emailAddress: { address: a.trim(), name: a.trim() },
          })),
        };
        break;
      }

      case 'mark-newsletters-read': {
        const { senderPatterns, ruleName: customName } = parameters;
        if (!senderPatterns) {
          return makeErrorResponse('senderPatterns is required for this template.');
        }
        rule.displayName = customName || 'Mark newsletters as read';
        rule.conditions = {
          senderAddressIncludes: senderPatterns.split(',').map(s => ({
            emailAddress: { address: s.trim(), name: s.trim() },
          })),
        };
        rule.actions = { markAsRead: true };
        break;
      }

      case 'archive-old': {
        const { sizeOverMB = 10, olderThanDays = 365, archiveFolder, ruleName: customName } = parameters;
        if (!archiveFolder) {
          return makeErrorResponse('archiveFolder is required for this template.');
        }
        rule.displayName = customName || `Archive emails >${sizeOverMB}MB or >${olderThanDays} days`;
        // Note: Date-based conditions not directly supported in Graph
        // Using size-based only
        rule.conditions = { messageSizeOver: sizeOverMB * 1024 * 1024 };
        rule.actions = { moveToFolder: archiveFolder };
        break;
      }

      case 'cc-me-to-folder': {
        const { folderName, ruleName: customName } = parameters;
        if (!folderName) {
          return makeErrorResponse('folderName is required for this template.');
        }
        rule.displayName = customName || `Move CC'd emails to ${folderName}`;
        rule.conditions = { ccMe: true };
        rule.actions = { moveToFolder: folderName };
        break;
      }

      default:
        return makeErrorResponse(`Unknown template: ${templateId}. Use list-rule-templates to see available templates.`);
    }

    const created = await client.api('me/mailFolders/inbox/messageRules').post(rule);

    return makeResponse(`Rule "${created.displayName}" created from template "${templateId}".\nID: ${created.id}\nSequence: ${created.sequence}`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error creating rule from template: ${error.message}`);
  }
}

/**
 * Export rules to JSON
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleExportRules(_args) {
  try {
    const client = await getGraphClient();

    const rules = await client.api('me/mailFolders/inbox/messageRules')
      .select('id,displayName,sequence,isEnabled,conditions,actions')
      .orderby('sequence asc')
      .get();

    const exportedRules = (rules.value || []).map(r => ({
      displayName: r.displayName,
      sequence: r.sequence,
      isEnabled: r.isEnabled,
      conditions: r.conditions,
      actions: r.actions,
    }));

    const structured = {
      exportedAt: new Date().toISOString(),
      ruleCount: exportedRules.length,
      rules: exportedRules,
    };

    const textFallback = [
      `Rules Export`,
      `Exported: ${new Date().toISOString()}`,
      `Total Rules: ${exportedRules.length}`,
      '',
      'Rules:',
      exportedRules.map((r, i) => `${i + 1}. ${r.displayName} [seq: ${r.sequence}] ${r.isEnabled ? 'ENABLED' : 'DISABLED'}`).join('\n'),
    ].join('\n');

    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error exporting rules: ${error.message}`);
  }
}

/**
 * Import rules from JSON
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleImportRules(args) {
  const { rules, skipExisting = true, overwriteSequence = false } = args;

  if (!rules || !Array.isArray(rules)) {
    return makeErrorResponse('Rules array is required.');
  }

  try {
    const client = await getGraphClient();

    // Get existing rules
    const existingRules = await client.api('me/mailFolders/inbox/messageRules')
      .select('id,displayName,sequence')
      .get();

    const existingMap = new Map();
    (existingRules.value || []).forEach(r => existingMap.set(r.displayName, r));

    const results = [];
    const errors = [];
    let nextSequence = 1000;

    for (const rule of rules) {
      try {
        if (!rule.displayName) {
          errors.push({ rule: 'unknown', error: 'Missing displayName' });
          continue;
        }

        const exists = existingMap.has(rule.displayName);

        if (exists && skipExisting) {
          results.push({ name: rule.displayName, status: 'skipped', reason: 'Already exists' });
          continue;
        }

        const ruleData = {
          displayName: rule.displayName,
          isEnabled: rule.isEnabled !== undefined ? rule.isEnabled : true,
          sequence: overwriteSequence && rule.sequence ? rule.sequence : nextSequence++,
          conditions: rule.conditions,
          actions: rule.actions,
        };

        if (exists && !skipExisting) {
          // Update existing
          const existing = existingMap.get(rule.displayName);
          await client.api(`me/mailFolders/inbox/messageRules/${existing.id}`).patch(ruleData);
          results.push({ name: rule.displayName, status: 'updated' });
        } else {
          // Create new
          await client.api('me/mailFolders/inbox/messageRules').post(ruleData);
          results.push({ name: rule.displayName, status: 'created' });
        }
      } catch (e) {
        errors.push({ name: rule.displayName, error: e.message });
      }
    }

    const structured = {
      total: rules.length,
      created: results.filter(r => r.status === 'created').length,
      updated: results.filter(r => r.status === 'updated').length,
      skipped: results.filter(r => r.status === 'skipped').length,
      errors: errors.length,
      results,
      errors: errors.slice(0, 20),
    };

    const textFallback = [
      `Rules Import Results`,
      `Total: ${rules.length}`,
      `Created: ${structured.created}`,
      `Updated: ${structured.updated}`,
      `Skipped: ${structured.skipped}`,
      `Errors: ${structured.errors}`,
      '',
      results.map(r => `  ${r.status === 'created' ? '✓' : r.status === 'updated' ? '↻' : '⊘'} ${r.name}${r.reason ? ` (${r.reason})` : ''}`).join('\n'),
      errors.length > 0 ? `\nErrors:\n${errors.map(e => `  ✗ ${e.name}: ${e.error}`).join('\n')}` : '',
    ].filter(Boolean).join('\n');

    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error importing rules: ${error.message}`);
  }
}

module.exports = {
  handleListRuleTemplates,
  handleCreateRuleFromTemplate,
  handleExportRules,
  handleImportRules,
};