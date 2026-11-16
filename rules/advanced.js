/**
 * Advanced rules tools for Outlook MCP server
 * Extended conditions and actions for inbox rules
 */
const { getGraphClient } = require('../utils/graph-client');
const config = require('../config');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * Create rule with advanced conditions and actions
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateAdvancedRule(args) {
  const {
    name,
    isEnabled = true,
    sequence,
    // Conditions
    fromAddresses,
    senderAddressIncludes,
    recipientAddressIncludes,
    containsSubject,
    containsBody,
    containsBodyOrSubject,
    importance,
    hasAttachments,
    sentToMe,
    ccMe,
    messageSizeOver,
    messageSizeUnder,
    categories,
    // Actions
    moveToFolder,
    copyToFolder,
    delete: deleteAction,
    markAsRead,
    markImportance,
    forwardTo,
    redirectTo,
    assignCategories,
    stopProcessingRules,
    flagForAction,
  } = args;

  if (!name) {
    return makeErrorResponse('Rule name is required.');
  }

  try {
    const client = await getGraphClient();

    const conditions = {};
    const actions = {};

    // Build conditions
    if (fromAddresses) {
      conditions.fromAddresses = fromAddresses.split(',').map(a => ({
        emailAddress: { address: a.trim(), name: a.trim() },
      }));
    }
    if (senderAddressIncludes) {
      conditions.senderAddressIncludes = senderAddressIncludes.split(',').map(a => ({
        emailAddress: { address: a.trim(), name: a.trim() },
      }));
    }
    if (recipientAddressIncludes) {
      conditions.recipientAddressIncludes = recipientAddressIncludes.split(',').map(a => ({
        emailAddress: { address: a.trim(), name: a.trim() },
      }));
    }
    if (containsSubject) {
      conditions.subjectContains = [containsSubject];
    }
    if (containsBody) {
      conditions.bodyContains = [containsBody];
    }
    if (containsBodyOrSubject) {
      conditions.bodyOrSubjectContains = [containsBodyOrSubject];
    }
    if (importance) {
      conditions.importance = importance;
    }
    if (hasAttachments !== undefined) {
      conditions.hasAttachments = hasAttachments;
    }
    if (sentToMe !== undefined) {
      conditions.sentToMe = sentToMe;
    }
    if (ccMe !== undefined) {
      conditions.ccMe = ccMe;
    }
    if (messageSizeOver !== undefined) {
      conditions.messageSizeOver = parseInt(messageSizeOver);
    }
    if (messageSizeUnder !== undefined) {
      conditions.messageSizeUnder = parseInt(messageSizeUnder);
    }
    if (categories) {
      conditions.categories = categories.split(',').map(c => c.trim());
    }

    // Build actions
    if (moveToFolder) {
      actions.moveToFolder = moveToFolder;
    }
    if (copyToFolder) {
      actions.copyToFolder = copyToFolder;
    }
    if (deleteAction) {
      actions.delete = true;
    }
    if (markAsRead !== undefined) {
      actions.markAsRead = markAsRead;
    }
    if (markImportance) {
      actions.markImportance = markImportance;
    }
    if (forwardTo) {
      actions.forwardTo = forwardTo.split(',').map(a => ({
        emailAddress: { address: a.trim(), name: a.trim() },
      }));
    }
    if (redirectTo) {
      actions.redirectTo = redirectTo.split(',').map(a => ({
        emailAddress: { address: a.trim(), name: a.trim() },
      }));
    }
    if (assignCategories) {
      actions.assignCategories = assignCategories.split(',').map(c => c.trim());
    }
    if (stopProcessingRules !== undefined) {
      actions.stopProcessingRules = stopProcessingRules;
    }
    if (flagForAction) {
      actions.flagForAction = flagForAction;
    }

    const rule = {
      displayName: name,
      isEnabled,
      sequence: sequence || config.DEFAULT_RULE_SEQUENCE,
      conditions: Object.keys(conditions).length > 0 ? conditions : undefined,
      actions: Object.keys(actions).length > 0 ? actions : undefined,
    };

    const created = await client.api('me/mailFolders/inbox/messageRules').post(rule);

    return makeResponse(`Advanced rule "${name}" created successfully.\nID: ${created.id}\nSequence: ${created.sequence}`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error creating advanced rule: ${error.message}`);
  }
}

/**
 * Update rule with advanced conditions and actions
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleUpdateAdvancedRule(args) {
  const {
    ruleName,
    isEnabled,
    sequence,
    // Conditions (set to null to clear, omit to keep existing)
    fromAddresses,
    senderAddressIncludes,
    recipientAddressIncludes,
    containsSubject,
    containsBody,
    containsBodyOrSubject,
    importance,
    hasAttachments,
    sentToMe,
    ccMe,
    messageSizeOver,
    messageSizeUnder,
    categories,
    clearConditions,
    // Actions
    moveToFolder,
    copyToFolder,
    delete: deleteAction,
    markAsRead,
    markImportance,
    forwardTo,
    redirectTo,
    assignCategories,
    stopProcessingRules,
    flagForAction,
    clearActions,
  } = args;

  if (!ruleName) {
    return makeErrorResponse('Rule name is required.');
  }

  try {
    const client = await getGraphClient();
    const rules = await client.api('me/mailFolders/inbox/messageRules').get();
    const rule = rules.value?.find(r => r.displayName === ruleName);

    if (!rule) {
      return makeErrorResponse(`Rule "${ruleName}" not found.`);
    }

    const patch = {};
    if (isEnabled !== undefined) patch.isEnabled = isEnabled;
    if (sequence !== undefined) patch.sequence = sequence;

    // Build conditions
    if (clearConditions) {
      patch.conditions = {};
    } else {
      const conditions = { ...rule.conditions };
      let conditionsChanged = false;

      if (fromAddresses !== undefined) {
        conditions.fromAddresses = fromAddresses ? fromAddresses.split(',').map(a => ({
          emailAddress: { address: a.trim(), name: a.trim() },
        })) : null;
        conditionsChanged = true;
      }
      if (senderAddressIncludes !== undefined) {
        conditions.senderAddressIncludes = senderAddressIncludes ? senderAddressIncludes.split(',').map(a => ({
          emailAddress: { address: a.trim(), name: a.trim() },
        })) : null;
        conditionsChanged = true;
      }
      if (recipientAddressIncludes !== undefined) {
        conditions.recipientAddressIncludes = recipientAddressIncludes ? recipientAddressIncludes.split(',').map(a => ({
          emailAddress: { address: a.trim(), name: a.trim() },
        })) : null;
        conditionsChanged = true;
      }
      if (containsSubject !== undefined) {
        conditions.subjectContains = containsSubject ? [containsSubject] : null;
        conditionsChanged = true;
      }
      if (containsBody !== undefined) {
        conditions.bodyContains = containsBody ? [containsBody] : null;
        conditionsChanged = true;
      }
      if (containsBodyOrSubject !== undefined) {
        conditions.bodyOrSubjectContains = containsBodyOrSubject ? [containsBodyOrSubject] : null;
        conditionsChanged = true;
      }
      if (importance !== undefined) {
        conditions.importance = importance || null;
        conditionsChanged = true;
      }
      if (hasAttachments !== undefined) {
        conditions.hasAttachments = hasAttachments;
        conditionsChanged = true;
      }
      if (sentToMe !== undefined) {
        conditions.sentToMe = sentToMe;
        conditionsChanged = true;
      }
      if (ccMe !== undefined) {
        conditions.ccMe = ccMe;
        conditionsChanged = true;
      }
      if (messageSizeOver !== undefined) {
        conditions.messageSizeOver = messageSizeOver ? parseInt(messageSizeOver) : null;
        conditionsChanged = true;
      }
      if (messageSizeUnder !== undefined) {
        conditions.messageSizeUnder = messageSizeUnder ? parseInt(messageSizeUnder) : null;
        conditionsChanged = true;
      }
      if (categories !== undefined) {
        conditions.categories = categories ? categories.split(',').map(c => c.trim()) : null;
        conditionsChanged = true;
      }

      if (conditionsChanged) {
        // Remove null values
        Object.keys(conditions).forEach(key => {
          if (conditions[key] === null) delete conditions[key];
        });
        patch.conditions = Object.keys(conditions).length > 0 ? conditions : {};
      }
    }

    // Build actions
    if (clearActions) {
      patch.actions = {};
    } else {
      const actions = { ...rule.actions };
      let actionsChanged = false;

      if (moveToFolder !== undefined) {
        actions.moveToFolder = moveToFolder || null;
        actionsChanged = true;
      }
      if (copyToFolder !== undefined) {
        actions.copyToFolder = copyToFolder || null;
        actionsChanged = true;
      }
      if (deleteAction !== undefined) {
        actions.delete = deleteAction;
        actionsChanged = true;
      }
      if (markAsRead !== undefined) {
        actions.markAsRead = markAsRead;
        actionsChanged = true;
      }
      if (markImportance !== undefined) {
        actions.markImportance = markImportance || null;
        actionsChanged = true;
      }
      if (forwardTo !== undefined) {
        actions.forwardTo = forwardTo ? forwardTo.split(',').map(a => ({
          emailAddress: { address: a.trim(), name: a.trim() },
        })) : null;
        actionsChanged = true;
      }
      if (redirectTo !== undefined) {
        actions.redirectTo = redirectTo ? redirectTo.split(',').map(a => ({
          emailAddress: { address: a.trim(), name: a.trim() },
        })) : null;
        actionsChanged = true;
      }
      if (assignCategories !== undefined) {
        actions.assignCategories = assignCategories ? assignCategories.split(',').map(c => c.trim()) : null;
        actionsChanged = true;
      }
      if (stopProcessingRules !== undefined) {
        actions.stopProcessingRules = stopProcessingRules;
        actionsChanged = true;
      }
      if (flagForAction !== undefined) {
        actions.flagForAction = flagForAction || null;
        actionsChanged = true;
      }

      if (actionsChanged) {
        // Remove null values
        Object.keys(actions).forEach(key => {
          if (actions[key] === null) delete actions[key];
        });
        patch.actions = Object.keys(actions).length > 0 ? actions : {};
      }
    }

    await client.api(`me/mailFolders/inbox/messageRules/${rule.id}`).patch(patch);

    return makeResponse(`Advanced rule "${ruleName}" updated successfully.`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error updating advanced rule: ${error.message}`);
  }
}

/**
 * Get rule details with all conditions and actions
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetRule(args) {
  const { ruleName } = args;

  if (!ruleName) {
    return makeErrorResponse('Rule name is required.');
  }

  try {
    const client = await getGraphClient();
    const rules = await client.api('me/mailFolders/inbox/messageRules').get();
    const rule = rules.value?.find(r => r.displayName === ruleName);

    if (!rule) {
      return makeErrorResponse(`Rule "${ruleName}" not found.`);
    }

    const structured = {
      id: rule.id,
      name: rule.displayName,
      sequence: rule.sequence,
      isEnabled: rule.isEnabled,
      conditions: rule.conditions ? {
        fromAddresses: rule.conditions.fromAddresses?.map(a => a.emailAddress?.address).filter(Boolean) || [],
        senderAddressIncludes: rule.conditions.senderAddressIncludes?.map(a => a.emailAddress?.address).filter(Boolean) || [],
        recipientAddressIncludes: rule.conditions.recipientAddressIncludes?.map(a => a.emailAddress?.address).filter(Boolean) || [],
        subjectContains: rule.conditions.subjectContains || [],
        bodyContains: rule.conditions.bodyContains || [],
        bodyOrSubjectContains: rule.conditions.bodyOrSubjectContains || [],
        importance: rule.conditions.importance,
        hasAttachments: rule.conditions.hasAttachments,
        sentToMe: rule.conditions.sentToMe,
        ccMe: rule.conditions.ccMe,
        messageSizeOver: rule.conditions.messageSizeOver,
        messageSizeUnder: rule.conditions.messageSizeUnder,
        categories: rule.conditions.categories || [],
      } : null,
      actions: rule.actions ? {
        moveToFolder: rule.actions.moveToFolder,
        copyToFolder: rule.actions.copyToFolder,
        delete: rule.actions.delete,
        markAsRead: rule.actions.markAsRead,
        markImportance: rule.actions.markImportance,
        forwardTo: rule.actions.forwardTo?.map(a => a.emailAddress?.address).filter(Boolean) || [],
        redirectTo: rule.actions.redirectTo?.map(a => a.emailAddress?.address).filter(Boolean) || [],
        assignCategories: rule.actions.assignCategories || [],
        stopProcessingRules: rule.actions.stopProcessingRules,
        flagForAction: rule.actions.flagForAction,
      } : null,
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

    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error getting rule: ${error.message}`);
  }
}

/**
 * Test/simulate a rule against an email
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleTestRule(args) {
  const { ruleName, emailId } = args;

  if (!ruleName) {
    return makeErrorResponse('Rule name is required.');
  }
  if (!emailId) {
    return makeErrorResponse('Email ID is required to test against.');
  }

  try {
    const client = await getGraphClient();

    // Get the rule
    const rules = await client.api('me/mailFolders/inbox/messageRules').get();
    const rule = rules.value?.find(r => r.displayName === ruleName);

    if (!rule) {
      return makeErrorResponse(`Rule "${ruleName}" not found.`);
    }

    // Get the email
    const email = await client.api(`me/messages/${emailId}`).select(
      'id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,bodyPreview,body,hasAttachments,importance,isRead,categories,size'
    ).get();

    // Simulate rule matching
    const matches = [];
    const mismatches = [];

    if (rule.conditions) {
      const c = rule.conditions;

      // Check fromAddresses
      if (c.fromAddresses?.length) {
        const senderEmail = email.from?.emailAddress?.address?.toLowerCase();
        const senderName = email.from?.emailAddress?.name?.toLowerCase();
        const matched = c.fromAddresses.some(a =>
          a.emailAddress?.address?.toLowerCase() === senderEmail ||
          a.emailAddress?.name?.toLowerCase() === senderName
        );
        if (matched) matches.push('fromAddresses'); else mismatches.push('fromAddresses');
      }

      // Check senderAddressIncludes
      if (c.senderAddressIncludes?.length) {
        const senderEmail = email.from?.emailAddress?.address?.toLowerCase();
        const senderName = email.from?.emailAddress?.name?.toLowerCase();
        const matched = c.senderAddressIncludes.some(a =>
          senderEmail?.includes(a.emailAddress?.address?.toLowerCase()) ||
          senderName?.includes(a.emailAddress?.name?.toLowerCase())
        );
        if (matched) matches.push('senderAddressIncludes'); else mismatches.push('senderAddressIncludes');
      }

      // Check containsSubject
      if (c.subjectContains?.length) {
        const subject = email.subject?.toLowerCase() || '';
        const matched = c.subjectContains.some(s => subject.includes(s.toLowerCase()));
        if (matched) matches.push('subjectContains'); else mismatches.push('subjectContains');
      }

      // Check containsBody
      if (c.bodyContains?.length) {
        const body = (email.bodyPreview || '').toLowerCase();
        const matched = c.bodyContains.some(s => body.includes(s.toLowerCase()));
        if (matched) matches.push('bodyContains'); else mismatches.push('bodyContains');
      }

      // Check importance
      if (c.importance) {
        const matched = email.importance === c.importance;
        if (matched) matches.push(`importance:${c.importance}`); else mismatches.push(`importance:${c.importance}`);
      }

      // Check hasAttachments
      if (c.hasAttachments !== undefined) {
        const matched = email.hasAttachments === c.hasAttachments;
        if (matched) matches.push('hasAttachments'); else mismatches.push('hasAttachments');
      }

      // Check sentToMe
      if (c.sentToMe !== undefined) {
        const me = await client.api('me').select('mail,userPrincipalName').get();
        const myEmail = (me.mail || me.userPrincipalName || '').toLowerCase();
        const toEmails = email.toRecipients?.map(r => r.emailAddress?.address?.toLowerCase()) || [];
        const matched = toEmails.includes(myEmail);
        if (matched) matches.push('sentToMe'); else mismatches.push('sentToMe');
      }

      // Check ccMe
      if (c.ccMe !== undefined) {
        const me = await client.api('me').select('mail,userPrincipalName').get();
        const myEmail = (me.mail || me.userPrincipalName || '').toLowerCase();
        const ccEmails = email.ccRecipients?.map(r => r.emailAddress?.address?.toLowerCase()) || [];
        const matched = ccEmails.includes(myEmail);
        if (matched) matches.push('ccMe'); else mismatches.push('ccMe');
      }

      // Check messageSizeOver
      if (c.messageSizeOver !== undefined) {
        const matched = (email.size || 0) > c.messageSizeOver;
        if (matched) matches.push('messageSizeOver'); else mismatches.push('messageSizeOver');
      }

      // Check messageSizeUnder
      if (c.messageSizeUnder !== undefined) {
        const matched = (email.size || 0) < c.messageSizeUnder;
        if (matched) matches.push('messageSizeUnder'); else mismatches.push('messageSizeUnder');
      }

      // Check categories
      if (c.categories?.length) {
        const emailCats = email.categories || [];
        const matched = c.categories.some(cat => emailCats.includes(cat));
        if (matched) matches.push('categories'); else mismatches.push('categories');
      }
    }

    const wouldMatch = mismatches.length === 0 && (matches.length > 0 || !rule.conditions);

    const structured = {
      ruleName: rule.displayName,
      ruleId: rule.id,
      emailId: email.id,
      emailSubject: email.subject,
      wouldMatch,
      matchedConditions: matches,
      unmatchedConditions: mismatches,
      actionsThatWouldRun: rule.actions ? Object.keys(rule.actions).filter(k => rule.actions[k] !== null && rule.actions[k] !== false) : [],
    };

    const textFallback = [
      `Rule Test: "${rule.displayName}"`,
      `Email: ${email.subject}`,
      `Would Match: ${wouldMatch ? 'YES ✓' : 'NO ✗'}`,
      '',
      `Matched Conditions (${matches.length}):`,
      matches.length > 0 ? matches.map(m => `  ✓ ${m}`).join('\n') : '  (no conditions or all matched)',
      '',
      `Unmatched Conditions (${mismatches.length}):`,
      mismatches.length > 0 ? mismatches.map(m => `  ✗ ${m}`).join('\n') : '  None',
      '',
      `Actions That Would Run:`,
      structured.actionsThatWouldRun.length > 0
        ? structured.actionsThatWouldRun.map(a => `  → ${a}`).join('\n')
        : '  None',
    ].join('\n');

    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error testing rule: ${error.message}`);
  }
}

/**
 * Enable/disable multiple rules at once
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleBulkEnableRules(args) {
  const { ruleNames, isEnabled } = args;

  if (!ruleNames || !Array.isArray(ruleNames) || ruleNames.length === 0) {
    return makeErrorResponse('ruleNames array is required.');
  }
  if (isEnabled === undefined) {
    return makeErrorResponse('isEnabled (true/false) is required.');
  }

  try {
    const client = await getGraphClient();
    const rules = await client.api('me/mailFolders/inbox/messageRules').get();
    const allRules = rules.value || [];

    const results = [];
    const errors = [];

    for (const ruleName of ruleNames) {
      const rule = allRules.find(r => r.displayName === ruleName);
      if (!rule) {
        errors.push({ ruleName, error: 'Rule not found' });
        continue;
      }

      try {
        await client.api(`me/mailFolders/inbox/messageRules/${rule.id}`).patch({ isEnabled });
        results.push({ ruleName, status: 'success' });
      } catch (e) {
        errors.push({ ruleName, error: e.message });
      }
    }

    const action = isEnabled ? 'Enabled' : 'Disabled';

    const structured = {
      action,
      successful: results.length,
      failed: errors.length,
      results,
      errors,
    };

    const textFallback = [
      `Bulk ${action} Rules`,
      `Successful: ${results.length}`,
      `Failed: ${errors.length}`,
      '',
      results.length > 0 ? 'Success:' : '',
      ...results.map(r => `  ✓ ${r.ruleName}`),
      errors.length > 0 ? 'Failed:' : '',
      ...errors.map(e => `  ✗ ${e.ruleName}: ${e.error}`),
    ].filter(Boolean).join('\n');

    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error bulk enabling rules: ${error.message}`);
  }
}

module.exports = {
  handleCreateAdvancedRule,
  handleUpdateAdvancedRule,
  handleGetRule,
  handleTestRule,
  handleBulkEnableRules,
};