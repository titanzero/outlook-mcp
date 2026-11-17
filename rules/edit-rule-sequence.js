const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { getInboxRules } = require('./list');

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

module.exports = {
  handleEditRuleSequence,
};
