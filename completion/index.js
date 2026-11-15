/**
 * Completion handlers for autocompletion in tool arguments
 * These are registered as completion callbacks on resource templates
 */
const { getGraphClient } = require('../utils/graph-client');
const { getAllFolders } = require('../email/folder-utils');
const { isAuthError } = require('../utils/response-helpers');

/**
 * Get folder completions
 * @param {string} value - Partial folder name/path
 * @returns {Promise<Array<{value: string, label: string}>>}
 */
async function getFolderCompletions(value) {
  try {
    const folders = await getAllFolders();
    return folders
      .map(f => f.path || f.displayName)
      .filter(name => name.toLowerCase().includes(value.toLowerCase()))
      .slice(0, 20)
      .map(s => ({ value: s, label: s }));
  } catch (error) {
    if (isAuthError(error)) {
      return [];
    }
    console.error('Folder completion error:', error);
    return [];
  }
}

/**
 * Get calendar completions
 * @param {string} value - Partial calendar name
 * @returns {Promise<Array<{value: string, label: string}>>}
 */
async function getCalendarCompletions(value) {
  try {
    const client = await getGraphClient();
    const response = await client.api('me/calendars').select('id,name').get();
    const calendars = response.value || [];

    return calendars
      .map(c => c.name)
      .filter(name => name.toLowerCase().includes(value.toLowerCase()))
      .slice(0, 20)
      .map(name => ({ value: name, label: name }));
  } catch (error) {
    if (isAuthError(error)) {
      return [];
    }
    console.error('Calendar completion error:', error);
    return [];
  }
}

/**
 * Get contact email completions
 * @param {string} value - Partial email/name
 * @returns {Promise<Array<{value: string, label: string}>>}
 */
async function getContactCompletions(value) {
  try {
    const client = await getGraphClient();
    const response = await client
      .api('me/contacts')
      .select('id,displayName,emailAddresses')
      .top(50)
      .get();

    const contacts = response.value || [];
    const suggestions = [];

    contacts.forEach(c => {
      c.emailAddresses?.forEach(e => {
        if (e.address.toLowerCase().includes(value.toLowerCase())) {
          suggestions.push({ value: e.address, label: `${c.displayName} <${e.address}>` });
        }
      });
      if (c.displayName.toLowerCase().includes(value.toLowerCase())) {
        suggestions.push({ value: c.displayName, label: c.displayName });
      }
    });

    return suggestions.slice(0, 20);
  } catch (error) {
    if (isAuthError(error)) {
      return [];
    }
    console.error('Contact completion error:', error);
    return [];
  }
}

/**
 * Get task list completions
 * @param {string} value - Partial task list name
 * @returns {Promise<Array<{value: string, label: string}>>}
 */
async function getTaskListCompletions(value) {
  try {
    const client = await getGraphClient();
    const response = await client.api('me/todo/lists').select('id,displayName').get();
    const lists = response.value || [];

    return lists
      .map(l => l.displayName)
      .filter(name => name.toLowerCase().includes(value.toLowerCase()))
      .slice(0, 20)
      .map(name => ({ value: name, label: name }));
  } catch (error) {
    if (isAuthError(error)) {
      return [];
    }
    console.error('Task list completion error:', error);
    return [];
  }
}

/**
 * Get category completions
 * @param {string} value - Partial category name
 * @returns {Promise<Array<{value: string, label: string}>>}
 */
async function getCategoryCompletions(value) {
  const commonCategories = [
    'Red Category',
    'Orange Category',
    'Yellow Category',
    'Green Category',
    'Blue Category',
    'Purple Category',
    'Pink Category',
    'Grey Category',
    'Important',
    'Urgent',
    'Follow Up',
    'Waiting',
    'Reference',
    'Personal',
    'Work',
    'Project',
    'Meeting',
    'Travel',
    'Finance',
    'Legal',
    'HR',
  ];

  return commonCategories
    .filter(cat => cat.toLowerCase().includes(value.toLowerCase()))
    .slice(0, 20)
    .map(cat => ({ value: cat, label: cat }));
}

/**
 * Get rule completions
 * @param {string} value - Partial rule name
 * @returns {Promise<Array<{value: string, label: string}>>}
 */
async function getRuleCompletions(value) {
  try {
    const client = await getGraphClient();
    const response = await client
      .api('me/mailFolders/inbox/messageRules')
      .select('id,displayName')
      .get();
    const rules = response.value || [];

    return rules
      .map(r => r.displayName)
      .filter(name => name.toLowerCase().includes(value.toLowerCase()))
      .slice(0, 20)
      .map(name => ({ value: name, label: name }));
  } catch (error) {
    if (isAuthError(error)) {
      return [];
    }
    console.error('Rule completion error:', error);
    return [];
  }
}

/**
 * Get importance level completions
 * @param {string} value - Partial importance
 * @returns {Promise<Array<{value: string, label: string}>>}
 */
async function getImportanceCompletions(value) {
  const levels = ['low', 'normal', 'high'];
  return levels.filter(l => l.startsWith(value.toLowerCase())).map(l => ({ value: l, label: l }));
}

/**
 * Get email folder completions (well-known)
 * @param {string} value - Partial folder name
 * @returns {Promise<Array<{value: string, label: string}>>}
 */
async function getEmailFolderCompletions(value) {
  const wellKnown = [
    'inbox',
    'sent',
    'drafts',
    'deleted',
    'junk',
    'archive',
    'outbox',
    'conversationhistory',
  ];
  return wellKnown
    .filter(f => f.startsWith(value.toLowerCase()))
    .map(f => ({ value: f, label: f }));
}

/**
 * Register completion callbacks on resource templates
 * This is called from each resource module to add completions
 * @param {McpServer} server - The MCP server instance
 */
function registerCompletionCallbacks(_server) {
  // The completions are registered on resource templates via the ResourceTemplate class
  // when using server.resource() with a template that has completeCallback
  // This function is a placeholder for future use if needed
  console.error('Completion callbacks available for resource templates');
}

/**
 * Get event completions
 * @param {string} value - Partial event subject
 * @returns {Promise<Array<{value: string, label: string}>>}
 */
async function getEventCompletions(value) {
  try {
    const client = await getGraphClient();
    const response = await client
      .api('me/events')
      .select('id,subject')
      .filter(`startswith(subject,'${value.replace(/'/g, "''")}')`)
      .top(20)
      .get();
    const events = response.value || [];

    return events
      .map(e => ({ value: e.id, label: `${e.subject} [${e.id.substring(0, 8)}...]` }))
      .slice(0, 20);
  } catch (error) {
    if (isAuthError(error)) {
      return [];
    }
    console.error('Event completion error:', error);
    return [];
  }
}

module.exports = {
  getFolderCompletions,
  getCalendarCompletions,
  getContactCompletions,
  getTaskListCompletions,
  getCategoryCompletions,
  getRuleCompletions,
  getImportanceCompletions,
  getEmailFolderCompletions,
  getEventCompletions,
  registerCompletionCallbacks,
};
