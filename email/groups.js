/**
 * Distribution lists / Microsoft 365 Groups tool for Outlook MCP server
 */
const { getGraphClient, graphGetPaginated } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * List all groups the user is a member of
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListGroups(args) {
  const { count = 50 } = args;
  
  try {
    const client = await getGraphClient();
    
    const response = await client
      .api('me/memberOf/microsoft.graph.group')
      .top(Math.min(count, 100))
      .select('id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes,visibility,createdDateTime')
      .get();
    
    const groups = response.value || [];
    
    const structured = groups.map(g => ({
      id: g.id,
      name: g.displayName,
      description: g.description,
      email: g.mail,
      mailEnabled: g.mailEnabled,
      securityEnabled: g.securityEnabled,
      isUnifiedGroup: g.groupTypes?.includes('Unified') || false,
      visibility: g.visibility,
      createdDateTime: g.createdDateTime,
    }));
    
    const textFallback = [
      `Groups (${groups.length}):`,
      '',
      groups.length === 0
        ? '  No groups found.'
        : groups.map((g, i) => {
            const type = g.groupTypes?.includes('Unified') ? 'Microsoft 365 Group' : 
                         g.mailEnabled ? 'Distribution List' : 'Security Group';
            return `${i + 1}. ${g.displayName} (${type})\n   Email: ${g.mail || 'none'}\n   ID: ${g.id}`;
          }).join('\n\n'),
    ].join('\n');
    
    return makeResponse(formatResponse({ groups: structured }, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error listing groups: ${error.message}`);
  }
}

/**
 * Get group details including members
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetGroup(args) {
  const { groupId } = args;
  
  if (!groupId) {
    return makeErrorResponse('Group ID is required (from list-groups).');
  }
  
  try {
    const client = await getGraphClient();
    
    // Get group info
    const group = await client
      .api(`groups/${groupId}`)
      .select('id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes,visibility,createdDateTime,owners,members')
      .get();
    
    // Get members
    let members = [];
    try {
      const membersResponse = await graphGetPaginated(client, `groups/${groupId}/members`, {
        $select: 'id,displayName,userPrincipalName,mail,jobTitle,department',
      });
      members = membersResponse.value || [];
    } catch (e) {
      // Might not have permission to list members
    }
    
    // Get owners
    let owners = [];
    try {
      const ownersResponse = await graphGetPaginated(client, `groups/${groupId}/owners`, {
        $select: 'id,displayName,userPrincipalName,mail',
      });
      owners = ownersResponse.value || [];
    } catch (e) {
      // Might not have permission
    }
    
    const structured = {
      id: group.id,
      name: group.displayName,
      description: group.description,
      email: group.mail,
      mailEnabled: group.mailEnabled,
      securityEnabled: group.securityEnabled,
      isUnifiedGroup: group.groupTypes?.includes('Unified') || false,
      visibility: group.visibility,
      createdDateTime: group.createdDateTime,
      members: members.map(m => ({
        id: m.id,
        name: m.displayName,
        email: m.mail || m.userPrincipalName,
        jobTitle: m.jobTitle,
        department: m.department,
      })),
      owners: owners.map(o => ({
        id: o.id,
        name: o.displayName,
        email: o.mail || o.userPrincipalName,
      })),
    };
    
    const type = group.groupTypes?.includes('Unified') ? 'Microsoft 365 Group' : 
                 group.mailEnabled ? 'Distribution List' : 'Security Group';
    
    const textFallback = [
      `Group: ${group.displayName} (${type})`,
      `ID: ${group.id}`,
      `Email: ${group.mail || 'none'}`,
      `Description: ${group.description || 'none'}`,
      `Visibility: ${group.visibility}`,
      `Created: ${group.createdDateTime}`,
      '',
      `Owners (${owners.length}):`,
      owners.length > 0 ? owners.map(o => `  - ${o.displayName} <${o.mail || o.userPrincipalName}>`).join('\n') : '  (none)',
      '',
      `Members (${members.length}):`,
      members.length > 0 ? members.map(m => `  - ${m.displayName} <${m.mail || m.userPrincipalName}> (${m.jobTitle || 'no title'})`).join('\n') : '  (none or no permission)',
    ].join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error getting group: ${error.message}`);
  }
}

/**
 * Get group members (paginated)
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListGroupMembers(args) {
  const { groupId, count = 100 } = args;
  
  if (!groupId) {
    return makeErrorResponse('Group ID is required (from list-groups).');
  }
  
  try {
    const client = await getGraphClient();
    
    const response = await graphGetPaginated(client, `groups/${groupId}/members`, {
      $select: 'id,displayName,userPrincipalName,mail,jobTitle,department,officeLocation',
      $top: Math.min(count, 999),
    });
    
    const members = response.value || [];
    
    const structured = {
      groupId,
      members: members.map(m => ({
        id: m.id,
        name: m.displayName,
        email: m.mail || m.userPrincipalName,
        jobTitle: m.jobTitle,
        department: m.department,
        officeLocation: m.officeLocation,
      })),
    };
    
    const textFallback = [
      `Group Members (${members.length}):`,
      '',
      members.length === 0
        ? '  No members found.'
        : members.map((m, i) => `${i + 1}. ${m.displayName} <${m.mail || m.userPrincipalName}> (${m.jobTitle || 'no title'}, ${m.department || 'no dept'})`).join('\n'),
    ].join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error listing group members: ${error.message}`);
  }
}

/**
 * Send email to a group
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSendEmailToGroup(args) {
  const { groupId, subject, body, importance = 'normal' } = args;
  
  if (!groupId) {
    return makeErrorResponse('Group ID is required (from list-groups).');
  }
  if (!subject) {
    return makeErrorResponse('Subject is required.');
  }
  if (!body) {
    return makeErrorResponse('Body is required.');
  }
  
  try {
    const client = await getGraphClient();
    
    // Get group email address
    const group = await client.api(`groups/${groupId}`).select('mail,displayName').get();
    
    if (!group.mail) {
      return makeErrorResponse(`Group "${group.displayName}" does not have an email address.`);
    }
    
    // Send email to group
    const message = {
      message: {
        subject,
        body: {
          contentType: 'HTML',
          content: body,
        },
        toRecipients: [
          {
            emailAddress: {
              address: group.mail,
              name: group.displayName,
            },
          },
        ],
        importance,
      },
      saveToSentItems: true,
    };
    
    await client.api('me/sendMail').post(message);
    
    return makeResponse(`Email sent successfully to group "${group.displayName}" (${group.mail}).`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error sending email to group: ${error.message}`);
  }
}

module.exports = {
  handleListGroups,
  handleGetGroup,
  handleListGroupMembers,
  handleSendEmailToGroup,
};