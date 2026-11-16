/**
 * Contact lists (distribution lists) tools for Outlook MCP server
 */
const { getGraphClient, graphGetPaginated } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * List contact lists (distribution lists in contacts)
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListContactLists(args) {
  const { folderId, count = 50 } = args;
  
  try {
    const client = await getGraphClient();
    
    let endpoint = 'me/contactFolders/Contacts/contacts';
    if (folderId) {
      endpoint = `me/contactFolders/${folderId}/contacts`;
    }
    
    // Filter for contact lists (they have member property)
    const response = await graphGetPaginated(client, endpoint, {
      $filter: "microsoft.graph.contact/members/any()",
      $select: 'id,displayName,members,emailAddresses,companyName',
      $top: Math.min(count, 100),
    });
    
    const lists = response.value || [];
    
    const structured = lists.map(l => ({
      id: l.id,
      name: l.displayName,
      email: l.emailAddresses?.[0]?.address,
      companyName: l.companyName,
      memberCount: l.members?.length || 0,
      members: l.members?.map(m => ({
        id: m.id,
        displayName: m.displayName,
        email: m.emailAddresses?.[0]?.address,
      })) || [],
    }));
    
    const textFallback = [
      `Contact Lists (${lists.length})${folderId ? ` in folder ${folderId}` : ''}:`,
      '',
      lists.length === 0
        ? '  No contact lists found.'
        : lists.map((l, i) => {
            const members = l.members?.map(m => `${m.displayName} <${m.email}>`).join(', ') || 'no members';
            return `${i + 1}. ${l.displayName}${l.email ? ` (${l.email})` : ''} - ${l.memberCount} members\n   Members: ${members}`;
          }).join('\n\n'),
    ].join('\n');
    
    return makeResponse(formatResponse({ contactLists: structured }, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error listing contact lists: ${error.message}`);
  }
}

/**
 * Create a contact list (distribution list)
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateContactList(args) {
  const { name, memberIds, email, folderId, companyName } = args;
  
  if (!name) {
    return makeErrorResponse('Contact list name is required.');
  }
  if (!memberIds || memberIds.length === 0) {
    return makeErrorResponse('At least one member ID is required (from list-contacts).');
  }
  
  try {
    const client = await getGraphClient();
    
    // Get member contacts to build the list
    const members = [];
    for (const memberId of memberIds) {
      try {
        const contact = await client.api(`me/contacts/${memberId}`).select('id,displayName,emailAddresses').get();
        members.push({
          '@odata.type': '#microsoft.graph.contact',
          id: contact.id,
          displayName: contact.displayName,
          emailAddresses: contact.emailAddresses,
        });
      } catch (e) {
        return makeErrorResponse(`Member contact ${memberId} not found.`);
      }
    }
    
    const contactList = {
      displayName: name,
      members,
    };
    
    if (email) {
      contactList.emailAddresses = [{ address: email, name: name }];
    }
    if (companyName) {
      contactList.companyName = companyName;
    }
    
    let endpoint = 'me/contactFolders/Contacts/contacts';
    if (folderId) {
      endpoint = `me/contactFolders/${folderId}/contacts`;
    }
    
    const created = await client.api(endpoint).post(contactList);
    
    return makeResponse(`Contact list "${name}" created successfully with ${members.length} members.\nID: ${created.id}`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error creating contact list: ${error.message}`);
  }
}

/**
 * Update contact list (add/remove members)
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleUpdateContactList(args) {
  const { listId, name, addMemberIds, removeMemberIds } = args;
  
  if (!listId) {
    return makeErrorResponse('Contact list ID is required (from list-contact-lists).');
  }
  
  try {
    const client = await getGraphClient();
    
    // Get current list
    const list = await client.api(`me/contacts/${listId}`).select('id,displayName,members').get();
    
    let members = list.members || [];
    
    // Remove members
    if (removeMemberIds && removeMemberIds.length > 0) {
      members = members.filter(m => !removeMemberIds.includes(m.id));
    }
    
    // Add members
    if (addMemberIds && addMemberIds.length > 0) {
      for (const memberId of addMemberIds) {
        // Check if already a member
        if (!members.some(m => m.id === memberId)) {
          try {
            const contact = await client.api(`me/contacts/${memberId}`).select('id,displayName,emailAddresses').get();
            members.push({
              '@odata.type': '#microsoft.graph.contact',
              id: contact.id,
              displayName: contact.displayName,
              emailAddresses: contact.emailAddresses,
            });
          } catch (e) {
            return makeErrorResponse(`Contact ${memberId} not found.`);
          }
        }
      }
    }
    
    const patch = { members };
    if (name) patch.displayName = name;
    
    await client.api(`me/contacts/${listId}`).patch(patch);
    
    return makeResponse(`Contact list updated successfully. Now has ${members.length} members.`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error updating contact list: ${error.message}`);
  }
}

/**
 * Delete contact list
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDeleteContactList(args) {
  const { listId } = args;
  
  if (!listId) {
    return makeErrorResponse('Contact list ID is required (from list-contact-lists).');
  }
  
  try {
    const client = await getGraphClient();
    
    await client.api(`me/contacts/${listId}`).delete();
    
    return makeResponse('Contact list deleted successfully.');
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error deleting contact list: ${error.message}`);
  }
}

/**
 * Send email to contact list
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSendEmailToContactList(args) {
  const { listId, subject, body, importance = 'normal' } = args;
  
  if (!listId) {
    return makeErrorResponse('Contact list ID is required (from list-contact-lists).');
  }
  if (!subject) {
    return makeErrorResponse('Subject is required.');
  }
  if (!body) {
    return makeErrorResponse('Body is required.');
  }
  
  try {
    const client = await getGraphClient();
    
    // Get the contact list with members
    const list = await client.api(`me/contacts/${listId}`).select('id,displayName,members,emailAddresses').get();
    
    // Collect all member email addresses
    const recipients = [];
    
    // Add list email if it has one
    if (list.emailAddresses?.[0]?.address) {
      recipients.push({
        emailAddress: {
          address: list.emailAddresses[0].address,
          name: list.displayName,
        },
      });
    }
    
    // Add member emails
    if (list.members) {
      for (const member of list.members) {
        const email = member.emailAddresses?.[0]?.address;
        if (email) {
          recipients.push({
            emailAddress: {
              address: email,
              name: member.displayName,
            },
          });
        }
      }
    }
    
    if (recipients.length === 0) {
      return makeErrorResponse('Contact list has no email addresses to send to.');
    }
    
    // Send email
    const message = {
      message: {
        subject,
        body: {
          contentType: 'HTML',
          content: body,
        },
        toRecipients: recipients,
        importance,
      },
      saveToSentItems: true,
    };
    
    await client.api('me/sendMail').post(message);
    
    return makeResponse(`Email sent to contact list "${list.displayName}" (${recipients.length} recipients).`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error sending email to contact list: ${error.message}`);
  }
}

module.exports = {
  handleListContactLists,
  handleCreateContactList,
  handleUpdateContactList,
  handleDeleteContactList,
  handleSendEmailToContactList,
};