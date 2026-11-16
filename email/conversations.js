/**
 * Conversation/thread view tool for Outlook MCP server
 */
const { getGraphClient, graphGetPaginated } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * List conversations (threaded view)
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListConversations(args) {
  const { folder = 'inbox', count = 20 } = args;
  
  try {
    const client = await getGraphClient();
    
    // Get conversations from a folder
    // Use the conversations endpoint (beta) or group by conversationId
    let endpoint;
    if (folder === 'inbox') {
      endpoint = 'me/messages';
    } else {
      endpoint = `me/mailFolders/${folder}/messages`;
    }
    
    // Get messages and group by conversationId
    const response = await graphGetPaginated(client, endpoint, {
      $select: 'id,conversationId,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments,importance,bodyPreview',
      $orderby: 'receivedDateTime desc',
      $top: Math.min(count * 5, 500), // Get more messages to group into conversations
    });
    
    const messages = response.value || [];
    
    // Group by conversationId
    const conversationMap = new Map();
    messages.forEach(msg => {
      const convId = msg.conversationId;
      if (!conversationMap.has(convId)) {
        conversationMap.set(convId, {
          conversationId: convId,
          subject: msg.subject,
          participants: new Set(),
          messageCount: 0,
          lastMessageDate: msg.receivedDateTime,
          hasUnread: false,
          hasAttachments: false,
          importance: 'normal',
          latestPreview: msg.bodyPreview,
          messages: [],
        });
      }
      const conv = conversationMap.get(convId);
      conv.messageCount++;
      if (msg.from?.emailAddress?.address) {
        conv.participants.add(`${msg.from.emailAddress.name} <${msg.from.emailAddress.address}>`);
      }
      msg.toRecipients?.forEach(r => {
        if (r.emailAddress?.address) {
          conv.participants.add(`${r.emailAddress.name} <${r.emailAddress.address}>`);
        }
      });
      if (!msg.isRead) conv.hasUnread = true;
      if (msg.hasAttachments) conv.hasAttachments = true;
      if (msg.importance === 'high') conv.importance = 'high';
      else if (msg.importance === 'low' && conv.importance !== 'high') conv.importance = 'low';
      conv.messages.push({
        id: msg.id,
        from: msg.from?.emailAddress,
        receivedDateTime: msg.receivedDateTime,
        isRead: msg.isRead,
        bodyPreview: msg.bodyPreview,
      });
    });
    
    // Convert to array and sort by last message date
    const conversations = Array.from(conversationMap.values())
      .sort((a, b) => new Date(b.lastMessageDate) - new Date(a.lastMessageDate))
      .slice(0, count)
      .map(conv => ({
        conversationId: conv.conversationId,
        subject: conv.subject,
        participants: Array.from(conv.participants),
        messageCount: conv.messageCount,
        lastMessageDate: conv.lastMessageDate,
        hasUnread: conv.hasUnread,
        hasAttachments: conv.hasAttachments,
        importance: conv.importance,
        latestPreview: conv.latestPreview,
      }));
    
    const structured = { conversations };
    
    const textFallback = [
      `Conversations in ${folder} (${conversations.length}):`,
      '',
      conversations.length === 0
        ? '  No conversations found.'
        : conversations.map((c, i) => {
            const unread = c.hasUnread ? '[UNREAD] ' : '';
            const attach = c.hasAttachments ? ' 📎' : '';
            const imp = c.importance === 'high' ? ' ❗' : c.importance === 'low' ? ' 🔽' : '';
            return `${i + 1}. ${unread}${c.subject}${attach}${imp}\n   Participants: ${c.participants.join(', ')}\n   Messages: ${c.messageCount} | Last: ${c.lastMessageDate}\n   Preview: ${c.latestPreview || '(empty)'}\n   ConversationId: ${c.conversationId}`;
          }).join('\n\n'),
    ].join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error listing conversations: ${error.message}`);
  }
}

/**
 * Get full conversation thread by conversationId
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetConversation(args) {
  const { conversationId, count = 50 } = args;
  
  if (!conversationId) {
    return makeErrorResponse('Conversation ID is required (from list-conversations).');
  }
  
  try {
    const client = await getGraphClient();
    
    // Get all messages in this conversation
    const response = await graphGetPaginated(client, 'me/messages', {
      $filter: `conversationId eq '${conversationId}'`,
      $select: 'id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,bodyPreview,body,hasAttachments,importance,isRead,isDraft,categories,flag',
      $orderby: 'receivedDateTime asc',
      $top: count,
    });
    
    const messages = response.value || [];
    
    const structured = {
      conversationId,
      messageCount: messages.length,
      messages: messages.map(m => {
        const sender = m.from ? m.from.emailAddress : { name: 'Unknown', address: 'unknown' };
        const toRecipients = m.toRecipients?.map(r => r.emailAddress?.address).filter(Boolean) || [];
        const ccRecipients = m.ccRecipients?.map(r => r.emailAddress?.address).filter(Boolean) || [];
        const bccRecipients = m.bccRecipients?.map(r => r.emailAddress?.address).filter(Boolean) || [];
        
        return {
          id: m.id,
          subject: m.subject,
          from: { name: sender.name, address: sender.address },
          to: toRecipients,
          cc: ccRecipients,
          bcc: bccRecipients,
          receivedDateTime: m.receivedDateTime,
          sentDateTime: m.sentDateTime,
          bodyPreview: m.bodyPreview,
          body: m.body?.content,
          bodyContentType: m.body?.contentType,
          hasAttachments: m.hasAttachments,
          importance: m.importance,
          isRead: m.isRead,
          isDraft: m.isDraft,
          categories: m.categories || [],
          flag: m.flag ? {
            status: m.flag.flagStatus,
            startDateTime: m.flag.startDateTime?.dateTime,
            dueDateTime: m.flag.dueDateTime?.dateTime,
            completedDateTime: m.flag.completedDateTime?.dateTime,
          } : null,
        };
      }),
    };
    
    const textFallback = [
      `Conversation Thread (${messages.length} messages)`,
      `ConversationId: ${conversationId}`,
      '',
      messages.map((m, i) => {
        const sender = m.from ? m.from.emailAddress : { name: 'Unknown', address: 'unknown' };
        const readStatus = m.isRead ? '' : '[UNREAD] ';
        const draft = m.isDraft ? '[DRAFT] ' : '';
        return `${i + 1}. ${readStatus}${draft}${m.receivedDateTime || m.sentDateTime} - ${sender.name} <${sender.address}>\n   Subject: ${m.subject}\n   To: ${m.to.join(', ') || 'none'}${m.cc.length ? `\n   CC: ${m.cc.join(', ')}` : ''}\n   Preview: ${m.bodyPreview || '(empty)'}\n   ID: ${m.id}`;
      }).join('\n\n'),
    ].join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error getting conversation: ${error.message}`);
  }
}

module.exports = {
  handleListConversations,
  handleGetConversation,
};