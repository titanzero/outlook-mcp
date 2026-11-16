/**
 * Mail tips tools for Outlook MCP server
 * Check recipient availability, out-of-office, mailbox full, etc.
 */
const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * Get mail tips for recipients
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetMailTips(args) {
  const { emailAddresses } = args;
  
  if (!emailAddresses || !Array.isArray(emailAddresses) || emailAddresses.length === 0) {
    return makeErrorResponse('emailAddresses array is required (at least one recipient).');
  }
  
  if (emailAddresses.length > 20) {
    return makeErrorResponse('Maximum 20 email addresses allowed per request.');
  }
  
  try {
    const client = await getGraphClient();
    
    const response = await client.api('me/getMailTips').post({
      emailAddresses: emailAddresses.map(addr => ({ address: addr })),
      mailTipsOptions: 'All',
    });
    
    const tips = response.value || [];
    
    const structured = tips.map(t => ({
      emailAddress: t.emailAddress?.address,
      automaticReplies: t.automaticReplies ? {
        message: t.automaticReplies.message,
        scheduledStart: t.automaticReplies.scheduledStartTime?.dateTime,
        scheduledEnd: t.automaticReplies.scheduledEndTime?.dateTime,
      } : null,
      mailboxFull: t.mailboxFull,
      customMailTip: t.customMailTip,
      externalMemberCount: t.externalMemberCount,
      totalMemberCount: t.totalMemberCount,
      maxMessageSize: t.maxMessageSize,
      deliveryRestricted: t.deliveryRestricted,
      isModerated: t.isModerated,
      invalidRecipient: t.invalidRecipient,
      scope: t.scope,
    }));
    
    const textFallback = [
      `Mail Tips (${tips.length} recipients):`,
      '',
      tips.map((t, i) => {
        const lines = [`${i + 1}. ${t.emailAddress?.address}`];
        if (t.automaticReplies) {
          lines.push(`   Auto-reply: ${t.automaticReplies.message}`);
          if (t.automaticReplies.scheduledStart) {
            lines.push(`   From: ${t.automaticReplies.scheduledStart} to ${t.automaticReplies.scheduledEnd}`);
          }
        }
        if (t.mailboxFull) lines.push('   ⚠ Mailbox full');
        if (t.deliveryRestricted) lines.push('   ⚠ Delivery restricted');
        if (t.isModerated) lines.push('   ⚠ Moderated recipient');
        if (t.invalidRecipient) lines.push('   ⚠ Invalid recipient');
        if (t.customMailTip) lines.push(`   Custom: ${t.customMailTip}`);
        if (t.maxMessageSize) lines.push(`   Max message size: ${Math.round(t.maxMessageSize / 1024)} KB`);
        if (t.totalMemberCount) lines.push(`   Group: ${t.totalMemberCount} members (${t.externalMemberCount} external)`);
        return lines.join('\n');
      }).join('\n\n'),
    ].join('\n');
    
    return makeResponse(formatResponse({ mailTips: structured }, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error getting mail tips: ${error.message}`);
  }
}

/**
 * Get mailbox usage statistics
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetMailboxUsage(_args) {
  try {
    const client = await getGraphClient();
    
    // Get mailbox settings with quota info
    const settings = await client.api('me/mailboxSettings').get();
    
    // Try to get mailbox usage from me/mailFolders
    let totalMessages = 0;
    let totalSize = 0;
    
    try {
      const foldersResponse = await client.api('me/mailFolders').select('totalItemCount').get();
      const folders = foldersResponse.value || [];
      totalMessages = folders.reduce((sum, f) => sum + (f.totalItemCount || 0), 0);
    } catch (e) {
      // Ignore
    }
    
    const structured = {
      displayName: settings.displayName,
      userPrincipalName: settings.userPrincipalName,
      mail: settings.mail,
      totalMessages,
      timeZone: settings.timeZone,
      language: settings.language?.displayName,
      automaticReplies: settings.automaticRepliesSetting?.status,
      workingHours: settings.workingHours ? {
        daysOfWeek: settings.workingHours.daysOfWeek,
        startTime: settings.workingHours.startTime,
        endTime: settings.workingHours.endTime,
        timeZone: settings.workingHours.timeZone?.name,
      } : null,
    };
    
    const textFallback = [
      `Mailbox Usage`,
      `Name: ${structured.displayName || 'N/A'}`,
      `Email: ${structured.mail || structured.userPrincipalName || 'N/A'}`,
      `Total Messages (estimate): ${totalMessages.toLocaleString()}`,
      `Time Zone: ${structured.timeZone}`,
      `Language: ${structured.language || 'Not set'}`,
      `Auto-replies: ${structured.automaticReplies || 'disabled'}`,
      '',
      `Working Hours:`,
      structured.workingHours
        ? `  Days: ${structured.workingHours.daysOfWeek.join(', ')}\n  Time: ${structured.workingHours.startTime} - ${structured.workingHours.endTime} (${structured.workingHours.timeZone})`
        : '  Not configured',
    ].join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error getting mailbox usage: ${error.message}`);
  }
}

/**
 * Get user profile with extended info
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetUserProfile(_args) {
  try {
    const client = await getGraphClient();
    
    const user = await client.api('me').select(
      'id,displayName,givenName,surname,userPrincipalName,mail,mailNickname,jobTitle,officeLocation,department,companyName,businessPhones,mobilePhone,preferredLanguage,employeeId,employeeHireDate,employeeOrgData,createdDateTime'
    ).get();
    
    const structured = {
      id: user.id,
      displayName: user.displayName,
      givenName: user.givenName,
      surname: user.surname,
      userPrincipalName: user.userPrincipalName,
      mail: user.mail,
      mailNickname: user.mailNickname,
      jobTitle: user.jobTitle,
      officeLocation: user.officeLocation,
      department: user.department,
      companyName: user.companyName,
      businessPhones: user.businessPhones || [],
      mobilePhone: user.mobilePhone,
      preferredLanguage: user.preferredLanguage,
      employeeId: user.employeeId,
      employeeHireDate: user.employeeHireDate,
      createdDateTime: user.createdDateTime,
    };
    
    const textFallback = [
      `User Profile`,
      `Name: ${user.displayName}`,
      `Given Name: ${user.givenName || 'N/A'}`,
      `Surname: ${user.surname || 'N/A'}`,
      `UPN: ${user.userPrincipalName}`,
      `Email: ${user.mail || 'N/A'}`,
      `Mail Nickname: ${user.mailNickname || 'N/A'}`,
      `Job Title: ${user.jobTitle || 'N/A'}`,
      `Office: ${user.officeLocation || 'N/A'}`,
      `Department: ${user.department || 'N/A'}`,
      `Company: ${user.companyName || 'N/A'}`,
      `Phones: ${user.businessPhones?.join(', ') || 'N/A'}`,
      `Mobile: ${user.mobilePhone || 'N/A'}`,
      `Language: ${user.preferredLanguage || 'N/A'}`,
      `Employee ID: ${user.employeeId || 'N/A'}`,
      `Hire Date: ${user.employeeHireDate || 'N/A'}`,
      `Created: ${user.createdDateTime || 'N/A'}`,
    ].join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error getting user profile: ${error.message}`);
  }
}

module.exports = {
  handleGetMailTips,
  handleGetMailboxUsage,
  handleGetUserProfile,
};