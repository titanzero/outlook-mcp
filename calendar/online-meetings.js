/**
 * Online meeting providers tools for Outlook MCP server
 */
const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * List online meeting providers
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListOnlineMeetingProviders(_args) {
  try {
    const client = await getGraphClient();
    
    const response = await client.api('me/onlineMeetingProviders').get();
    const providers = response.value || [];
    
    const structured = providers.map(p => ({
      id: p.id,
      displayName: p.displayName,
      type: p.type,
      meetingUrl: p.meetingUrl,
      logoUrl: p.logoUrl,
      termsUrl: p.termsUrl,
      privacyUrl: p.privacyUrl,
    }));
    
    const textFallback = [
      `Online Meeting Providers (${providers.length}):`,
      '',
      providers.length === 0
        ? '  No online meeting providers configured.'
        : providers.map((p, i) => 
            `${i + 1}. ${p.displayName} (${p.type})\n   ID: ${p.id}\n   Meeting URL: ${p.meetingUrl || 'N/A'}${p.logoUrl ? `\n   Logo: ${p.logoUrl}` : ''}${p.termsUrl ? `\n   Terms: ${p.termsUrl}` : ''}${p.privacyUrl ? `\n   Privacy: ${p.privacyUrl}` : ''}`
          ).join('\n\n'),
    ].join('\n');
    
    return makeResponse(formatResponse({ providers: structured }, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    if (error.statusCode === 404) {
      return makeErrorResponse('Online meeting providers API not available.');
    }
    return makeErrorResponse(`Error listing online meeting providers: ${error.message}`);
  }
}

/**
 * Create an online meeting
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateOnlineMeeting(args) {
  const { 
    subject, 
    startTime, 
    endTime, 
    attendees = [], 
    isEntryExitAnnounced = true,
    allowedPresenters = 'organization',
    allowAttendeeToEnableCamera = true,
    allowAttendeeToEnableMic = true,
    allowMeetingChat = true,
    autoAdmittedUsers = 'everyone',
    description 
  } = args;
  
  if (!subject) {
    return makeErrorResponse('Subject is required.');
  }
  if (!startTime || !endTime) {
    return makeErrorResponse('startTime and endTime are required (ISO 8601 format).');
  }
  
  try {
    const client = await getGraphClient();
    
    const meeting = {
      subject,
      startTime: { dateTime: startTime, timeZone: 'UTC' },
      endTime: { dateTime: endTime, timeZone: 'UTC' },
      participants: {
        attendees: attendees.map(email => ({
          identity: { user: { email } },
          role: 'attendee',
        })),
      },
      isEntryExitAnnounced,
      allowedPresenters,
      allowAttendeeToEnableCamera,
      allowAttendeeToEnableMic,
      allowMeetingChat,
      autoAdmittedUsers,
    };
    
    if (description) {
      meeting.description = description;
    }
    
    const created = await client.api('me/onlineMeetings').post(meeting);
    
    const structured = {
      id: created.id,
      subject: created.subject,
      startTime: created.startTime?.dateTime,
      endTime: created.endTime?.dateTime,
      joinUrl: created.joinUrl,
      joinWebUrl: created.joinWebUrl,
      chatInfo: created.chatInfo,
      videoTeleconferenceId: created.videoTeleconferenceId,
      audioConferencing: created.audioConferencing,
    };
    
    const textFallback = [
      `Online Meeting Created`,
      `ID: ${created.id}`,
      `Subject: ${created.subject}`,
      `Start: ${created.startTime?.dateTime}`,
      `End: ${created.endTime?.dateTime}`,
      `Join URL: ${created.joinUrl || created.joinWebUrl}`,
      created.chatInfo?.threadId ? `Chat Thread: ${created.chatInfo.threadId}` : '',
      created.videoTeleconferenceId ? `Video Conf ID: ${created.videoTeleconferenceId}` : '',
      created.audioConferencing?.tollNumber ? `Dial-in: ${created.audioConferencing.tollNumber}` : '',
    ].filter(Boolean).join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error creating online meeting: ${error.message}`);
  }
}

/**
 * Get online meeting details
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetOnlineMeeting(args) {
  const { meetingId } = args;
  
  if (!meetingId) {
    return makeErrorResponse('Meeting ID is required.');
  }
  
  try {
    const client = await getGraphClient();
    
    const meeting = await client.api(`me/onlineMeetings/${meetingId}`).get();
    
    const structured = {
      id: meeting.id,
      subject: meeting.subject,
      startTime: meeting.startTime?.dateTime,
      endTime: meeting.endTime?.dateTime,
      joinUrl: meeting.joinUrl,
      joinWebUrl: meeting.joinWebUrl,
      chatInfo: meeting.chatInfo,
      videoTeleconferenceId: meeting.videoTeleconferenceId,
      audioConferencing: meeting.audioConferencing,
      isEntryExitAnnounced: meeting.isEntryExitAnnounced,
      allowedPresenters: meeting.allowedPresenters,
      allowAttendeeToEnableCamera: meeting.allowAttendeeToEnableCamera,
      allowAttendeeToEnableMic: meeting.allowAttendeeToEnableMic,
      allowMeetingChat: meeting.allowMeetingChat,
      autoAdmittedUsers: meeting.autoAdmittedUsers,
      description: meeting.description,
    };
    
    const textFallback = [
      `Online Meeting: ${meeting.subject}`,
      `ID: ${meeting.id}`,
      `Start: ${meeting.startTime?.dateTime}`,
      `End: ${meeting.endTime?.dateTime}`,
      `Join URL: ${meeting.joinUrl || meeting.joinWebUrl}`,
      meeting.chatInfo?.threadId ? `Chat Thread: ${meeting.chatInfo.threadId}` : '',
      meeting.videoTeleconferenceId ? `Video Conf ID: ${meeting.videoTeleconferenceId}` : '',
      meeting.audioConferencing?.tollNumber ? `Dial-in: ${meeting.audioConferencing.tollNumber}` : '',
      meeting.description ? `\nDescription: ${meeting.description}` : '',
    ].filter(Boolean).join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error getting online meeting: ${error.message}`);
  }
}

module.exports = {
  handleListOnlineMeetingProviders,
  handleCreateOnlineMeeting,
  handleGetOnlineMeeting,
};