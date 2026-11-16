/**
 * Working hours tools for Outlook MCP server
 */
const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * Get working hours
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetWorkingHours(_args) {
  try {
    const client = await getGraphClient();
    
    const settings = await client.api('me/mailboxSettings').select('workingHours').get();
    
    const wh = settings.workingHours;
    
    if (!wh) {
      return makeResponse('Working hours not configured.');
    }
    
    const dayLabels = {
      monday: 'Monday',
      tuesday: 'Tuesday',
      wednesday: 'Wednesday',
      thursday: 'Thursday',
      friday: 'Friday',
      saturday: 'Saturday',
      sunday: 'Sunday',
    };
    
    const structured = {
      daysOfWeek: wh.daysOfWeek || [],
      startTime: wh.startTime,
      endTime: wh.endTime,
      timeZone: wh.timeZone?.name || 'UTC',
    };
    
    const days = (wh.daysOfWeek || []).map(d => dayLabels[d] || d).join(', ');
    
    const textFallback = [
      `Working Hours`,
      `Days: ${days || 'Not set'}`,
      `Time: ${wh.startTime || 'Not set'} - ${wh.endTime || 'Not set'}`,
      `Time Zone: ${wh.timeZone?.name || 'UTC'}`,
    ].join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error getting working hours: ${error.message}`);
  }
}

/**
 * Set working hours
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSetWorkingHours(args) {
  const { daysOfWeek, startTime, endTime, timeZone } = args;
  
  if (!daysOfWeek || !Array.isArray(daysOfWeek) || daysOfWeek.length === 0) {
    return makeErrorResponse('daysOfWeek array is required (e.g., ["monday", "tuesday", "wednesday", "thursday", "friday"])');
  }
  if (!startTime) {
    return makeErrorResponse('startTime is required (HH:mm format)');
  }
  if (!endTime) {
    return makeErrorResponse('endTime is required (HH:mm format)');
  }
  
  const validDays = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'];
  for (const day of daysOfWeek) {
    if (!validDays.includes(day.toLowerCase())) {
      return makeErrorResponse(`Invalid day: ${day}. Valid days: ${validDays.join(', ')}`);
    }
  }
  
  // Validate time format HH:mm
  const timeRegex = /^([01]?[0-9]|2[0-3]):[0-5][0-9]$/;
  if (!timeRegex.test(startTime)) {
    return makeErrorResponse('startTime must be in HH:mm format (e.g., "09:00")');
  }
  if (!timeRegex.test(endTime)) {
    return makeErrorResponse('endTime must be in HH:mm format (e.g., "17:00")');
  }
  
  try {
    const client = await getGraphClient();
    
    const workingHours = {
      daysOfWeek: daysOfWeek.map(d => d.toLowerCase()),
      startTime,
      endTime,
      timeZone: { name: timeZone || 'UTC' },
    };
    
    await client.api('me/mailboxSettings').patch({ workingHours });
    
    const dayLabels = {
      monday: 'Monday',
      tuesday: 'Tuesday',
      wednesday: 'Wednesday',
      thursday: 'Thursday',
      friday: 'Friday',
      saturday: 'Saturday',
      sunday: 'Sunday',
    };
    
    const days = daysOfWeek.map(d => dayLabels[d.toLowerCase()] || d).join(', ');
    
    return makeResponse(`Working hours set successfully.\nDays: ${days}\nTime: ${startTime} - ${endTime}\nTime Zone: ${timeZone || 'UTC'}`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error setting working hours: ${error.message}`);
  }
}

/**
 * Clear working hours
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleClearWorkingHours(_args) {
  try {
    const client = await getGraphClient();
    
    await client.api('me/mailboxSettings').patch({ workingHours: null });
    
    return makeResponse('Working hours cleared successfully.');
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error clearing working hours: ${error.message}`);
  }
}

module.exports = {
  handleGetWorkingHours,
  handleSetWorkingHours,
  handleClearWorkingHours,
};