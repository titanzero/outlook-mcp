/**
 * Task recurrence tools for Outlook MCP server
 */
const { getGraphClient } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * Set recurrence on a task
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSetTaskRecurrence(args) {
  const { listId, taskId, pattern, range } = args;
  
  if (!listId) {
    return makeErrorResponse('listId is required (from list-task-lists).');
  }
  if (!taskId) {
    return makeErrorResponse('taskId is required (from list-tasks).');
  }
  if (!pattern) {
    return makeErrorResponse('Recurrence pattern is required.');
  }
  if (!range) {
    return makeErrorResponse('Recurrence range is required.');
  }
  
  try {
    const client = await getGraphClient();
    
    const recurrence = { pattern, range };
    
    await client.api(`me/todo/lists/${listId}/tasks/${taskId}`).patch({ recurrence });
    
    return makeResponse(`Task recurrence set successfully.`);
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error setting task recurrence: ${error.message}`);
  }
}

/**
 * Get recurrence for a task
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetTaskRecurrence(args) {
  const { listId, taskId } = args;
  
  if (!listId) {
    return makeErrorResponse('listId is required (from list-task-lists).');
  }
  if (!taskId) {
    return makeErrorResponse('taskId is required (from list-tasks).');
  }
  
  try {
    const client = await getGraphClient();
    
    const task = await client.api(`me/todo/lists/${listId}/tasks/${taskId}`).select('id,title,recurrence').get();
    
    if (!task.recurrence) {
      return makeResponse('This task has no recurrence pattern.');
    }
    
    const pattern = task.recurrence.pattern;
    const range = task.recurrence.range;
    
    const structured = {
      taskId: task.id,
      taskTitle: task.title,
      recurrence: {
        pattern: {
          type: pattern.type,
          interval: pattern.interval,
          month: pattern.month,
          dayOfMonth: pattern.dayOfMonth,
          daysOfWeek: pattern.daysOfWeek,
          firstDayOfWeek: pattern.firstDayOfWeek,
          index: pattern.index,
        },
        range: {
          type: range.type,
          startDate: range.startDate,
          endDate: range.endDate,
          recurrenceCount: range.recurrenceCount,
        },
      },
    };
    
    // Build human-readable description
    let patternDesc = '';
    switch (pattern.type) {
      case 'daily':
        patternDesc = `Every ${pattern.interval} day(s)`;
        break;
      case 'weekly':
        const days = pattern.daysOfWeek?.join(', ') || 'weekly';
        patternDesc = `Every ${pattern.interval} week(s) on ${days}`;
        break;
      case 'monthly':
        patternDesc = pattern.dayOfMonth
          ? `Every ${pattern.interval} month(s) on day ${pattern.dayOfMonth}`
          : `Every ${pattern.interval} month(s) on the ${pattern.index} ${pattern.daysOfWeek?.[0]}`;
        break;
      case 'yearly':
        patternDesc = `Every year on ${pattern.month} ${pattern.dayOfMonth}`;
        break;
      case 'absoluteMonthly':
        patternDesc = `Every ${pattern.interval} month(s) on day ${pattern.dayOfMonth}`;
        break;
      case 'relativeMonthly':
        patternDesc = `Every ${pattern.interval} month(s) on the ${pattern.index} ${pattern.daysOfWeek?.[0]}`;
        break;
      default:
        patternDesc = pattern.type;
    }
    
    let rangeDesc = '';
    switch (range.type) {
      case 'noEnd':
        rangeDesc = `Starting ${range.startDate}, no end date`;
        break;
      case 'endDate':
        rangeDesc = `From ${range.startDate} to ${range.endDate}`;
        break;
      case 'numbered':
        rangeDesc = `Starting ${range.startDate}, ${range.recurrenceCount} occurrences`;
        break;
    }
    
    const textFallback = [
      `Task Recurrence: ${task.title}`,
      `Task ID: ${task.id}`,
      '',
      `Pattern: ${patternDesc}`,
      `Range: ${rangeDesc}`,
    ].join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error getting task recurrence: ${error.message}`);
  }
}

/**
 * Remove recurrence from a task
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleRemoveTaskRecurrence(args) {
  const { listId, taskId } = args;
  
  if (!listId) {
    return makeErrorResponse('listId is required.');
  }
  if (!taskId) {
    return makeErrorResponse('taskId is required.');
  }
  
  try {
    const client = await getGraphClient();
    
    await client.api(`me/todo/lists/${listId}/tasks/${taskId}`).patch({ recurrence: null });
    
    return makeResponse('Task recurrence removed successfully.');
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error removing task recurrence: ${error.message}`);
  }
}

/**
 * Recurrence pattern helper - returns examples
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleRecurrenceExamples(_args) {
  const examples = `
TASK RECURRENCE PATTERN EXAMPLES

PATTERN TYPES:
1. DAILY
   {
     "type": "daily",
     "interval": 1
   }
   → Every day

2. WEEKLY
   {
     "type": "weekly",
     "interval": 1,
     "daysOfWeek": ["monday", "wednesday", "friday"]
   }
   → Every Monday, Wednesday, Friday

3. MONTHLY (by day of month)
   {
     "type": "monthly",
     "interval": 1,
     "dayOfMonth": 15
   }
   → 15th of every month

4. MONTHLY (by weekday)
   {
     "type": "monthly",
     "interval": 1,
     "index": "first",
     "daysOfWeek": ["monday"]
   }
   → First Monday of every month

5. YEARLY
   {
     "type": "yearly",
     "interval": 1,
     "month": 1,
     "dayOfMonth": 1
   }
   → January 1st every year

RANGE TYPES:
1. NO END
   {
     "type": "noEnd",
     "startDate": "2024-01-01"
   }
   → Repeats forever from start date

2. END DATE
   {
     "type": "endDate",
     "startDate": "2024-01-01",
     "endDate": "2024-12-31"
   }
   → Repeats until end date

3. NUMBERED (count)
   {
     "type": "numbered",
     "startDate": "2024-01-01",
     "recurrenceCount": 10
   }
   → 10 occurrences from start date

VALID VALUES:
- daysOfWeek: monday, tuesday, wednesday, thursday, friday, saturday, sunday
- index: first, second, third, fourth, last
- month: 1-12
- dayOfMonth: 1-31
- interval: 1-999

USAGE:
Use set-task-recurrence with pattern and range objects.
`;

  return makeResponse(examples.trim());
}

module.exports = {
  handleSetTaskRecurrence,
  handleGetTaskRecurrence,
  handleRemoveTaskRecurrence,
  handleRecurrenceExamples,
};