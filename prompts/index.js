/**
 * Prompts for common Outlook workflows
 * These provide pre-defined prompt templates that clients can invoke
 */
const { z } = require('zod');

/**
 * Register prompts with the MCP server
 * @param {McpServer} server - The MCP server instance
 */
function registerPrompts(server) {
  // Prompt: Summarize inbox
  server.prompt(
    'summarize-inbox',
    'Summarize recent emails in your inbox with key themes, action items, and urgent items',
    {
      folder: z.string().optional().describe('Folder to summarize (default: inbox)'),
      count: z
        .number()
        .int()
        .positive()
        .max(50)
        .optional()
        .describe('Number of emails to analyze (default: 20)'),
      timeRange: z
        .enum(['today', 'this-week', 'last-week', 'all'])
        .optional()
        .describe('Time range filter (default: today)'),
    },
    async ({ folder = 'inbox', count = 20, timeRange = 'today' }) => {
      const timeFilter = {
        today: 'received today',
        'this-week': 'received this week',
        'last-week': 'received last week',
        all: '',
      }[timeRange];

      return {
        messages: [
          {
            role: 'user',
            content: `Please summarize my ${folder} folder. Analyze the ${count} most recent emails ${timeFilter}. Provide:
1. Key themes/topics
2. Urgent items requiring action
3. Action items / follow-ups needed
4. Senders to prioritize
5. Any meetings or deadlines mentioned

Be concise and actionable.`,
          },
        ],
      };
    }
  );

  // Prompt: Schedule meeting
  server.prompt(
    'schedule-meeting',
    'Help schedule a meeting by checking availability and drafting the invite',
    {
      attendees: z.array(z.string()).describe('Attendee email addresses'),
      durationMinutes: z
        .number()
        .int()
        .positive()
        .optional()
        .describe('Meeting duration in minutes (default: 30)'),
      preferredTimes: z
        .array(z.string())
        .optional()
        .describe('Preferred date/time ranges in ISO 8601'),
      subject: z.string().optional().describe('Meeting subject/title'),
      description: z.string().optional().describe('Meeting description/agenda'),
    },
    async ({
      attendees,
      durationMinutes = 30,
      preferredTimes,
      subject = 'Meeting',
      description = '',
    }) => {
      const timeRanges = preferredTimes?.length
        ? preferredTimes.map(t => `- ${t}`).join('\n')
        : 'No specific preferences';

      return {
        messages: [
          {
            role: 'user',
            content: `I need to schedule a meeting with: ${attendees.join(', ')}
Duration: ${durationMinutes} minutes
Subject: ${subject}
Description: ${description}
Preferred times:
${timeRanges}

Please:
1. Check their free/busy schedules
2. Suggest 3 best time slots
3. Create the calendar event once I confirm`,
          },
        ],
      };
    }
  );

  // Prompt: Triage email
  server.prompt(
    'triage-email',
    'Analyze an email and suggest actions (reply, forward, flag, categorize, create task, etc.)',
    {
      emailId: z.string().describe('ID of the email to triage'),
    },
    async ({ emailId }) => {
      return {
        messages: [
          {
            role: 'user',
            content: `Please analyze email ${emailId} and suggest the best actions. Consider:
1. Does it need a reply? (urgent vs. can wait)
2. Should it be forwarded to someone?
3. Does it need follow-up? (flag with due date)
4. Should it be categorized?
5. Does it contain action items for a task?
6. Should it be moved to a folder?
7. Any calendar events to create?

Provide a prioritized action list with specific tool calls I can make.`,
          },
        ],
      };
    }
  );

  // Prompt: Daily briefing
  server.prompt(
    'daily-briefing',
    'Generate a daily briefing with calendar, tasks, and important emails',
    {
      date: z.string().optional().describe('Date for briefing in ISO 8601 (default: today)'),
      includeTasks: z.boolean().optional().describe('Include task lists (default: true)'),
      includeEmails: z.boolean().optional().describe('Include important emails (default: true)'),
    },
    async ({ date, includeTasks = true, includeEmails = true }) => {
      const targetDate = date || new Date().toISOString().split('T')[0];

      return {
        messages: [
          {
            role: 'user',
            content: `Generate a daily briefing for ${targetDate}. Include:
${includeTasks ? '1. Task lists and due tasks' : ''}
${includeEmails ? '2. Important unread emails (flagged, from key senders)' : ''}
3. Calendar events for the day
4. Any conflicts or preparation needed
5. Suggested focus areas

Format as a concise executive summary.`,
          },
        ],
      };
    }
  );

  // Prompt: Weekly review
  server.prompt(
    'weekly-review',
    'Generate a weekly review summary',
    {
      weekStart: z.string().optional().describe('Week start date in ISO 8601 (default: this week)'),
    },
    async ({ weekStart }) => {
      const start =
        weekStart || new Date(Date.now() - 7 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];

      return {
        messages: [
          {
            role: 'user',
            content: `Generate a weekly review for the week starting ${start}. Include:
1. Meetings attended / upcoming
2. Emails sent / received (key threads)
3. Tasks completed / created
4. Time allocation analysis
5. Accomplishments
6. Areas for improvement
7. Next week priorities

Format as a structured review.`,
          },
        ],
      };
    }
  );

  // Prompt: Follow-up draft
  server.prompt(
    'follow-up-draft',
    'Create a follow-up email draft based on a previous conversation',
    {
      emailId: z.string().describe('ID of the email to follow up on'),
      tone: z
        .enum(['professional', 'friendly', 'urgent', 'gentle'])
        .optional()
        .describe('Tone of follow-up (default: professional)'),
      context: z.string().optional().describe('Additional context for the follow-up'),
    },
    async ({ emailId, tone = 'professional', context = '' }) => {
      return {
        messages: [
          {
            role: 'user',
            content: `Create a follow-up email draft for email ${emailId}.
Tone: ${tone}
Additional context: ${context}

The draft should:
- Reference the original email
- Be polite but clear about what's needed
- Include a specific call to action
- Be concise`,
          },
        ],
      };
    }
  );

  // Prompt: Contact enrichment
  server.prompt(
    'enrich-contact',
    'Enrich a contact with additional information from emails and calendar',
    {
      contactId: z.string().describe('ID of the contact to enrich'),
    },
    async ({ contactId }) => {
      return {
        messages: [
          {
            role: 'user',
            content: `Analyze all interactions with contact ${contactId} and enrich their profile:
1. Recent email threads (key topics)
2. Meetings attended together
3. Shared attachments/documents
4. Communication frequency
5. Suggested fields to update (company, title, phone, notes)
6. Relationship strength indicator

Provide a summary with suggested updates.`,
          },
        ],
      };
    }
  );
}

module.exports = { registerPrompts };
