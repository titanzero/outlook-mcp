/**
 * Mailbox resources for MCP
 * Exposes mailbox settings as addressable resources
 */
const { getGraphClient } = require('../utils/graph-client');
const { isAuthError } = require('../utils/response-helpers');

/**
 * Register mailbox resources with the MCP server
 * @param {McpServer} server - The MCP server instance
 */
function registerMailboxResources(server) {
  // Resource for mailbox settings
  server.resource(
    'outlook-mailbox-settings',
    'outlook://mailbox/settings',
    {
      title: 'Mailbox Settings',
      description:
        'Access mailbox settings including timezone, language, auto-replies, and working hours',
    },
    async uri => {
      try {
        const client = await getGraphClient();
        const settings = await client.api('me/mailboxSettings').get();

        const structured = {
          timeZone: settings.timeZone,
          language: settings.language?.displayName || null,
          dateFormat: settings.dateFormat || null,
          timeFormat: settings.timeFormat || null,
          automaticReplies: {
            status: settings.automaticRepliesSetting?.status,
            scheduledStart:
              settings.automaticRepliesSetting?.scheduledStartDateTime?.dateTime || null,
            scheduledEnd: settings.automaticRepliesSetting?.scheduledEndDateTime?.dateTime || null,
            internalMessage: settings.automaticRepliesSetting?.internalReplyMessage || null,
            externalMessage: settings.automaticRepliesSetting?.externalReplyMessage || null,
          },
          workingHours: settings.workingHours
            ? {
              daysOfWeek: settings.workingHours.daysOfWeek || [],
              startTime: settings.workingHours.startTime,
              endTime: settings.workingHours.endTime,
              timeZone: settings.workingHours.timeZone?.name || null,
            }
            : null,
          userPurpose: settings.userPurpose,
          delegateMeetingMessageDeliveryOptions: settings.delegateMeetingMessageDeliveryOptions,
        };

        const ars = structured.automaticReplies;
        const wh = structured.workingHours;

        const __textFallback = [
          `Timezone: ${structured.timeZone}`,
          `Language: ${structured.language || 'not set'}`,
          `Date Format: ${structured.dateFormat || 'not set'}`,
          `Time Format: ${structured.timeFormat || 'not set'}`,
          '',
          `Auto-replies: ${ars.status}`,
          ars.scheduledStart ? `  From: ${ars.scheduledStart}` : '',
          ars.scheduledEnd ? `  To: ${ars.scheduledEnd}` : '',
          ars.internalMessage ? `  Internal: ${ars.internalMessage}` : '',
          ars.externalMessage ? `  External: ${ars.externalMessage}` : '',
          '',
          wh
            ? `Working Hours: ${wh.daysOfWeek?.join(', ')} ${wh.startTime} - ${wh.endTime} (${wh.timeZone || 'UTC'})`
            : 'Working Hours: not set',
        ]
          .filter(Boolean)
          .join('\n');

        return {
          contents: [
            {
              uri: uri.toString(),
              mimeType: 'application/json',
              text: JSON.stringify(structured, null, 2),
            },
          ],
        };
      } catch (error) {
        if (isAuthError(error)) {
          throw error;
        }
        throw new Error(`Error reading mailbox settings: ${error.message}`);
      }
    }
  );

  // Resource for user profile
  server.resource(
    'outlook-user-profile',
    'outlook://user/profile',
    {
      title: 'User Profile',
      description: "Access the authenticated user's profile information",
    },
    async uri => {
      try {
        const client = await getGraphClient();
        const user = await client
          .api('me')
          .select(
            'id,displayName,givenName,surname,userPrincipalName,mail,jobTitle,officeLocation,department,companyName,businessPhones,mobilePhone,preferredLanguage'
          )
          .get();

        const structured = {
          id: user.id,
          displayName: user.displayName,
          givenName: user.givenName,
          surname: user.surname,
          userPrincipalName: user.userPrincipalName,
          mail: user.mail,
          jobTitle: user.jobTitle,
          officeLocation: user.officeLocation,
          department: user.department,
          companyName: user.companyName,
          businessPhones: user.businessPhones || [],
          mobilePhone: user.mobilePhone,
          preferredLanguage: user.preferredLanguage,
        };

        const __textFallback = [
          `User: ${user.displayName}`,
          `ID: ${user.id}`,
          `UPN: ${user.userPrincipalName}`,
          `Email: ${user.mail || 'not set'}`,
          `Title: ${user.jobTitle || 'not set'}`,
          `Office: ${user.officeLocation || 'not set'}`,
          `Department: ${user.department || 'not set'}`,
          `Company: ${user.companyName || 'not set'}`,
          `Phones: ${user.businessPhones?.join(', ') || 'none'}`,
          `Mobile: ${user.mobilePhone || 'not set'}`,
          `Language: ${user.preferredLanguage || 'not set'}`,
        ].join('\n');

        return {
          contents: [
            {
              uri: uri.toString(),
              mimeType: 'application/json',
              text: JSON.stringify(structured, null, 2),
            },
          ],
        };
      } catch (error) {
        if (isAuthError(error)) {
          throw error;
        }
        throw new Error(`Error reading user profile: ${error.message}`);
      }
    }
  );
}

module.exports = { registerMailboxResources };
