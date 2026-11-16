/**
 * Tests for mailbox settings functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleGetMailboxSettings } = require('../../mailbox/index');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('handleGetMailboxSettings', () => {
  let mockClient;

  const mockSettings = {
    timeZone: 'Central European Standard Time',
    language: { displayName: 'English (United States)', locale: 'en-US' },
    dateFormat: 'dd/MM/yyyy',
    timeFormat: 'HH:mm',
    automaticRepliesSetting: {
      status: 'disabled',
      scheduledStartDateTime: null,
      scheduledEndDateTime: null,
      internalReplyMessage: null,
      externalReplyMessage: null,
    },
  };

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockClient = {
      api: jest.fn(() => ({
        get: jest.fn().mockResolvedValue(mockSettings),
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('successful mailbox settings retrieval', () => {
    test('should get mailbox settings', async () => {
      const result = await handleGetMailboxSettings({});

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(mockClient.api).toHaveBeenCalledWith('me/mailboxSettings');

      expect(result.content[0].text).toContain('Timezone: Central European Standard Time');
      expect(result.content[0].text).toContain('Language: English (United States)');
      expect(result.content[0].text).toContain('Auto-replies: disabled');
    });

    test('should include auto-reply details when scheduled', async () => {
      const scheduledSettings = {
        ...mockSettings,
        automaticRepliesSetting: {
          status: 'scheduled',
          scheduledStartDateTime: { dateTime: '2024-12-20T09:00:00', timeZone: 'UTC' },
          scheduledEndDateTime: { dateTime: '2025-01-05T17:00:00', timeZone: 'UTC' },
          internalReplyMessage: 'I am out of office.',
          externalReplyMessage: 'I will reply when I return.',
        },
      };

      mockClient = {
        api: jest.fn(() => ({
          get: jest.fn().mockResolvedValue(scheduledSettings),
        })),
      };
      getGraphClient.mockResolvedValue(mockClient);

      const result = await handleGetMailboxSettings({});

      expect(result.content[0].text).toContain('Auto-replies: scheduled');
      expect(result.content[0].text).toContain('From: 2024-12-20T09:00:00');
      expect(result.content[0].text).toContain('To:   2025-01-05T17:00:00');
      expect(result.content[0].text).toContain('Internal: I am out of office.');
      expect(result.content[0].text).toContain('External: I will reply when I return.');
    });

    test('should handle missing language gracefully', async () => {
      const noLanguageSettings = { ...mockSettings, language: null };
      mockClient = {
        api: jest.fn(() => ({
          get: jest.fn().mockResolvedValue(noLanguageSettings),
        })),
      };
      getGraphClient.mockResolvedValue(mockClient);

      const result = await handleGetMailboxSettings({});

      expect(result.content[0].text).toContain('Language: not set');
    });

    test('should handle missing automaticRepliesSetting gracefully', async () => {
      const noArsSettings = { ...mockSettings, automaticRepliesSetting: null };
      mockClient = {
        api: jest.fn(() => ({
          get: jest.fn().mockResolvedValue(noArsSettings),
        })),
      };
      getGraphClient.mockResolvedValue(mockClient);

      const result = await handleGetMailboxSettings({});

      expect(result.content[0].text).toContain('Auto-replies: ');
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleGetMailboxSettings({});

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle Graph API error', async () => {
      mockClient = {
        api: jest.fn(() => ({
          get: jest.fn().mockRejectedValue(new Error('Graph API Error')),
        })),
      };
      getGraphClient.mockResolvedValue(mockClient);

      const result = await handleGetMailboxSettings({});

      expect(result.content[0].text).toBe('Error fetching mailbox settings: Graph API Error');
    });
  });
});
