/**
 * Tests for mailbox out-of-office functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleSetOutOfOffice } = require('../../mailbox/index');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('handleSetOutOfOffice', () => {
  let mockClient;
  let mockPatch;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockPatch = jest.fn().mockResolvedValue({});
    mockClient = {
      api: jest.fn(() => ({
        patch: mockPatch,
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('successful out-of-office setting', () => {
    test('should enable out-of-office', async () => {
      const result = await handleSetOutOfOffice({
        status: 'enabled',
        internalMessage: 'I am out of office.',
        externalMessage: 'I will reply when I return.',
      });

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(mockClient.api).toHaveBeenCalledWith('me/mailboxSettings');
      expect(mockPatch).toHaveBeenCalledWith({
        automaticRepliesSetting: {
          status: 'enabled',
          internalReplyMessage: 'I am out of office.',
          externalReplyMessage: 'I will reply when I return.',
        },
      });

      expect(result.content[0].text).toBe('Out-of-office set to "enabled" successfully.');
    });

    test('should schedule out-of-office with start and end times', async () => {
      await handleSetOutOfOffice({
        status: 'scheduled',
        startTime: '2024-12-20T09:00:00Z',
        endTime: '2025-01-05T17:00:00Z',
        internalMessage: 'Out of office for holidays.',
      });

      expect(mockPatch).toHaveBeenCalledWith({
        automaticRepliesSetting: {
          status: 'scheduled',
          scheduledStartDateTime: { dateTime: '2024-12-20T09:00:00Z', timeZone: 'UTC' },
          scheduledEndDateTime: { dateTime: '2025-01-05T17:00:00Z', timeZone: 'UTC' },
          internalReplyMessage: 'Out of office for holidays.',
        },
      });
    });

    test('should disable out-of-office', async () => {
      const result = await handleSetOutOfOffice({ status: 'disabled' });

      expect(mockPatch).toHaveBeenCalledWith({
        automaticRepliesSetting: { status: 'disabled' },
      });

      expect(result.content[0].text).toBe('Out-of-office set to "disabled" successfully.');
    });

    test('should handle only internal message', async () => {
      await handleSetOutOfOffice({
        status: 'enabled',
        internalMessage: 'Internal only message.',
      });

      expect(mockPatch).toHaveBeenCalledWith({
        automaticRepliesSetting: {
          status: 'enabled',
          internalReplyMessage: 'Internal only message.',
        },
      });
    });

    test('should handle only external message', async () => {
      await handleSetOutOfOffice({
        status: 'enabled',
        externalMessage: 'External only message.',
      });

      expect(mockPatch).toHaveBeenCalledWith({
        automaticRepliesSetting: {
          status: 'enabled',
          externalReplyMessage: 'External only message.',
        },
      });
    });

    test('should handle no messages', async () => {
      const result = await handleSetOutOfOffice({ status: 'enabled' });

      expect(mockPatch).toHaveBeenCalledWith({
        automaticRepliesSetting: { status: 'enabled' },
      });

      expect(result.content[0].text).toBe('Out-of-office set to "enabled" successfully.');
    });
  });

  describe('validation', () => {
    test('should return error when status is missing', async () => {
      const result = await handleSetOutOfOffice({});

      expect(result.content[0].text).toContain('status is required');
      expect(result.isError).toBe(true);
      expect(getGraphClient).not.toHaveBeenCalled();
    });

    test('should return error when status is invalid', async () => {
      const result = await handleSetOutOfOffice({ status: 'invalid' });

      expect(result.content[0].text).toContain('Invalid status. Use one of: enabled, scheduled, disabled.');
    });

    test('should return error when scheduled but startTime missing', async () => {
      const result = await handleSetOutOfOffice({
        status: 'scheduled',
        endTime: '2025-01-05T17:00:00Z',
      });

      expect(result.content[0].text).toContain('startTime and endTime are required when status is "scheduled"');
    });

    test('should return error when scheduled but endTime missing', async () => {
      const result = await handleSetOutOfOffice({
        status: 'scheduled',
        startTime: '2024-12-20T09:00:00Z',
      });

      expect(result.content[0].text).toContain('startTime and endTime are required when status is "scheduled"');
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleSetOutOfOffice({ status: 'enabled' });

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle Graph API error', async () => {
      mockPatch.mockRejectedValue(new Error('Graph API Error'));

      const result = await handleSetOutOfOffice({ status: 'enabled' });

      expect(result.content[0].text).toBe('Error setting out-of-office: Graph API Error');
    });
  });
});
