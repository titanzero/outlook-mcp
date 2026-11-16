// Force text mode so existing assertions match plain-text output
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleGetReminders, handleSnoozeReminder, handleDismissReminder, handleSetReminder } = require('../../calendar/reminders');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('Event Reminders Tools', () => {
  let mockClient;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockClient = {
      api: jest.fn(() => mockClient),
      get: jest.fn(),
      patch: jest.fn(),
      filter: jest.fn(() => mockClient),
      select: jest.fn(() => mockClient),
      orderby: jest.fn(() => mockClient),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('handleGetReminders', () => {
    test('should get reminders in date range', async () => {
      const mockEvents = [
        { id: 'event-1', subject: 'Meeting', start: { dateTime: '2024-01-15T10:00:00Z', timeZone: 'UTC' }, end: { dateTime: '2024-01-15T11:00:00Z', timeZone: 'UTC' }, location: { displayName: 'Room A' }, isReminderOn: true, reminderMinutesBeforeStart: 15, organizer: { emailAddress: { address: 'org@company.com' } } },
      ];

      mockClient.get.mockResolvedValue({ value: mockEvents });

      const result = await handleGetReminders({
        startTime: '2024-01-15T00:00:00Z',
        endTime: '2024-01-16T00:00:00Z',
      });

      expect(result.content[0].text).toContain('Event Reminders (1)');
      expect(result.content[0].text).toContain('Meeting');
      expect(result.content[0].text).toContain('15 min before');
    });

    test('should return error when startTime/endTime missing', async () => {
      const result = await handleGetReminders({ startTime: '2024-01-15T00:00:00Z' });

      expect(result.content[0].text).toBe('startTime and endTime are required (ISO 8601 format).');
    });
  });

  describe('handleSnoozeReminder', () => {
    test('should snooze reminder', async () => {
      // Use a future date that's far enough for snoozing
      const futureDate = new Date(Date.now() + 24 * 60 * 60 * 1000).toISOString();
      const mockEvent = { id: 'event-1', start: { dateTime: futureDate } };

      mockClient.get.mockResolvedValue(mockEvent);
      mockClient.patch.mockResolvedValue({});

      const result = await handleSnoozeReminder({ eventId: 'event-1', snoozeMinutes: 5 });

      expect(result.content[0].text).toContain('Reminder snoozed');
    });

    test('should dismiss reminder if snooze time is after event start', async () => {
      const mockEvent = { id: 'event-1', start: { dateTime: new Date(Date.now() + 60000).toISOString() } }; // 1 min from now

      mockClient.get.mockResolvedValue(mockEvent);
      mockClient.patch.mockResolvedValue({});

      const result = await handleSnoozeReminder({ eventId: 'event-1', snoozeMinutes: 10 });

      expect(result.content[0].text).toContain('Reminder dismissed');
    });
  });

  describe('handleDismissReminder', () => {
    test('should dismiss reminder', async () => {
      mockClient.patch.mockResolvedValue({});

      const result = await handleDismissReminder({ eventId: 'event-1' });

      expect(result.content[0].text).toBe('Reminder dismissed successfully.');
    });
  });

  describe('handleSetReminder', () => {
    test('should set reminder', async () => {
      mockClient.patch.mockResolvedValue({});

      const result = await handleSetReminder({ eventId: 'event-1', minutesBeforeStart: 30 });

      expect(result.content[0].text).toBe('Reminder set to 30 minutes before event start.');
    });

    test('should return error when minutesBeforeStart is 0 (treated as invalid)', async () => {
      const result = await handleSetReminder({ eventId: 'event-1', minutesBeforeStart: 0 });

      expect(result.content[0].text).toContain('minutesBeforeStart is required');
    });

    test('should return error for invalid minutes', async () => {
      const result = await handleSetReminder({ eventId: 'event-1', minutesBeforeStart: -1 });

      expect(result.content[0].text).toContain('minutesBeforeStart is required');
    });
  });
});