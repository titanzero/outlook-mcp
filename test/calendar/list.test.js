/**
 * Tests for calendar list events functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const handleListEvents = require('../../calendar/list');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');
jest.mock('../../config', () => ({
  DEFAULT_TIMEZONE: 'Central European Standard Time',
  MAX_RESULT_COUNT: 100,
  CALENDAR_SELECT_FIELDS: 'id,subject,start,end,location,attendees,bodyPreview,isAllDay',
}));

function createMockClient(events) {
  const mockChain = {
    query: jest.fn().mockReturnThis(),
    top: jest.fn().mockReturnThis(),
    orderby: jest.fn().mockReturnThis(),
    select: jest.fn().mockReturnThis(),
    get: jest.fn().mockResolvedValue({ value: events }),
  };
  return {
    api: jest.fn().mockReturnValue(mockChain),
  };
}

describe('handleListEvents', () => {
  let mockClient;

  const mockEvents = [
    {
      id: 'event-1',
      subject: 'Team Meeting',
      start: { dateTime: '2024-01-15T10:00:00', timeZone: 'UTC' },
      end: { dateTime: '2024-01-15T11:00:00', timeZone: 'UTC' },
      location: { displayName: 'Conference Room A' },
      attendees: [
        { emailAddress: { name: 'Alice', address: 'alice@example.com' }, status: { response: 'accepted' } },
        { emailAddress: { name: 'Bob', address: 'bob@example.com' }, status: { response: 'tentative' } },
      ],
      bodyPreview: 'Weekly team sync',
      isAllDay: false,
    },
    {
      id: 'event-2',
      subject: 'All Day Event',
      start: { dateTime: '2024-01-16T00:00:00', timeZone: 'UTC' },
      end: { dateTime: '2024-01-17T00:00:00', timeZone: 'UTC' },
      location: { displayName: 'Office' },
      attendees: [],
      bodyPreview: 'Company holiday',
      isAllDay: true,
    },
  ];

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockClient = createMockClient(mockEvents);
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('successful event listing', () => {
    test('should list events with default count', async () => {
      const result = await handleListEvents({});

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(mockClient.api).toHaveBeenCalledWith('me/calendarView');
      expect(result.content[0].text).toContain('Found 2 events:');
      expect(result.content[0].text).toContain('Team Meeting');
      expect(result.content[0].text).toContain('All Day Event');
      expect(result.content[0].text).toContain('Conference Room A');
      expect(result.content[0].text).toContain('Alice (accepted)');
      expect(result.content[0].text).toContain('Bob (tentative)');
      expect(result.content[0].text).toContain('Weekly team sync');
    });

    test('should respect custom count parameter', async () => {
      await handleListEvents({ count: 5 });

      expect(mockClient.api).toHaveBeenCalledWith('me/calendarView');
      expect(mockClient.api().top).toHaveBeenCalledWith(5);
    });

    test('should cap count at MAX_RESULT_COUNT', async () => {
      await handleListEvents({ count: 500 });

      expect(mockClient.api().top).toHaveBeenCalledWith(100);
    });

    test('should handle events without location', async () => {
      const eventsNoLocation = [
        {
          id: 'event-3',
          subject: 'No Location Event',
          start: { dateTime: '2024-01-15T10:00:00', timeZone: 'UTC' },
          end: { dateTime: '2024-01-15T11:00:00', timeZone: 'UTC' },
          location: null,
          attendees: [],
          bodyPreview: '',
          isAllDay: false,
        },
      ];

      mockClient = createMockClient(eventsNoLocation);
      getGraphClient.mockResolvedValue(mockClient);

      const result = await handleListEvents({});

      expect(result.content[0].text).toContain('No Location Event');
      expect(result.content[0].text).toContain('No location');
    });

    test('should handle events without attendees', async () => {
      const eventsNoAttendees = [
        {
          id: 'event-4',
          subject: 'Solo Event',
          start: { dateTime: '2024-01-15T10:00:00', timeZone: 'UTC' },
          end: { dateTime: '2024-01-15T11:00:00', timeZone: 'UTC' },
          location: { displayName: 'Home' },
          attendees: null,
          bodyPreview: 'Personal time',
          isAllDay: false,
        },
      ];

      mockClient = createMockClient(eventsNoAttendees);
      getGraphClient.mockResolvedValue(mockClient);

      const result = await handleListEvents({});

      expect(result.content[0].text).toContain('Solo Event');
      expect(result.content[0].text).not.toContain('Attendees:');
    });

    test('should handle all-day events correctly', async () => {
      const result = await handleListEvents({});

      expect(result.content[0].text).toContain('(All Day)');
    });

    test('should return "No calendar events found" when empty', async () => {
      mockClient = createMockClient([]);
      getGraphClient.mockResolvedValue(mockClient);

      const result = await handleListEvents({});

      expect(result.content[0].text).toBe('No calendar events found.');
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleListEvents({});

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle Graph API error', async () => {
      const mockChain = {
        query: jest.fn().mockReturnThis(),
        top: jest.fn().mockReturnThis(),
        orderby: jest.fn().mockReturnThis(),
        select: jest.fn().mockReturnThis(),
        get: jest.fn().mockRejectedValue(new Error('Graph API Error')),
      };
      mockClient = { api: jest.fn().mockReturnValue(mockChain) };
      getGraphClient.mockResolvedValue(mockClient);

      const result = await handleListEvents({});

      expect(result.content[0].text).toBe('Error listing events: Graph API Error');
    });
  });
});
