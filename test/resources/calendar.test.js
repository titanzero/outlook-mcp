/**
 * Tests for calendar resources
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { registerCalendarResources } = require('../../calendar/resources');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');
jest.mock('@modelcontextprotocol/sdk/server/mcp.js', () => ({
  ResourceTemplate: jest.fn().mockImplementation((uri, options) => ({ uri, options })),
}));

describe('registerCalendarResources', () => {
  let mockServer;
  let mockClient;

  const mockEvents = [
    {
      id: 'event-1',
      subject: 'Team Meeting',
      start: { dateTime: '2024-01-15T10:00:00', timeZone: 'UTC' },
      end: { dateTime: '2024-01-15T11:00:00', timeZone: 'UTC' },
      location: { displayName: 'Conference Room A' },
      attendees: [{ emailAddress: { name: 'Alice', address: 'alice@example.com' } }],
      bodyPreview: 'Weekly sync',
      isAllDay: false,
    },
  ];

  const mockCalendars = [
    { id: 'cal-1', name: 'Personal', color: 'auto', isDefaultCalendar: true, canEdit: true },
    { id: 'cal-2', name: 'Work', color: 'lightBlue', isDefaultCalendar: false, canEdit: true },
  ];

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});

    mockClient = {
      api: jest.fn(() => ({
        query: jest.fn().mockReturnThis(),
        top: jest.fn().mockReturnThis(),
        orderby: jest.fn().mockReturnThis(),
        select: jest.fn().mockReturnThis(),
        get: jest.fn().mockResolvedValue({ value: mockEvents }),
      })),
    };

    getGraphClient.mockResolvedValue(mockClient);

    mockServer = { resource: jest.fn() };
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('should register calendar resources', () => {
    registerCalendarResources(mockServer);
    // Calendar resources: events, calendars, event detail
    expect(mockServer.resource).toHaveBeenCalled();
  });
});
