/**
 * Tests for calendar list-calendars functionality
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const handleListCalendars = require('../../calendar/list-calendars');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('handleListCalendars', () => {
  let mockClient;

  const mockCalendars = [
    {
      id: 'cal-1',
      name: 'Personal',
      color: 'auto',
      isDefaultCalendar: true,
      canEdit: true,
      owner: { address: 'user@example.com' },
    },
    {
      id: 'cal-2',
      name: 'Work',
      color: 'lightBlue',
      isDefaultCalendar: false,
      canEdit: true,
      owner: { address: 'user@example.com' },
    },
    {
      id: 'cal-3',
      name: 'Shared Calendar',
      color: 'lightGreen',
      isDefaultCalendar: false,
      canEdit: false,
      owner: { address: 'other@example.com' },
    },
  ];

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    const mockChain = {
      select: jest.fn().mockReturnThis(),
      get: jest.fn().mockResolvedValue({ value: mockCalendars }),
    };
    mockClient = {
      api: jest.fn().mockReturnValue(mockChain),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('successful calendar listing', () => {
    test('should list all calendars', async () => {
      const result = await handleListCalendars({});

      expect(getGraphClient).toHaveBeenCalledTimes(1);
      expect(mockClient.api).toHaveBeenCalledWith('me/calendars');
      expect(mockClient.api().select).toHaveBeenCalledWith('id,name,color,isDefaultCalendar,canEdit,owner');

      expect(result.content[0].text).toContain('Calendars (3):');
      expect(result.content[0].text).toContain('Personal [default] (id: cal-1)');
      expect(result.content[0].text).toContain('Work (id: cal-2)');
      expect(result.content[0].text).toContain('Shared Calendar [read-only] (id: cal-3)');
    });

    test('should return structured data', async () => {
      const result = await handleListCalendars({});

      // The response should include structured data for TOON
      expect(result.content[0].text).toBeDefined();
    });
  });

  describe('empty results', () => {
    test('should return appropriate message when no calendars found', async () => {
      mockClient = {
        api: jest.fn(() => ({
          select: jest.fn(() => ({
            get: jest.fn().mockResolvedValue({ value: [] }),
          })),
        })),
      };
      getGraphClient.mockResolvedValue(mockClient);

      const result = await handleListCalendars({});

      expect(result.content[0].text).toBe('No calendars found.');
    });
  });

  describe('error handling', () => {
    test('should handle authentication error', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      getGraphClient.mockRejectedValue(authError);

      const result = await handleListCalendars({});

      expect(result.content[0].text).toBe('Authentication required');
    });

    test('should handle Graph API error', async () => {
      mockClient = {
        api: jest.fn(() => ({
          select: jest.fn(() => ({
            get: jest.fn().mockRejectedValue(new Error('Graph API Error')),
          })),
        })),
      };
      getGraphClient.mockResolvedValue(mockClient);

      const result = await handleListCalendars({});

      expect(result.content[0].text).toBe('Error listing calendars: Graph API Error');
    });
  });
});
