// Force text mode so existing assertions match plain-text output
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { handleListOnlineMeetingProviders, handleCreateOnlineMeeting, handleGetOnlineMeeting } = require('../../calendar/online-meetings');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('Online Meetings Tools', () => {
  let mockClient;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    mockClient = {
      api: jest.fn(() => mockClient),
      get: jest.fn(),
      post: jest.fn(),
    };
    getGraphClient.mockResolvedValue(mockClient);
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('handleListOnlineMeetingProviders', () => {
    test('should list online meeting providers', async () => {
      const mockProviders = [
        { id: 'teams', displayName: 'Microsoft Teams', type: 'teamsForBusiness', meetingUrl: 'https://teams.microsoft.com', logoUrl: 'https://logo.url', termsUrl: 'https://terms.url', privacyUrl: 'https://privacy.url' },
      ];

      mockClient.get.mockResolvedValue({ value: mockProviders });

      const result = await handleListOnlineMeetingProviders({});

      expect(result.content[0].text).toContain('Online Meeting Providers (1)');
      expect(result.content[0].text).toContain('Microsoft Teams (teamsForBusiness)');
    });

    test('should handle API not available', async () => {
      const error = new Error('Not Found');
      error.statusCode = 404;
      mockClient.get.mockRejectedValue(error);

      const result = await handleListOnlineMeetingProviders({});

      expect(result.content[0].text).toContain('Online meeting providers API not available');
    });
  });

  describe('handleCreateOnlineMeeting', () => {
    test('should create online meeting', async () => {
      const mockMeeting = {
        id: 'meeting-1',
        subject: 'Team Sync',
        startTime: { dateTime: '2024-01-15T10:00:00Z' },
        endTime: { dateTime: '2024-01-15T11:00:00Z' },
        joinUrl: 'https://teams.microsoft.com/meeting/123',
        joinWebUrl: 'https://teams.microsoft.com/l/meetup-join/123',
        chatInfo: { threadId: 'chat-123' },
        videoTeleconferenceId: 'vc-123',
        audioConferencing: { tollNumber: '+1 555-123-4567', tollFreeNumber: null, conferenceId: '123456' },
      };

      mockClient.post.mockResolvedValue(mockMeeting);

      const result = await handleCreateOnlineMeeting({
        subject: 'Team Sync',
        startTime: '2024-01-15T10:00:00Z',
        endTime: '2024-01-15T11:00:00Z',
        attendees: ['user1@company.com'],
      });

      expect(result.content[0].text).toContain('Online Meeting Created');
      expect(result.content[0].text).toContain('Team Sync');
      expect(result.content[0].text).toContain('Join URL: https://teams.microsoft.com/meeting/123');
    });

    test('should return error for missing required fields', async () => {
      const result = await handleCreateOnlineMeeting({ subject: 'Test' });

      expect(result.content[0].text).toBe('startTime and endTime are required (ISO 8601 format).');
    });
  });

  describe('handleGetOnlineMeeting', () => {
    test('should get online meeting details', async () => {
      const mockMeeting = {
        id: 'meeting-1',
        subject: 'Team Sync',
        startTime: { dateTime: '2024-01-15T10:00:00Z' },
        endTime: { dateTime: '2024-01-15T11:00:00Z' },
        joinUrl: 'https://teams.microsoft.com/meeting/123',
        chatInfo: { threadId: 'chat-123' },
        videoTeleconferenceId: 'vc-123',
        audioConferencing: { tollNumber: '+1 555-123-4567' },
      };

      mockClient.get.mockResolvedValue(mockMeeting);

      const result = await handleGetOnlineMeeting({ meetingId: 'meeting-1' });

      expect(result.content[0].text).toContain('Online Meeting: Team Sync');
      expect(result.content[0].text).toContain('Join URL: https://teams.microsoft.com/meeting/123');
    });

    test('should return error when meetingId is missing', async () => {
      const result = await handleGetOnlineMeeting({});

      expect(result.content[0].text).toBe('Meeting ID is required.');
    });
  });
});