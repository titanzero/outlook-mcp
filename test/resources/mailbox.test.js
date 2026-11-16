/**
 * Tests for mailbox resources
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { registerMailboxResources } = require('../../mailbox/resources');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');

describe('registerMailboxResources', () => {
  let mockServer;
  let mockClient;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});

    mockClient = {
      api: jest.fn(() => ({
        get: jest.fn().mockResolvedValue({
          timeZone: 'UTC',
          language: { locale: 'en-US', displayName: 'English (United States)' },
          automaticRepliesSetting: { status: 'disabled' },
        }),
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);

    mockServer = { resource: jest.fn() };
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('should register mailbox resources', () => {
    registerMailboxResources(mockServer);
    expect(mockServer.resource).toHaveBeenCalled();
  });
});
