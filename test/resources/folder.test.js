/**
 * Tests for folder resources
 */
process.env.OUTLOOK_RESPONSE_FORMAT = 'text';

const { registerFolderResources } = require('../../folder/resources');
const { getGraphClient } = require('../../utils/graph-client');

jest.mock('../../utils/graph-client');
jest.mock('@modelcontextprotocol/sdk/server/mcp.js', () => ({
  ResourceTemplate: jest.fn().mockImplementation((uri, options) => ({ uri, options })),
}));

describe('registerFolderResources', () => {
  let mockServer;
  let mockClient;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});

    mockClient = {
      api: jest.fn(() => ({
        select: jest.fn().mockReturnThis(),
        get: jest.fn().mockResolvedValue({ value: [] }),
      })),
    };
    getGraphClient.mockResolvedValue(mockClient);

    mockServer = { resource: jest.fn() };
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('should register folder resources', () => {
    registerFolderResources(mockServer);
    expect(mockServer.resource).toHaveBeenCalled();
  });
});
