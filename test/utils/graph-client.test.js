/**
 * Tests for the Graph client wrapper
 */
const { getGraphClient, graphGetPaginated, resetClient, OutlookAuthProvider } = require('../../utils/graph-client');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../auth', () => ({
  ensureAuthenticated: jest.fn(),
}));
jest.mock('@microsoft/microsoft-graph-client', () => ({
  Client: {
    initWithMiddleware: jest.fn(),
  },
}));

const { Client } = require('@microsoft/microsoft-graph-client');

describe('graph-client', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    resetClient();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('getGraphClient', () => {
    test('should call ensureAuthenticated and return client instance', async () => {
      ensureAuthenticated.mockResolvedValue('test_access_token');
      const mockClient = { api: jest.fn() };
      Client.initWithMiddleware.mockReturnValue(mockClient);

      const client = await getGraphClient();

      expect(ensureAuthenticated).toHaveBeenCalledTimes(1);
      expect(Client.initWithMiddleware).toHaveBeenCalledWith({
        authProvider: expect.any(OutlookAuthProvider),
      });
      expect(client).toBe(mockClient);
    });

    test('should reuse cached client if token unchanged', async () => {
      ensureAuthenticated.mockResolvedValue('same_token');
      const mockClient = { api: jest.fn() };
      Client.initWithMiddleware.mockReturnValue(mockClient);

      const client1 = await getGraphClient();
      const client2 = await getGraphClient();

      expect(Client.initWithMiddleware).toHaveBeenCalledTimes(1);
      expect(client1).toBe(client2);
    });

    test('should recreate client if token changed', async () => {
      ensureAuthenticated.mockResolvedValueOnce('token_1');
      const mockClient1 = { api: jest.fn() };
      Client.initWithMiddleware.mockReturnValueOnce(mockClient1);

      const client1 = await getGraphClient();

      ensureAuthenticated.mockResolvedValueOnce('token_2');
      const mockClient2 = { api: jest.fn() };
      Client.initWithMiddleware.mockReturnValueOnce(mockClient2);

      const client2 = await getGraphClient();

      expect(Client.initWithMiddleware).toHaveBeenCalledTimes(2);
      expect(client1).not.toBe(client2);
    });

    test('should propagate auth errors', async () => {
      const authError = new Error('Authentication required');
      authError.isAuthError = true;
      ensureAuthenticated.mockRejectedValue(authError);

      await expect(getGraphClient()).rejects.toThrow('Authentication required');
    });
  });

  describe('OutlookAuthProvider', () => {
    test('getAccessToken should call ensureAuthenticated', async () => {
      ensureAuthenticated.mockResolvedValue('provider_token');

      const provider = new OutlookAuthProvider();
      const token = await provider.getAccessToken();

      expect(ensureAuthenticated).toHaveBeenCalledTimes(1);
      expect(token).toBe('provider_token');
    });

    test('should propagate auth errors', async () => {
      const authError = new Error('Auth failed');
      ensureAuthenticated.mockRejectedValue(authError);

      const provider = new OutlookAuthProvider();
      await expect(provider.getAccessToken()).rejects.toThrow('Auth failed');
    });
  });

  describe('graphGetPaginated', () => {
    let mockClient;

    beforeEach(() => {
      mockClient = {
        api: jest.fn(),
      };
    });

    test('should fetch single page without pagination', async () => {
      const mockResponse = {
        value: [{ id: '1' }, { id: '2' }],
        '@odata.count': 2,
      };

      const chain = {
        top: jest.fn().mockReturnThis(),
        select: jest.fn().mockReturnThis(),
        filter: jest.fn().mockReturnThis(),
        orderby: jest.fn().mockReturnThis(),
        get: jest.fn().mockResolvedValue(mockResponse),
      };
      mockClient.api.mockReturnValue(chain);

      const result = await graphGetPaginated(mockClient, 'me/messages', { $top: 10 }, 10);

      expect(mockClient.api).toHaveBeenCalledWith('me/messages');
      expect(chain.top).toHaveBeenCalledWith(10);
      expect(chain.get).toHaveBeenCalledTimes(1);
      expect(result.value).toHaveLength(2);
      expect(result['@odata.count']).toBe(2);
    });

    test('should follow @odata.nextLink for pagination', async () => {
      const page1 = { value: [{ id: '1' }, { id: '2' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/messages?$skiptoken=abc' };
      const page2 = { value: [{ id: '3' }], '@odata.nextLink': null };

      const chain1 = {
        top: jest.fn().mockReturnThis(),
        select: jest.fn().mockReturnThis(),
        get: jest.fn().mockResolvedValue(page1),
      };
      const chain2 = {
        get: jest.fn().mockResolvedValue(page2),
      };

      mockClient.api
        .mockReturnValueOnce(chain1)
        .mockReturnValueOnce(chain2);

      const result = await graphGetPaginated(mockClient, 'me/messages', { $top: 2 }, 10);

      expect(mockClient.api).toHaveBeenCalledTimes(2);
      expect(chain1.get).toHaveBeenCalledTimes(1);
      expect(chain2.get).toHaveBeenCalledTimes(1);
      expect(result.value).toHaveLength(3);
      expect(result['@odata.count']).toBe(3);
    });

    test('should respect maxCount limit', async () => {
      const page1 = { value: [{ id: '1' }, { id: '2' }, { id: '3' }], '@odata.nextLink': 'https://next' };
      const page2 = { value: [{ id: '4' }, { id: '5' }], '@odata.nextLink': null };

      const chain1 = { top: jest.fn().mockReturnThis(), get: jest.fn().mockResolvedValue(page1) };
      const chain2 = { get: jest.fn().mockResolvedValue(page2) };

      mockClient.api
        .mockReturnValueOnce(chain1)
        .mockReturnValueOnce(chain2);

      const result = await graphGetPaginated(mockClient, 'me/messages', { $top: 10 }, 3);

      expect(result.value).toHaveLength(3);
      expect(result.value.map(i => i.id)).toEqual(['1', '2', '3']);
    });

    test('should apply all query parameter types', async () => {
      const mockResponse = { value: [], '@odata.nextLink': null };
      const chain = {
        top: jest.fn().mockReturnThis(),
        select: jest.fn().mockReturnThis(),
        filter: jest.fn().mockReturnThis(),
        orderby: jest.fn().mockReturnThis(),
        search: jest.fn().mockReturnThis(),
        count: jest.fn().mockReturnThis(),
        expand: jest.fn().mockReturnThis(),
        skip: jest.fn().mockReturnThis(),
        query: jest.fn().mockReturnThis(),
        get: jest.fn().mockResolvedValue(mockResponse),
      };
      mockClient.api.mockReturnValue(chain);

      await graphGetPaginated(mockClient, 'me/messages', {
        $top: 10,
        $select: 'id,subject',
        $filter: 'isRead eq false',
        $orderby: 'receivedDateTime desc',
        $search: 'test',
        $count: true,
        $expand: 'attachments',
        $skip: 5,
        startDateTime: '2024-01-01T00:00:00Z',
      });

      expect(chain.top).toHaveBeenCalledWith(10);
      expect(chain.select).toHaveBeenCalledWith('id,subject');
      expect(chain.filter).toHaveBeenCalledWith('isRead eq false');
      expect(chain.orderby).toHaveBeenCalledWith('receivedDateTime desc');
      expect(chain.search).toHaveBeenCalledWith('test');
      expect(chain.count).toHaveBeenCalledWith(true);
      expect(chain.expand).toHaveBeenCalledWith('attachments');
      expect(chain.skip).toHaveBeenCalledWith(5);
      expect(chain.query).toHaveBeenCalledWith({ startDateTime: '2024-01-01T00:00:00Z' });
    });

    test('should handle API errors', async () => {
      const chain = { top: jest.fn().mockReturnThis(), get: jest.fn().mockRejectedValue(new Error('Graph API Error')) };
      mockClient.api.mockReturnValue(chain);

      await expect(graphGetPaginated(mockClient, 'me/messages', {}, 10)).rejects.toThrow('Graph API Error');
    });

    test('should return empty array when no value in response', async () => {
      const mockResponse = { '@odata.nextLink': null };
      const chain = { top: jest.fn().mockReturnThis(), get: jest.fn().mockResolvedValue(mockResponse) };
      mockClient.api.mockReturnValue(chain);

      const result = await graphGetPaginated(mockClient, 'me/messages', {}, 10);

      expect(result.value).toEqual([]);
    });

    test('should handle non-array value gracefully', async () => {
      const mockResponse = { value: 'not an array', '@odata.nextLink': null };
      const chain = { top: jest.fn().mockReturnThis(), get: jest.fn().mockResolvedValue(mockResponse) };
      mockClient.api.mockReturnValue(chain);

      const result = await graphGetPaginated(mockClient, 'me/messages', {}, 10);

      expect(result.value).toEqual([]);
    });
  });

  describe('resetClient', () => {
    test('should reset cached client and token', async () => {
      ensureAuthenticated.mockResolvedValue('token_1');
      const mockClient = { api: jest.fn() };
      Client.initWithMiddleware.mockReturnValue(mockClient);

      await getGraphClient();
      resetClient();

      ensureAuthenticated.mockResolvedValue('token_2');
      const mockClient2 = { api: jest.fn() };
      Client.initWithMiddleware.mockReturnValueOnce(mockClient2);

      const client = await getGraphClient();

      expect(Client.initWithMiddleware).toHaveBeenCalledTimes(2);
      expect(client).toBe(mockClient2);
    });
  });
});
