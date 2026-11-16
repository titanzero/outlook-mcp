/**
 * Tests for the auth index module (ensureAuthenticated, buildAuthError)
 */
const { ensureAuthenticated, tokenManager, registerAuthTools } = require('../../auth/index');

jest.mock('../../auth/token-manager', () => ({
  getAccessToken: jest.fn(),
  getLastErrorReason: jest.fn(),
  loadTokenCacheSync: jest.fn(),
  clearTokens: jest.fn(),
  exchangeCodeForTokens: jest.fn(),
  getExpiryTime: jest.fn(),
}));

const tokenManagerMock = require('../../auth/token-manager');

describe('auth/index', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('ensureAuthenticated', () => {
    test('should throw buildAuthError when forceNew is true', async () => {
      tokenManagerMock.getLastErrorReason.mockReturnValue({ code: 'TOKEN_FILE_MISSING' });

      await expect(ensureAuthenticated(true)).rejects.toThrow('Authentication required');
    });

    test('should return access token when valid', async () => {
      tokenManagerMock.getAccessToken.mockResolvedValue('valid_token_123');

      const token = await ensureAuthenticated();

      expect(tokenManagerMock.getAccessToken).toHaveBeenCalledTimes(1);
      expect(token).toBe('valid_token_123');
    });

    test('should throw buildAuthError when getAccessToken returns null', async () => {
      tokenManagerMock.getAccessToken.mockResolvedValue(null);
      tokenManagerMock.getLastErrorReason.mockReturnValue({ code: 'TOKEN_FILE_MISSING', path: '/tmp/tokens.json' });

      await expect(ensureAuthenticated()).rejects.toThrow('token file not found');
    });

    test('should propagate auth error details when token is null', async () => {
      tokenManagerMock.getAccessToken.mockResolvedValue(null);
      tokenManagerMock.getLastErrorReason.mockReturnValue({
        code: 'TOKEN_FILE_INVALID_JSON',
        path: '/tmp/tokens.json',
      });

      try {
        await ensureAuthenticated();
      } catch (error) {
        expect(error.isAuthError).toBe(true);
        expect(error.authReason).toBe('TOKEN_FILE_INVALID_JSON');
        expect(error.authDetails).toBeDefined();
        expect(error.message).toContain('invalid JSON');
      }
    });
  });

  describe('buildAuthError (via ensureAuthenticated)', () => {
    const errorCases = [
      {
        code: 'TOKEN_FILE_MISSING',
        path: '/home/user/.outlook-mcp-tokens.json',
        expected: 'token file not found',
      },
      {
        code: 'TOKEN_FILE_INVALID_JSON',
        path: '/home/user/.outlook-mcp-tokens.json',
        expected: 'invalid JSON',
      },
      {
        code: 'TOKEN_FILE_INVALID_SHAPE',
        path: '/home/user/.outlook-mcp-tokens.json',
        expected: 'missing required fields',
      },
      {
        code: 'CLIENT_CONFIG_MISSING',
        expected: 'OUTLOOK_CLIENT_ID and/or OUTLOOK_CLIENT_SECRET are not set',
      },
      {
        code: 'REFRESH_FAILED_INVALID_CLIENT',
        expected: 'Microsoft rejected the client credentials',
      },
      {
        code: 'CODE_EXCHANGE_INVALID_CLIENT',
        expected: 'Microsoft rejected the client credentials',
      },
      {
        code: 'REFRESH_FAILED_NO_REFRESH_TOKEN',
        expected: 'no refresh token found',
      },
      {
        code: 'REFRESH_NETWORK_ERROR',
        message: 'ECONNREFUSED',
        expected: 'network error',
      },
      {
        code: 'CODE_EXCHANGE_NETWORK_ERROR',
        message: 'ENOTFOUND',
        expected: 'network error',
      },
      {
        code: 'UNKNOWN_CODE',
        message: 'custom error message',
        expected: 'custom error message',
      },
    ];

    test.each(errorCases)('should format error for %s', async ({ code, path, message, expected }) => {
      tokenManagerMock.getAccessToken.mockResolvedValue(null);
      tokenManagerMock.getLastErrorReason.mockReturnValue({ code, path, message });

      try {
        await ensureAuthenticated();
      } catch (error) {
        expect(error.isAuthError).toBe(true);
        expect(error.authReason).toBe(code);
        expect(error.message).toContain(expected);
      }
    });

    test('should set isAuthError=true and include authDetails', async () => {
      tokenManagerMock.getAccessToken.mockResolvedValue(null);
      tokenManagerMock.getLastErrorReason.mockReturnValue({
        code: 'TOKEN_FILE_MISSING',
        path: '/tmp/tokens.json',
      });

      try {
        await ensureAuthenticated();
      } catch (error) {
        expect(error.isAuthError).toBe(true);
        expect(error.authDetails).toEqual({
          code: 'TOKEN_FILE_MISSING',
          path: '/tmp/tokens.json',
        });
      }
    });
  });

  describe('registerAuthTools', () => {
    test('should register three tools on the server', () => {
      const mockServer = {
        tool: jest.fn(),
      };

      registerAuthTools(mockServer);

      expect(mockServer.tool).toHaveBeenCalledTimes(3);

      const toolCalls = mockServer.tool.mock.calls.map(c => c[0]);
      expect(toolCalls).toContain('about');
      expect(toolCalls).toContain('authenticate');
      expect(toolCalls).toContain('check-auth-status');
    });

    test('should register about tool with readOnlyHint', () => {
      const mockServer = { tool: jest.fn() };
      registerAuthTools(mockServer);

      const aboutCall = mockServer.tool.mock.calls.find(c => c[0] === 'about');
      expect(aboutCall[3].annotations.readOnlyHint).toBe(true);
      expect(aboutCall[3].annotations.openWorldHint).toBe(false);
    });

    test('should register authenticate tool with annotations', () => {
      const mockServer = { tool: jest.fn() };
      registerAuthTools(mockServer);

      const authCall = mockServer.tool.mock.calls.find(c => c[0] === 'authenticate');
      expect(authCall[3].annotations.readOnlyHint).toBe(false);
      expect(authCall[3].annotations.destructiveHint).toBe(false);
      expect(authCall[3].annotations.idempotentHint).toBe(false);
      expect(authCall[3].annotations.openWorldHint).toBe(true);
    });

    test('should register check-auth-status tool with readOnlyHint', () => {
      const mockServer = { tool: jest.fn() };
      registerAuthTools(mockServer);

      const statusCall = mockServer.tool.mock.calls.find(c => c[0] === 'check-auth-status');
      expect(statusCall[3].annotations.readOnlyHint).toBe(true);
      expect(statusCall[3].annotations.openWorldHint).toBe(false);
    });
  });

  describe('exports', () => {
    test('should export tokenManager', () => {
      expect(tokenManager).toBeDefined();
      expect(typeof tokenManager.getAccessToken).toBe('function');
    });

    test('should export registerAuthTools', () => {
      expect(typeof registerAuthTools).toBe('function');
    });

    test('should export ensureAuthenticated', () => {
      expect(typeof ensureAuthenticated).toBe('function');
    });
  });
});
