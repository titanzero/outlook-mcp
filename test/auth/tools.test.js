/**
 * Tests for the auth tools module (handleAbout, handleAuthenticate, handleCheckAuthStatus)
 */
const { handleAbout, handleAuthenticate, handleCheckAuthStatus } = require('../../auth/tools');
const tokenManager = require('../../auth/token-manager');
const { makeResponse, makeErrorResponse } = require('../../utils/response-helpers');

jest.mock('../../config', () => ({
  SERVER_VERSION: '1.0.0-test',
  AUTH_CONFIG: {
    clientId: 'test-client-id',
    clientSecret: 'test-client-secret',
    authServerUrl: 'http://localhost:3333',
  },
}));

jest.mock('../../auth/token-manager', () => ({
  clearTokens: jest.fn(),
  loadTokenCacheSync: jest.fn(),
}));

jest.mock('../../utils/response-helpers', () => ({
  makeResponse: jest.fn(text => ({ content: [{ text }], isError: false })),
  makeErrorResponse: jest.fn(text => ({ content: [{ text }], isError: true })),
}));

describe('auth/tools', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  describe('handleAbout', () => {
    test('should return server info response', async () => {
      const result = await handleAbout();

      expect(makeResponse).toHaveBeenCalledWith(
        expect.stringContaining('MODULAR Outlook Assistant MCP Server v1.0.0-test')
      );
      expect(result.isError).toBe(false);
    });

    test('should include Microsoft Graph API mention', async () => {
      await handleAbout();
      expect(makeResponse).toHaveBeenCalledWith(
        expect.stringContaining('Microsoft Graph API')
      );
    });
  });

  describe('handleAuthenticate', () => {
    test('should return auth URL when not forced', async () => {
      const result = await handleAuthenticate({ force: false });

      expect(makeResponse).toHaveBeenCalledWith(
        expect.stringContaining('Authentication required')
      );
      expect(makeResponse).toHaveBeenCalledWith(
        expect.stringContaining('http://localhost:3333/auth?client_id=test-client-id')
      );
      expect(tokenManager.clearTokens).not.toHaveBeenCalled();
    });

    test('should clear tokens and return auth URL when forced', async () => {
      tokenManager.clearTokens.mockResolvedValue();

      const result = await handleAuthenticate({ force: true });

      expect(tokenManager.clearTokens).toHaveBeenCalledTimes(1);
      expect(makeResponse).toHaveBeenCalledWith(
        expect.stringContaining('Existing tokens were cleared')
      );
      expect(makeResponse).toHaveBeenCalledWith(
        expect.stringContaining('http://localhost:3333/auth?client_id=test-client-id')
      );
    });

    test('should return error when clientId not configured', async () => {
      // Use a fresh module with different config
      jest.resetModules();
      jest.mock('../../config', () => ({
        SERVER_VERSION: '1.0.0-test',
        AUTH_CONFIG: {
          clientId: '',
          clientSecret: 'test-secret',
          authServerUrl: 'http://localhost:3333',
        },
      }));
      jest.mock('../../auth/token-manager', () => ({
        clearTokens: jest.fn(),
        loadTokenCacheSync: jest.fn(),
      }));
      jest.mock('../../utils/response-helpers', () => ({
        makeResponse: jest.fn(text => ({ content: [{ text }], isError: false })),
        makeErrorResponse: jest.fn(text => ({ content: [{ text }], isError: true })),
      }));

      const { handleAuthenticate: freshHandleAuthenticate } = require('../../auth/tools');
      const { makeErrorResponse: freshMakeErrorResponse } = require('../../utils/response-helpers');

      const result = await freshHandleAuthenticate({});

      expect(freshMakeErrorResponse).toHaveBeenCalledWith(
        expect.stringContaining('Authentication configuration is missing')
      );
      expect(result.isError).toBe(true);
    });

    test('should handle errors during token clear', async () => {
      tokenManager.clearTokens.mockRejectedValue(new Error('Clear failed'));

      const result = await handleAuthenticate({ force: true });

      expect(makeErrorResponse).toHaveBeenCalledWith(
        expect.stringContaining('Error starting authentication flow: Clear failed')
      );
      expect(result.isError).toBe(true);
    });

    test('should use default empty object when args is undefined', async () => {
      const result = await handleAuthenticate(undefined);

      expect(makeResponse).toHaveBeenCalledWith(
        expect.stringContaining('Authentication required')
      );
    });
  });

  describe('handleCheckAuthStatus', () => {
    test('should return "Not authenticated" when no tokens', async () => {
      tokenManager.loadTokenCacheSync.mockReturnValue(null);

      const result = await handleCheckAuthStatus();

      expect(tokenManager.loadTokenCacheSync).toHaveBeenCalledTimes(1);
      expect(makeResponse).toHaveBeenCalledWith('Not authenticated');
    });

    test('should return "Not authenticated" when no access_token', async () => {
      tokenManager.loadTokenCacheSync.mockReturnValue({ refresh_token: 'rt' });

      const result = await handleCheckAuthStatus();

      expect(makeResponse).toHaveBeenCalledWith('Not authenticated');
    });

    test('should return "Authenticated and ready" when valid tokens', async () => {
      const mockTokens = {
        access_token: 'valid_token',
        expires_at: Date.now() + 3600000,
      };
      tokenManager.loadTokenCacheSync.mockReturnValue(mockTokens);

      const result = await handleCheckAuthStatus();

      expect(makeResponse).toHaveBeenCalledWith('Authenticated and ready');
    });

    test('should log debug info to console.error', async () => {
      tokenManager.loadTokenCacheSync.mockReturnValue({
        access_token: 'valid_token',
        expires_at: Date.now() + 3600000,
      });

      await handleCheckAuthStatus();

      expect(console.error).toHaveBeenCalledWith('[CHECK-AUTH-STATUS] Starting authentication status check');
      expect(console.error).toHaveBeenCalledWith('[CHECK-AUTH-STATUS] Tokens loaded: YES');
      expect(console.error).toHaveBeenCalledWith('[CHECK-AUTH-STATUS] Access token present');
      expect(console.error).toHaveBeenCalledWith(
        expect.stringContaining('[CHECK-AUTH-STATUS] Token expires at:')
      );
      expect(console.error).toHaveBeenCalledWith(
        expect.stringContaining('[CHECK-AUTH-STATUS] Current time:')
      );
    });

    test('should log "NO" when no tokens', async () => {
      tokenManager.loadTokenCacheSync.mockReturnValue(null);

      await handleCheckAuthStatus();

      expect(console.error).toHaveBeenCalledWith('[CHECK-AUTH-STATUS] Tokens loaded: NO');
      expect(console.error).toHaveBeenCalledWith('[CHECK-AUTH-STATUS] No valid access token found');
    });
  });
});
