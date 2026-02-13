/**
 * Tests for the unified token-manager module.
 * Replaces the old token-storage.test.js.
 */
const fs = require('fs').promises;
const https = require('https');
const path = require('path');

// Mock fs.promises and https before requiring the module
jest.mock('fs', () => ({
  promises: {
    readFile: jest.fn(),
    writeFile: jest.fn(),
    unlink: jest.fn(),
  },
  existsSync: jest.fn(),
  readFileSync: jest.fn(),
}));
jest.mock('https');

// Mock config
jest.mock('../../config', () => ({
  AUTH_CONFIG: {
    clientId: 'test-client-id',
    clientSecret: 'test-client-secret',
    redirectUri: 'http://localhost:3333/auth/callback',
    scopes: ['offline_access', 'User.Read'],
    tokenStorePath: '/mock/home/.outlook-mcp-tokens.json',
  },
  TOKEN_ENDPOINT: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
  TOKEN_REFRESH_BUFFER_MS: 5 * 60 * 1000,
}));

// We need to get a fresh module for each test to reset cachedTokens
let tokenManager;

// Helper to reload the module fresh 
function reloadModule() {
  jest.resetModules();
  // Re-apply mocks after module reset
  jest.mock('fs', () => ({
    promises: {
      readFile: jest.fn(),
      writeFile: jest.fn(),
      unlink: jest.fn(),
    },
    existsSync: jest.fn(),
    readFileSync: jest.fn(),
  }));
  jest.mock('https');
  jest.mock('../../config', () => ({
    AUTH_CONFIG: {
      clientId: 'test-client-id',
      clientSecret: 'test-client-secret',
      redirectUri: 'http://localhost:3333/auth/callback',
      scopes: ['offline_access', 'User.Read'],
      tokenStorePath: '/mock/home/.outlook-mcp-tokens.json',
    },
    TOKEN_ENDPOINT: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
    TOKEN_REFRESH_BUFFER_MS: 5 * 60 * 1000,
  }));
  tokenManager = require('../../auth/token-manager');
}

describe('token-manager', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    reloadModule();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    jest.spyOn(console, 'warn').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
    console.warn.mockRestore();
  });

  describe('loadTokens', () => {
    test('should load and parse tokens from file', async () => {
      const mockTokens = { access_token: 'loaded_token', expires_at: Date.now() + 3600000 };
      const fsPromises = require('fs').promises;
      fsPromises.readFile.mockResolvedValue(JSON.stringify(mockTokens));

      const result = await tokenManager.loadTokens();

      expect(fsPromises.readFile).toHaveBeenCalledWith('/mock/home/.outlook-mcp-tokens.json', 'utf8');
      expect(result).toEqual(mockTokens);
    });

    test('should return null if file not found (ENOENT)', async () => {
      const fsPromises = require('fs').promises;
      fsPromises.readFile.mockRejectedValue({ code: 'ENOENT' });

      const result = await tokenManager.loadTokens();

      expect(result).toBeNull();
    });

    test('should return null for other read errors', async () => {
      const fsPromises = require('fs').promises;
      fsPromises.readFile.mockRejectedValue(new Error('Permission denied'));

      const result = await tokenManager.loadTokens();

      expect(result).toBeNull();
      const errorReason = tokenManager.getLastErrorReason();
      expect(errorReason).toBeTruthy();
      expect(errorReason.code).toBe('TOKEN_FILE_READ_ERROR');
    });

    test('should return null and set error reason for invalid JSON', async () => {
      const fsPromises = require('fs').promises;
      fsPromises.readFile.mockResolvedValue('invalid json {');

      const result = await tokenManager.loadTokens();

      expect(result).toBeNull();
      const errorReason = tokenManager.getLastErrorReason();
      expect(errorReason).toBeTruthy();
      expect(errorReason.code).toBe('TOKEN_FILE_INVALID_JSON');
      expect(errorReason.path).toBe('/mock/home/.outlook-mcp-tokens.json');
      expect(console.error).toHaveBeenCalledWith(
        expect.stringContaining('Invalid token JSON')
      );
    });

    test('should return null and set error reason for missing access_token', async () => {
      const fsPromises = require('fs').promises;
      fsPromises.readFile.mockResolvedValue(JSON.stringify({ refresh_token: 'rt' }));

      const result = await tokenManager.loadTokens();

      expect(result).toBeNull();
      const errorReason = tokenManager.getLastErrorReason();
      expect(errorReason).toBeTruthy();
      expect(errorReason.code).toBe('TOKEN_FILE_INVALID_SHAPE');
      expect(console.error).toHaveBeenCalledWith(
        expect.stringContaining('has no access_token')
      );
    });

    test('should return null and set error reason for non-object JSON', async () => {
      const fsPromises = require('fs').promises;
      fsPromises.readFile.mockResolvedValue('"just a string"');

      const result = await tokenManager.loadTokens();

      expect(result).toBeNull();
      const errorReason = tokenManager.getLastErrorReason();
      expect(errorReason).toBeTruthy();
      expect(errorReason.code).toBe('TOKEN_FILE_INVALID_SHAPE');
    });

    test('should clear error reason on successful load', async () => {
      const fsPromises = require('fs').promises;
      // First, trigger an error
      fsPromises.readFile.mockRejectedValueOnce({ code: 'ENOENT' });
      await tokenManager.loadTokens();
      expect(tokenManager.getLastErrorReason()).toBeTruthy();

      // Then, successful load
      const mockTokens = { access_token: 'loaded_token', expires_at: Date.now() + 3600000 };
      fsPromises.readFile.mockResolvedValue(JSON.stringify(mockTokens));
      await tokenManager.loadTokens();

      expect(tokenManager.getLastErrorReason()).toBeNull();
    });

    test('should return cached tokens on subsequent calls', async () => {
      const mockTokens = { access_token: 'cached_token', expires_at: Date.now() + 3600000 };
      const fsPromises = require('fs').promises;
      fsPromises.readFile.mockResolvedValue(JSON.stringify(mockTokens));

      const result1 = await tokenManager.loadTokens();
      const result2 = await tokenManager.loadTokens();

      expect(result1).toEqual(mockTokens);
      expect(result2).toEqual(mockTokens);
      // Only read once from disk; second call uses cache
      expect(fsPromises.readFile).toHaveBeenCalledTimes(1);
    });
  });

  describe('saveTokens', () => {
    test('should write tokens to file and update cache', async () => {
      const fsPromises = require('fs').promises;
      fsPromises.writeFile.mockResolvedValue();
      const tokens = { access_token: 'save_token', expires_at: Date.now() + 3600000 };

      await tokenManager.saveTokens(tokens);

      expect(fsPromises.writeFile).toHaveBeenCalledWith(
        '/mock/home/.outlook-mcp-tokens.json',
        JSON.stringify(tokens, null, 2)
      );
    });

    test('should warn if no tokens to save', async () => {
      const fsPromises = require('fs').promises;

      await tokenManager.saveTokens(null);

      expect(fsPromises.writeFile).not.toHaveBeenCalled();
      expect(console.warn).toHaveBeenCalledWith('[token-manager] No tokens to save');
    });
  });

  describe('isExpired', () => {
    test('should return true if tokens are null', () => {
      expect(tokenManager.isExpired(null)).toBe(true);
    });

    test('should return true if no expires_at', () => {
      expect(tokenManager.isExpired({ access_token: 'tok' })).toBe(true);
    });

    test('should return true if within 5-min buffer', () => {
      const tokens = { expires_at: Date.now() + (4 * 60 * 1000) }; // 4 min from now
      expect(tokenManager.isExpired(tokens)).toBe(true);
    });

    test('should return false if outside buffer', () => {
      const tokens = { expires_at: Date.now() + (10 * 60 * 1000) }; // 10 min from now
      expect(tokenManager.isExpired(tokens)).toBe(false);
    });
  });

  describe('getExpiryTime', () => {
    test('should return 0 when no tokens cached', () => {
      expect(tokenManager.getExpiryTime()).toBe(0);
    });

    test('should return expires_at after loading tokens', async () => {
      const expiryTime = Date.now() + 3600000;
      const mockTokens = { access_token: 'tok', expires_at: expiryTime };
      const fsPromises = require('fs').promises;
      fsPromises.readFile.mockResolvedValue(JSON.stringify(mockTokens));

      await tokenManager.loadTokens();

      expect(tokenManager.getExpiryTime()).toBe(expiryTime);
    });
  });

  describe('clearTokens', () => {
    test('should delete token file and clear cache', async () => {
      const fsPromises = require('fs').promises;
      fsPromises.unlink.mockResolvedValue();

      await tokenManager.clearTokens();

      expect(fsPromises.unlink).toHaveBeenCalledWith('/mock/home/.outlook-mcp-tokens.json');
    });

    test('should handle ENOENT gracefully when deleting', async () => {
      const fsPromises = require('fs').promises;
      fsPromises.unlink.mockRejectedValue({ code: 'ENOENT' });

      // Should not throw
      await tokenManager.clearTokens();

      expect(fsPromises.unlink).toHaveBeenCalled();
    });

    test('should handle other errors during delete', async () => {
      const fsPromises = require('fs').promises;
      fsPromises.unlink.mockRejectedValue(new Error('Disk error'));

      // Should not throw
      await tokenManager.clearTokens();

      expect(console.error).toHaveBeenCalled();
    });
  });

  describe('loadTokenCacheSync', () => {
    test('should return null when file does not exist', () => {
      const fsSync = require('fs');
      fsSync.existsSync.mockReturnValue(false);

      expect(tokenManager.loadTokenCacheSync()).toBeNull();
      const errorReason = tokenManager.getLastErrorReason();
      expect(errorReason).toBeTruthy();
      expect(errorReason.code).toBe('TOKEN_FILE_MISSING');
    });

    test('should read and return tokens from file', () => {
      const fsSync = require('fs');
      const mockTokens = { access_token: 'sync_token', expires_at: Date.now() + 3600000 };
      fsSync.existsSync.mockReturnValue(true);
      fsSync.readFileSync.mockReturnValue(JSON.stringify(mockTokens));

      const result = tokenManager.loadTokenCacheSync();

      expect(result).toEqual(mockTokens);
      // Successful load clears error
      expect(tokenManager.getLastErrorReason()).toBeNull();
    });

    test('should return null and set error reason for parse errors', () => {
      const fsSync = require('fs');
      fsSync.existsSync.mockReturnValue(true);
      fsSync.readFileSync.mockReturnValue('invalid json');

      expect(tokenManager.loadTokenCacheSync()).toBeNull();
      const errorReason = tokenManager.getLastErrorReason();
      expect(errorReason).toBeTruthy();
      expect(errorReason.code).toBe('TOKEN_FILE_INVALID_JSON');
      expect(console.error).toHaveBeenCalledWith(
        expect.stringContaining('Invalid token JSON')
      );
    });

    test('should return null and set error reason for missing access_token', () => {
      const fsSync = require('fs');
      fsSync.existsSync.mockReturnValue(true);
      fsSync.readFileSync.mockReturnValue(JSON.stringify({ refresh_token: 'rt' }));

      expect(tokenManager.loadTokenCacheSync()).toBeNull();
      const errorReason = tokenManager.getLastErrorReason();
      expect(errorReason).toBeTruthy();
      expect(errorReason.code).toBe('TOKEN_FILE_INVALID_SHAPE');
    });
  });

  describe('getAccessToken', () => {
    test('should return cached token if not expired', async () => {
      const mockTokens = { access_token: 'valid_token', expires_at: Date.now() + 3600000 };
      const fsPromises = require('fs').promises;
      fsPromises.readFile.mockResolvedValue(JSON.stringify(mockTokens));

      // Load tokens first
      await tokenManager.loadTokens();

      const token = await tokenManager.getAccessToken();

      expect(token).toBe('valid_token');
      // Valid token clears error
      expect(tokenManager.getLastErrorReason()).toBeNull();
    });

    test('should return null when no tokens and refresh fails', async () => {
      const fsPromises = require('fs').promises;
      fsPromises.readFile.mockRejectedValue({ code: 'ENOENT' });

      const token = await tokenManager.getAccessToken();

      expect(token).toBeNull();
      const errorReason = tokenManager.getLastErrorReason();
      expect(errorReason).toBeTruthy();
      expect(errorReason.code).toBe('TOKEN_FILE_MISSING');
    });

    test('should reload tokens from disk if cache is null', async () => {
      const fsPromises = require('fs').promises;
      // First call: file doesn't exist
      fsPromises.readFile.mockRejectedValueOnce({ code: 'ENOENT' });
      const token1 = await tokenManager.getAccessToken();
      expect(token1).toBeNull();

      // File appears later (simulated by successful read)
      const mockTokens = { access_token: 'new_token', expires_at: Date.now() + 3600000 };
      fsPromises.readFile.mockResolvedValue(JSON.stringify(mockTokens));

      // Second call should reload from disk
      const token2 = await tokenManager.getAccessToken();
      expect(token2).toBe('new_token');
    });
  });

  describe('getLastErrorReason', () => {
    test('should return null initially', () => {
      expect(tokenManager.getLastErrorReason()).toBeNull();
    });

    test('should return error reason after load failure', async () => {
      const fsPromises = require('fs').promises;
      fsPromises.readFile.mockRejectedValue({ code: 'ENOENT' });

      await tokenManager.loadTokens();

      const errorReason = tokenManager.getLastErrorReason();
      expect(errorReason).toBeTruthy();
      expect(errorReason.code).toBe('TOKEN_FILE_MISSING');
      expect(errorReason.path).toBe('/mock/home/.outlook-mcp-tokens.json');
    });
  });

  describe('exchangeCodeForTokens', () => {
    let mockRequest;

    beforeEach(() => {
      mockRequest = {
        on: jest.fn((event, cb) => {
          if (event === 'error') mockRequest.errorHandler = cb;
          return mockRequest;
        }),
        write: jest.fn(),
        end: jest.fn(),
      };
      const httpsModule = require('https');
      httpsModule.request.mockImplementation((options, callback) => {
        mockRequest.callback = callback;
        return mockRequest;
      });
    });

    test('should reject if client ID is not configured', async () => {
      // Override config to have no clientId
      jest.resetModules();
      jest.mock('fs', () => ({
        promises: { readFile: jest.fn(), writeFile: jest.fn(), unlink: jest.fn() },
        existsSync: jest.fn(),
        readFileSync: jest.fn(),
      }));
      jest.mock('https');
      jest.mock('../../config', () => ({
        AUTH_CONFIG: {
          clientId: '',
          clientSecret: '',
          redirectUri: 'http://localhost:3333/auth/callback',
          scopes: [],
          tokenStorePath: '/tmp/test-tokens.json',
        },
        TOKEN_ENDPOINT: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
        TOKEN_REFRESH_BUFFER_MS: 5 * 60 * 1000,
      }));
      const tm = require('../../auth/token-manager');

      await expect(tm.exchangeCodeForTokens('code123'))
        .rejects.toThrow('Client ID or Client Secret is not configured');
      
      const errorReason = tm.getLastErrorReason();
      expect(errorReason).toBeTruthy();
      expect(errorReason.code).toBe('CLIENT_CONFIG_MISSING');
    });

    test('should exchange code and save tokens on success', async () => {
      const fsPromises = require('fs').promises;
      fsPromises.writeFile.mockResolvedValue();

      const tokenResponse = {
        access_token: 'new_access_token',
        refresh_token: 'new_refresh_token',
        expires_in: 3600,
        scope: 'User.Read',
        token_type: 'Bearer'
      };

      const exchangePromise = tokenManager.exchangeCodeForTokens('auth_code_123');

      // Simulate successful HTTPS response
      const mockRes = {
        statusCode: 200,
        on: (event, cb) => {
          if (event === 'data') cb(Buffer.from(JSON.stringify(tokenResponse)));
          if (event === 'end') cb();
        }
      };
      mockRequest.callback(mockRes);

      const tokens = await exchangePromise;

      expect(tokens.access_token).toBe('new_access_token');
      expect(tokens.refresh_token).toBe('new_refresh_token');
      expect(tokens.expires_at).toBeGreaterThan(Date.now());
      expect(fsPromises.writeFile).toHaveBeenCalled();
    });

    test('should reject on API error and set error reason', async () => {
      const errorResponse = { error: 'invalid_grant', error_description: 'Bad auth code' };

      const exchangePromise = tokenManager.exchangeCodeForTokens('bad_code');

      const mockRes = {
        statusCode: 400,
        on: (event, cb) => {
          if (event === 'data') cb(Buffer.from(JSON.stringify(errorResponse)));
          if (event === 'end') cb();
        }
      };
      mockRequest.callback(mockRes);

      await expect(exchangePromise).rejects.toThrow('Bad auth code');
      
      const errorReason = tokenManager.getLastErrorReason();
      expect(errorReason).toBeTruthy();
      expect(errorReason.code).toBe('CODE_EXCHANGE_FAILED');
      expect(errorReason.statusCode).toBe(400);
    });

    test('should reject on network error and set error reason', async () => {
      const exchangePromise = tokenManager.exchangeCodeForTokens('code');

      mockRequest.errorHandler(new Error('Network fail'));

      await expect(exchangePromise).rejects.toThrow('Network fail');
      
      const errorReason = tokenManager.getLastErrorReason();
      expect(errorReason).toBeTruthy();
      expect(errorReason.code).toBe('CODE_EXCHANGE_NETWORK_ERROR');
    });
  });
});
