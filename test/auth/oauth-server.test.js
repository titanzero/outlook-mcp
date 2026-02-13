const express = require('express');
const request = require('supertest');
const { setupOAuthRoutes, createAuthConfig } = require('../../auth/oauth-server');
const appConfig = require('../../config');

// Mock the token-manager module used internally by oauth-server
jest.mock('../../auth/token-manager', () => ({
  exchangeCodeForTokens: jest.fn(),
  getAccessToken: jest.fn(),
  getExpiryTime: jest.fn(),
}));

const tokenManager = require('../../auth/token-manager');

const mockAuthConfig = {
  clientId: 'test-client-id',
  clientSecret: 'test-client-secret',
  redirectUri: 'http://localhost:3334/auth/callback',
  scopes: ['test_scope', 'openid'],
  tokenEndpoint: 'https://login.example.com/token',
  authEndpoint: 'https://login.example.com/authorize',
};

describe('OAuth Server Routes', () => {
  let app;

  beforeEach(() => {
    jest.clearAllMocks();
    app = express();
    setupOAuthRoutes(app, mockAuthConfig);
  });

  describe('GET /auth', () => {
    it('should redirect to the OAuth provider with correct parameters', async () => {
      const response = await request(app).get('/auth');
      expect(response.status).toBe(302);

      const redirectUrl = new URL(response.headers.location);
      expect(redirectUrl.origin).toBe(mockAuthConfig.authEndpoint.split('/authorize')[0]);
      expect(redirectUrl.pathname).toBe('/authorize');
      expect(redirectUrl.searchParams.get('client_id')).toBe(mockAuthConfig.clientId);
      expect(redirectUrl.searchParams.get('response_type')).toBe('code');
      expect(redirectUrl.searchParams.get('redirect_uri')).toBe(mockAuthConfig.redirectUri);
      expect(redirectUrl.searchParams.get('scope')).toBe(mockAuthConfig.scopes.join(' '));
      expect(redirectUrl.searchParams.get('response_mode')).toBe('query');
      expect(redirectUrl.searchParams.get('state')).toBeDefined();
      expect(redirectUrl.searchParams.get('state').length).toBe(32);
    });

    it('should return 500 if clientId is not configured', async () => {
      const tempApp = express();
      const noClientIdAuthConfig = { ...mockAuthConfig, clientId: null };
      setupOAuthRoutes(tempApp, noClientIdAuthConfig);

      const response = await request(tempApp).get('/auth');
      expect(response.status).toBe(500);
      expect(response.text).toContain('Authorization Failed');
      expect(response.text).toContain('Configuration Error');
      expect(response.text).toContain('Client ID is not configured.');
    });
  });

  describe('GET /auth/callback', () => {
    const mockAuthCode = 'mock_auth_code';
    const mockState = 'mock_state_value';

    it('should exchange code for tokens and return success HTML', async () => {
      tokenManager.exchangeCodeForTokens.mockResolvedValue({ access_token: 'mock_access_token' });

      const response = await request(app).get(`/auth/callback?code=${mockAuthCode}&state=${mockState}`);

      expect(tokenManager.exchangeCodeForTokens).toHaveBeenCalledWith(mockAuthCode);
      expect(response.status).toBe(200);
      expect(response.text).toContain('Authentication Successful');
    });

    it('should return 400 and error HTML if OAuth provider returns an error', async () => {
      const oauthError = 'access_denied';
      const oauthErrorDesc = 'User denied access';
      const response = await request(app).get(`/auth/callback?error=${oauthError}&error_description=${oauthErrorDesc}&state=${mockState}`);

      expect(response.status).toBe(400);
      expect(response.text).toContain('Authorization Failed');
      expect(response.text).toContain(oauthError);
      expect(response.text).toContain(oauthErrorDesc);
      expect(tokenManager.exchangeCodeForTokens).not.toHaveBeenCalled();
    });

    it('should return 400 if no code is provided', async () => {
      const response = await request(app).get(`/auth/callback?state=${mockState}`);
      expect(response.status).toBe(400);
      expect(response.text).toContain('Authorization Failed');
      expect(response.text).toContain('Missing Authorization Code');
      expect(tokenManager.exchangeCodeForTokens).not.toHaveBeenCalled();
    });

    it('should return 400 if state is missing from callback', async () => {
      const consoleErrorSpy = jest.spyOn(console, 'error').mockImplementation();
      const response = await request(app).get(`/auth/callback?code=${mockAuthCode}`);

      expect(response.status).toBe(400);
      expect(response.text).toContain('Authorization Failed');
      expect(response.text).toContain('Missing State Parameter');
      expect(tokenManager.exchangeCodeForTokens).not.toHaveBeenCalled();
      consoleErrorSpy.mockRestore();
    });

    it('should return 500 if token exchange fails', async () => {
      const exchangeError = new Error('Token exchange process failed');
      tokenManager.exchangeCodeForTokens.mockRejectedValue(exchangeError);

      const response = await request(app).get(`/auth/callback?code=${mockAuthCode}&state=${mockState}`);

      expect(response.status).toBe(500);
      expect(response.text).toContain('Token Exchange Failed');
      expect(response.text).toContain(exchangeError.message);
    });
  });

  describe('GET /token-status', () => {
    it('should return valid status if token exists and is valid', async () => {
      const mockExpiry = Date.now() + 3600000;
      tokenManager.getAccessToken.mockResolvedValue('valid_token_123');
      tokenManager.getExpiryTime.mockReturnValue(mockExpiry);

      const response = await request(app).get('/token-status');

      expect(response.status).toBe(200);
      expect(response.text).toContain('Token Status');
      expect(response.text).toContain('Access token is valid.');
    });

    it('should return "no valid token" status if token is not found', async () => {
      tokenManager.getAccessToken.mockResolvedValue(null);

      const response = await request(app).get('/token-status');

      expect(response.status).toBe(200);
      expect(response.text).toContain('Token Status');
      expect(response.text).toContain('No valid access token found. Please authenticate.');
    });

    it('should return 500 if checking token status throws an error', async () => {
      const statusError = new Error('Failed to check token status');
      tokenManager.getAccessToken.mockRejectedValue(statusError);

      const response = await request(app).get('/token-status');

      expect(response.status).toBe(500);
      expect(response.text).toContain('Token Status');
      expect(response.text).toContain(`Error checking token status: ${statusError.message}`);
    });
  });

  describe('createAuthConfig', () => {
    const originalEnv = process.env;

    beforeEach(() => {
      jest.resetModules();
      process.env = { ...originalEnv };
    });

    afterAll(() => {
      process.env = originalEnv;
    });

    it('should use shared config defaults when OUTLOOK variables are not set', () => {
      delete process.env.OUTLOOK_CLIENT_ID;
      delete process.env.OUTLOOK_CLIENT_SECRET;
      delete process.env.OUTLOOK_REDIRECT_URI;
      delete process.env.OUTLOOK_SCOPES;

      const config = createAuthConfig();
      expect(config.clientId).toBe(appConfig.AUTH_CONFIG.clientId);
      expect(config.clientSecret).toBe(appConfig.AUTH_CONFIG.clientSecret);
      expect(config.redirectUri).toBe(appConfig.AUTH_CONFIG.redirectUri);
      expect(config.scopes).toEqual(appConfig.AUTH_CONFIG.scopes);
      expect(config.tokenEndpoint).toBe(appConfig.TOKEN_ENDPOINT);
      expect(config.authEndpoint).toBe(appConfig.AUTH_ENDPOINT);
    });

    it('should use OUTLOOK_* environment variables when provided', () => {
      process.env.OUTLOOK_CLIENT_ID = 'env_client_id';
      process.env.OUTLOOK_CLIENT_SECRET = 'env_client_secret';
      process.env.OUTLOOK_REDIRECT_URI = 'http://env.redirect/uri';
      process.env.OUTLOOK_SCOPES = 'scope1 scope2';

      const config = createAuthConfig();

      expect(config.clientId).toBe('env_client_id');
      expect(config.clientSecret).toBe('env_client_secret');
      expect(config.redirectUri).toBe('http://env.redirect/uri');
      expect(config.scopes).toEqual(['scope1', 'scope2']);
      expect(config.tokenEndpoint).toBe(appConfig.TOKEN_ENDPOINT);
      expect(config.authEndpoint).toBe(appConfig.AUTH_ENDPOINT);
    });

    it('should ignore blank OUTLOOK_SCOPES and fall back to config scopes', () => {
      process.env.OUTLOOK_CLIENT_ID = 'client_id_val';
      process.env.OUTLOOK_SCOPES = '   ';
      const config = createAuthConfig();
      expect(config.clientId).toBe('client_id_val');
      expect(config.scopes).toEqual(appConfig.AUTH_CONFIG.scopes);
    });
  });
});
