import { PublicClientApplication } from '@azure/msal-node';
import keytar from 'keytar';
import { fileURLToPath } from 'url';
import path from 'path';
import fs from 'fs';
import logger from './logger.js';
const endpoints = await import('./endpoints.json', {
    with: { type: 'json' },
});
const SERVICE_NAME = 'ms-365-mcp-server';
const TOKEN_CACHE_ACCOUNT = 'msal-token-cache';
const FALLBACK_DIR = path.dirname(fileURLToPath(import.meta.url));
const FALLBACK_PATH = path.join(FALLBACK_DIR, '..', '.token-cache.json');
const DEFAULT_CONFIG = {
    auth: {
        clientId: process.env.MS365_MCP_CLIENT_ID || '084a3e9f-a9f4-43f7-89f9-d229cf97853e',
        authority: `https://login.microsoftonline.com/${process.env.MS365_MCP_TENANT_ID || 'common'}`,
    },
};
const SCOPE_HIERARCHY = {
    'Mail.ReadWrite': ['Mail.Read'],
    'Calendars.ReadWrite': ['Calendars.Read'],
    'Files.ReadWrite': ['Files.Read'],
    'Tasks.ReadWrite': ['Tasks.Read'],
    'Contacts.ReadWrite': ['Contacts.Read'],
};
function buildScopesFromEndpoints(includeWorkAccountScopes = false) {
    const scopesSet = new Set();
    endpoints.default.forEach((endpoint) => {
        if (endpoint.requiresWorkAccount && !includeWorkAccountScopes) {
            return;
        }
        if (endpoint.scopes && Array.isArray(endpoint.scopes)) {
            endpoint.scopes.forEach((scope) => scopesSet.add(scope));
        }
    });
    Object.entries(SCOPE_HIERARCHY).forEach(([higherScope, lowerScopes]) => {
        if (lowerScopes.every((scope) => scopesSet.has(scope))) {
            lowerScopes.forEach((scope) => scopesSet.delete(scope));
            scopesSet.add(higherScope);
        }
    });
    return Array.from(scopesSet);
}
function buildAllScopes() {
    return buildScopesFromEndpoints(true);
}
class AuthManager {
    constructor(config = DEFAULT_CONFIG, scopes = buildScopesFromEndpoints()) {
        logger.info(`And scopes are ${scopes.join(', ')}`, scopes);
        this.config = config;
        this.scopes = scopes;
        this.msalApp = new PublicClientApplication(this.config);
        this.accessToken = null;
        this.tokenExpiry = null;
        const oauthTokenFromEnv = process.env.MS365_MCP_OAUTH_TOKEN;
        this.oauthToken = oauthTokenFromEnv ?? null;
        this.isOAuthMode = oauthTokenFromEnv != null;
    }
    async loadTokenCache() {
        try {
            let cacheData;
            try {
                const cachedData = await keytar.getPassword(SERVICE_NAME, TOKEN_CACHE_ACCOUNT);
                if (cachedData) {
                    cacheData = cachedData;
                }
            }
            catch (keytarError) {
                logger.warn(`Keychain access failed, falling back to file storage: ${keytarError.message}`);
            }
            if (!cacheData && fs.existsSync(FALLBACK_PATH)) {
                cacheData = fs.readFileSync(FALLBACK_PATH, 'utf8');
            }
            if (cacheData) {
                this.msalApp.getTokenCache().deserialize(cacheData);
            }
        }
        catch (error) {
            logger.error(`Error loading token cache: ${error.message}`);
        }
    }
    async saveTokenCache() {
        try {
            const cacheData = this.msalApp.getTokenCache().serialize();
            try {
                await keytar.setPassword(SERVICE_NAME, TOKEN_CACHE_ACCOUNT, cacheData);
            }
            catch (keytarError) {
                logger.warn(`Keychain save failed, falling back to file storage: ${keytarError.message}`);
                fs.writeFileSync(FALLBACK_PATH, cacheData);
            }
        }
        catch (error) {
            logger.error(`Error saving token cache: ${error.message}`);
        }
    }
    async setOAuthToken(token) {
        this.oauthToken = token;
        this.isOAuthMode = true;
    }
    async getToken(forceRefresh = false) {
        if (this.isOAuthMode && this.oauthToken) {
            return this.oauthToken;
        }
        if (this.accessToken && this.tokenExpiry && this.tokenExpiry > Date.now() && !forceRefresh) {
            return this.accessToken;
        }
        const accounts = await this.msalApp.getTokenCache().getAllAccounts();
        if (accounts.length > 0) {
            const silentRequest = {
                account: accounts[0],
                scopes: this.scopes,
            };
            try {
                const response = await this.msalApp.acquireTokenSilent(silentRequest);
                this.accessToken = response.accessToken;
                this.tokenExpiry = response.expiresOn ? new Date(response.expiresOn).getTime() : null;
                return this.accessToken;
            }
            catch (error) {
                logger.error('Silent token acquisition failed');
                throw new Error('Silent token acquisition failed');
            }
        }
        throw new Error('No valid token found');
    }
    async acquireTokenByDeviceCode(hack) {
        const deviceCodeRequest = {
            scopes: this.scopes,
            deviceCodeCallback: (response) => {
                const text = ['\n', response.message, '\n'].join('');
                if (hack) {
                    hack(text + 'After login run the "verify login" command');
                }
                else {
                    console.log(text);
                }
                logger.info('Device code login initiated');
            },
        };
        try {
            logger.info('Requesting device code...');
            logger.info(`Requesting scopes: ${this.scopes.join(', ')}`);
            const response = await this.msalApp.acquireTokenByDeviceCode(deviceCodeRequest);
            logger.info(`Granted scopes: ${response?.scopes?.join(', ') || 'none'}`);
            logger.info('Device code login successful');
            this.accessToken = response?.accessToken || null;
            this.tokenExpiry = response?.expiresOn ? new Date(response.expiresOn).getTime() : null;
            await this.saveTokenCache();
            return this.accessToken;
        }
        catch (error) {
            logger.error(`Error in device code flow: ${error.message}`);
            throw error;
        }
    }
    async testLogin() {
        try {
            logger.info('Testing login...');
            const token = await this.getToken();
            if (!token) {
                logger.error('Login test failed - no token received');
                return {
                    success: false,
                    message: 'Login failed - no token received',
                };
            }
            logger.info('Token retrieved successfully, testing Graph API access...');
            try {
                const response = await fetch('https://graph.microsoft.com/v1.0/me', {
                    headers: {
                        Authorization: `Bearer ${token}`,
                    },
                });
                if (response.ok) {
                    const userData = await response.json();
                    logger.info('Graph API user data fetch successful');
                    return {
                        success: true,
                        message: 'Login successful',
                        userData: {
                            displayName: userData.displayName,
                            userPrincipalName: userData.userPrincipalName,
                        },
                    };
                }
                else {
                    const errorText = await response.text();
                    logger.error(`Graph API user data fetch failed: ${response.status} - ${errorText}`);
                    return {
                        success: false,
                        message: `Login successful but Graph API access failed: ${response.status}`,
                    };
                }
            }
            catch (graphError) {
                logger.error(`Error fetching user data: ${graphError.message}`);
                return {
                    success: false,
                    message: `Login successful but Graph API access failed: ${graphError.message}`,
                };
            }
        }
        catch (error) {
            logger.error(`Login test failed: ${error.message}`);
            return {
                success: false,
                message: `Login failed: ${error.message}`,
            };
        }
    }
    async logout() {
        try {
            const accounts = await this.msalApp.getTokenCache().getAllAccounts();
            for (const account of accounts) {
                await this.msalApp.getTokenCache().removeAccount(account);
            }
            this.accessToken = null;
            this.tokenExpiry = null;
            try {
                await keytar.deletePassword(SERVICE_NAME, TOKEN_CACHE_ACCOUNT);
            }
            catch (keytarError) {
                logger.warn(`Keychain deletion failed: ${keytarError.message}`);
            }
            if (fs.existsSync(FALLBACK_PATH)) {
                fs.unlinkSync(FALLBACK_PATH);
            }
            return true;
        }
        catch (error) {
            logger.error(`Error during logout: ${error.message}`);
            throw error;
        }
    }
    async hasWorkAccountPermissions() {
        try {
            const accounts = await this.msalApp.getTokenCache().getAllAccounts();
            if (accounts.length === 0) {
                return false;
            }
            const workScopes = endpoints.default
                .filter((e) => e.requiresWorkAccount)
                .flatMap((e) => e.scopes || []);
            try {
                await this.msalApp.acquireTokenSilent({
                    scopes: workScopes.slice(0, 1),
                    account: accounts[0],
                });
                return true;
            }
            catch {
                return false;
            }
        }
        catch (error) {
            logger.error(`Error checking work account permissions: ${error.message}`);
            return false;
        }
    }
    async expandToWorkAccountScopes(hack) {
        try {
            logger.info('Expanding to work account scopes...');
            const allScopes = buildAllScopes();
            const deviceCodeRequest = {
                scopes: allScopes,
                deviceCodeCallback: (response) => {
                    const text = [
                        '\n',
                        '🔄 This feature requires additional permissions (work account scopes)',
                        '\n',
                        response.message,
                        '\n',
                    ].join('');
                    if (hack) {
                        hack(text + 'After login run the "verify login" command');
                    }
                    else {
                        console.log(text);
                    }
                    logger.info('Work account scope expansion initiated');
                },
            };
            const response = await this.msalApp.acquireTokenByDeviceCode(deviceCodeRequest);
            logger.info('Work account scope expansion successful');
            this.accessToken = response?.accessToken || null;
            this.tokenExpiry = response?.expiresOn ? new Date(response.expiresOn).getTime() : null;
            this.scopes = allScopes;
            await this.saveTokenCache();
            return true;
        }
        catch (error) {
            logger.error(`Error expanding to work account scopes: ${error.message}`);
            return false;
        }
    }
    requiresWorkAccountScope(toolName) {
        const endpoint = endpoints.default.find((e) => e.toolName === toolName);
        return endpoint?.requiresWorkAccount === true;
    }
}
export default AuthManager;
export { buildScopesFromEndpoints, buildAllScopes };
