import { ProxyOAuthServerProvider } from '@modelcontextprotocol/sdk/server/auth/providers/proxyProvider.js';
import logger from './logger.js';
export class MicrosoftOAuthProvider extends ProxyOAuthServerProvider {
    constructor(authManager) {
        const tenantId = process.env.MS365_MCP_TENANT_ID || 'common';
        const clientId = process.env.MS365_MCP_CLIENT_ID || '084a3e9f-a9f4-43f7-89f9-d229cf97853e';
        super({
            endpoints: {
                authorizationUrl: `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`,
                tokenUrl: `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
                revocationUrl: `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/logout`,
            },
            verifyAccessToken: async (token) => {
                try {
                    const response = await fetch('https://graph.microsoft.com/v1.0/me', {
                        headers: {
                            Authorization: `Bearer ${token}`,
                        },
                    });
                    if (response.ok) {
                        const userData = await response.json();
                        logger.info(`OAuth token verified for user: ${userData.userPrincipalName}`);
                        await authManager.setOAuthToken(token);
                        return {
                            token,
                            clientId,
                            scopes: [],
                        };
                    }
                    else {
                        throw new Error(`Token verification failed: ${response.status}`);
                    }
                }
                catch (error) {
                    logger.error(`OAuth token verification error: ${error}`);
                    throw error;
                }
            },
            getClient: async (client_id) => {
                return {
                    client_id,
                    redirect_uris: ['http://localhost:3000/callback'],
                };
            },
        });
        this.authManager = authManager;
    }
}
