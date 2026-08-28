import { PublicClientApplication } from '@azure/msal-node';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const msalCachePath = path.join(__dirname, '.msal-cache.json');
const tokenFilePath = path.join(__dirname, '.access-token.txt');

const clientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e';
const scopes = ['Notes.Read.All', 'Notes.ReadWrite.All', 'User.Read'];

const msalApp = new PublicClientApplication({
  auth: {
    clientId: clientId,
    authority: 'https://login.microsoftonline.com/common'
  },
  cache: {
    cachePlugin: {
      beforeCacheAccess: async (cacheContext) => {
        try {
          const cache = JSON.parse(fs.readFileSync(msalCachePath, 'utf8'));
          cacheContext.tokenCache.deserialize(cache);
        } catch {}
      },
      afterCacheAccess: async (cacheContext) => {
        if (cacheContext.cacheHasChanged) {
          fs.writeFileSync(msalCachePath, JSON.stringify(cacheContext.tokenCache.serialize()));
        }
      }
    }
  }
});

async function authenticate() {
  try {
    const accounts = await msalApp.getTokenCache().getAllAccounts();
    if (accounts.length > 0) {
      console.log('Found existing session. Attempting silent token refresh...');
      try {
        const response = await msalApp.acquireTokenSilent({
          scopes: scopes,
          account: accounts[0]
        });
        console.log('Token refreshed successfully. No need to re-authenticate.');
        console.log('You can now use the MCP server.');
        return;
      } catch (silentError) {
        console.log('Silent refresh failed, starting device code flow...');
      }
    }

    console.log('Starting authentication...');
    console.log('Check your browser or go to https://login.microsoftonline.com/device\n');

    const response = await msalApp.acquireTokenByDeviceCode({
      scopes: scopes,
      deviceCodeCallback: (info) => {
        console.log(info.message);
      }
    });

    console.log('\nAuthentication successful!');
    console.log('The token cache has been saved to:', msalCachePath);
    console.log('You can now use the MCP server - it will automatically reuse this session.');

    fs.writeFileSync(tokenFilePath, JSON.stringify({ token: response.accessToken }));

  } catch (error) {
    console.error('Authentication error:', error);
  }
}

authenticate();