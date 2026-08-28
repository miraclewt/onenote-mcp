#!/usr/bin/env node

import { McpServer } from './typescript-sdk/dist/esm/server/mcp.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { StdioServerTransport } from './typescript-sdk/dist/esm/server/stdio.js';
import { PublicClientApplication } from '@azure/msal-node';
import dotenv from 'dotenv';
import { fileURLToPath } from 'url';
import path from 'path';
import fs from 'fs';
import fetch from 'node-fetch';
import { z } from 'zod';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const tokenFilePath = path.join(__dirname, '.access-token.txt');
const msalCachePath = path.join(__dirname, '.msal-cache.json');

const server = new McpServer(
  {
    name: "onenote",
    version: "1.0.0",
    description: "OneNote MCP Server"
  },
  {
    capabilities: {
      tools: {
        listChanged: true
      }
    }
  }
);

let accessToken = null;
let graphClient = null;

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

async function acquireTokenFromCache() {
  try {
    const accounts = await msalApp.getTokenCache().getAllAccounts();
    if (accounts.length > 0) {
      const response = await msalApp.acquireTokenSilent({
        scopes: scopes,
        account: accounts[0]
      });
      return response.accessToken;
    }
  } catch (error) {
    console.error('Silent token acquisition failed:', error.message);
  }
  return null;
}

async function acquireTokenFromDeviceCode() {
  const response = await msalApp.acquireTokenByDeviceCode({
    scopes: scopes,
    deviceCodeCallback: (response) => {
      console.error('\n' + response.message);
    }
  });
  return response.accessToken;
}

function readSavedToken() {
  try {
    if (fs.existsSync(tokenFilePath)) {
      const tokenData = fs.readFileSync(tokenFilePath, 'utf8');
      try {
        const parsedToken = JSON.parse(tokenData);
        return parsedToken.token;
      } catch {
        return tokenData;
      }
    }
  } catch {}
  return null;
}

if (process.env.GRAPH_ACCESS_TOKEN) {
  accessToken = process.env.GRAPH_ACCESS_TOKEN;
}

function initGraphClient(token) {
  graphClient = Client.initWithMiddleware({
    authProvider: {
      getAccessToken: async () => token
    }
  });
}

async function ensureGraphClient() {
  if (graphClient) return graphClient;

  if (accessToken) {
    initGraphClient(accessToken);
    return graphClient;
  }

  let token = await acquireTokenFromCache();
  if (token) {
    accessToken = token;
    initGraphClient(token);
    return graphClient;
  }

  token = readSavedToken();
  if (token) {
    accessToken = token;
    initGraphClient(token);
    return graphClient;
  }

  throw new Error("Access token not found. Use the 'authenticate' tool to sign in.");
}

// Tool: authenticate
server.tool(
  "authenticate",
  "Sign in to Microsoft OneNote. You will receive a URL and code to complete login in your browser.",
  async () => {
    try {
      const token = await acquireTokenFromDeviceCode();
      accessToken = token;
      initGraphClient(token);
      fs.writeFileSync(tokenFilePath, JSON.stringify({ token: accessToken }));
      return {
        content: [
          {
            type: "text",
            text: "Authentication successful! You are now signed in to OneNote."
          }
        ]
      };
    } catch (error) {
      console.error("Authentication error:", error);
      throw new Error(`Authentication failed: ${error.message}`);
    }
  }
);

// Tool: saveAccessToken
server.tool(
  "saveAccessToken",
  "Save a Microsoft Graph access token manually. Get a token from https://developer.microsoft.com/graph/graph-explorer",
  {
    token: z.string().describe("Microsoft Graph access token")
  },
  async (params) => {
    try {
      const token = params.token.trim();
      const testClient = Client.initWithMiddleware({
        authProvider: {
          getAccessToken: async () => token
        }
      });
      const me = await testClient.api("/me").get();

      accessToken = token;
      graphClient = testClient;
      fs.writeFileSync(tokenFilePath, JSON.stringify({ token: accessToken }));

      return {
        content: [
          {
            type: "text",
            text: `Access token saved. Authenticated as: ${me.displayName} (${me.userPrincipalName || me.mail || "unknown"})`
          }
        ]
      };
    } catch (error) {
      console.error("Error saving access token:", error);
      throw new Error(`Failed to save access token: ${error.message}. Make sure the token is valid and has Notes permissions.`);
    }
  }
);

// Tool: listNotebooks
server.tool(
  "listNotebooks",
  "List all OneNote notebooks",
  async () => {
    try {
      await ensureGraphClient();
      const response = await graphClient.api("/me/onenote/notebooks").get();
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(response.value)
          }
        ]
      };
    } catch (error) {
      console.error("Error listing notebooks:", error);
      throw new Error(`Failed to list notebooks: ${error.message}`);
    }
  }
);

// Tool: getNotebook
server.tool(
  "getNotebook",
  "Get details of a specific notebook",
  async () => {
    try {
      await ensureGraphClient();
      const response = await graphClient.api("/me/onenote/notebooks").get();
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(response.value[0])
          }
        ]
      };
    } catch (error) {
      console.error("Error getting notebook:", error);
      throw new Error(`Failed to get notebook: ${error.message}`);
    }
  }
);

// Tool: listSections
server.tool(
  "listSections",
  "List all sections in a notebook",
  async () => {
    try {
      await ensureGraphClient();
      const response = await graphClient.api("/me/onenote/sections").get();
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(response.value)
          }
        ]
      };
    } catch (error) {
      console.error("Error listing sections:", error);
      throw new Error(`Failed to list sections: ${error.message}`);
    }
  }
);

// Tool: listPages
server.tool(
  "listPages",
  "List all pages in a section",
  async () => {
    try {
      await ensureGraphClient();
      const sectionsResponse = await graphClient.api("/me/onenote/sections").get();

      if (sectionsResponse.value.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: "[]"
            }
          ]
        };
      }

      const sectionId = sectionsResponse.value[0].id;
      const response = await graphClient.api(`/me/onenote/sections/${sectionId}/pages`).get();

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(response.value)
          }
        ]
      };
    } catch (error) {
      console.error("Error listing pages:", error);
      throw new Error(`Failed to list pages: ${error.message}`);
    }
  }
);

// Tool: getPage
server.tool(
  "getPage",
  "Get the content of a page",
  async (params) => {
    try {
      console.error("GetPage called with params:", params);
      await ensureGraphClient();

      const pagesResponse = await graphClient.api('/me/onenote/pages').get();
      console.error("Got", pagesResponse.value.length, "pages");

      let targetPage;

      if (params.random_string && params.random_string.length > 0) {
        const pageId = params.random_string;
        console.error("Looking for page with ID:", pageId);

        targetPage = pagesResponse.value.find(p => p.id === pageId);

        if (!targetPage) {
          console.error("No exact match, trying title search");
          targetPage = pagesResponse.value.find(p =>
            p.title && p.title.toLowerCase().includes(params.random_string.toLowerCase())
          );
        }

        if (!targetPage) {
          console.error("No title match, trying partial ID match");
          targetPage = pagesResponse.value.find(p =>
            p.id.includes(pageId) || pageId.includes(p.id)
          );
        }
      } else {
        console.error("No ID provided, using first page");
        targetPage = pagesResponse.value[0];
      }

      if (!targetPage) {
        throw new Error("Page not found");
      }

      console.error("Target page found:", targetPage.title);
      console.error("Page ID:", targetPage.id);

      try {
        const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${targetPage.id}/content`;
        console.error("Fetching content from:", url);

        const response = await fetch(url, {
          headers: {
            'Authorization': `Bearer ${accessToken}`
          }
        });

        if (!response.ok) {
          throw new Error(`HTTP error! Status: ${response.status} ${response.statusText}`);
        }

        const content = await response.text();
        console.error(`Content received! Length: ${content.length} characters`);

        return {
          content: [
            {
              type: "text",
              text: content
            }
          ]
        };
      } catch (error) {
        console.error("Error getting content:", error);

        return {
          content: [
            {
              type: "text",
              text: `Error retrieving page content: ${error.message}`
            }
          ]
        };
      }
    } catch (error) {
      console.error("Error in getPage:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error in getPage: ${error.message}`
          }
        ]
      };
    }
  }
);

// Tool: createPage
server.tool(
  "createPage",
  "Create a new page in a section",
  async () => {
    try {
      await ensureGraphClient();
      const sectionsResponse = await graphClient.api("/me/onenote/sections").get();

      if (sectionsResponse.value.length === 0) {
        throw new Error("No sections found");
      }

      const sectionId = sectionsResponse.value[0].id;

      const simpleHtml = `
        <!DOCTYPE html>
        <html>
          <head>
            <title>New Page</title>
          </head>
          <body>
            <p>This is a new page created via the Microsoft Graph API</p>
          </body>
        </html>
      `;

      const response = await graphClient
        .api(`/me/onenote/sections/${sectionId}/pages`)
        .header("Content-Type", "application/xhtml+xml")
        .post(simpleHtml);

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(response)
          }
        ]
      };
    } catch (error) {
      console.error("Error creating page:", error);
      throw new Error(`Failed to create page: ${error.message}`);
    }
  }
);

// Tool: searchPages
server.tool(
  "searchPages",
  "Search for pages across notebooks",
  async (params) => {
    try {
      await ensureGraphClient();

      const response = await graphClient.api("/me/onenote/pages").get();

      if (params.random_string && params.random_string.length > 0) {
        const searchTerm = params.random_string.toLowerCase();
        const filteredPages = response.value.filter(page => {
          if (page.title && page.title.toLowerCase().includes(searchTerm)) {
            return true;
          }
          return false;
        });

        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(filteredPages)
            }
          ]
        };
      } else {
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify(response.value)
            }
          ]
        };
      }
    } catch (error) {
      console.error("Error searching pages:", error);
      throw new Error(`Failed to search pages: ${error.message}`);
    }
  }
);

async function main() {
  try {
    const transport = new StdioServerTransport();
    await server.connect(transport);

    console.error('Server started successfully.');
    console.error('Use the "authenticate" tool to sign in to OneNote.');
    console.error('If you already signed in before, the token will be reused automatically.');

    process.on('SIGINT', () => {
      process.exit(0);
    });
  } catch (error) {
    console.error('Error starting server:', error);
    process.exit(1);
  }
}

main();