# OneNote MCP Server

A Model Context Protocol (MCP) server that enables AI assistants to interact with Microsoft OneNote.

> Based on [azure-onenote-mcp-server](https://github.com/ZubeidHendricks/azure-onenote-mcp-server) by Zubeid Hendricks, with simplified authentication and improved usability.

## What Does This Do?

- List your OneNote notebooks, sections, and pages
- Create new pages with HTML content
- Search through your notes
- Read complete page content
- Access notebooks shared by other users
- Analyze and summarize notes directly through the AI interface

## Installation

### Prerequisites

- Node.js 16+ ([nodejs.org](https://nodejs.org/))
- A Microsoft account with OneNote access
- Git ([git-scm.com](https://git-scm.com/))

### Step 1: Clone & Install Dependencies

```bash
git clone https://github.com/yourusername/onenote-mcp.git
cd onenote-mcp
npm install
```

> **Note:** The TypeScript SDK dependency is included in this repo. If you need to rebuild it: `cd typescript-sdk && npm install && npm run build && cd ..`

### Step 2: Start the Server

```bash
npm start
```

You should see:
```
Server started successfully.
Use the "authenticate" tool to sign in to OneNote.
```

## Authentication

### Option A: Interactive Device Code Flow (recommended for first use)

Ask your AI assistant to authenticate with OneNote. The AI will provide a URL and code — open the URL, enter the code, and sign in with your Microsoft account. The token is saved for future use.

### Option B: Pre-configured Token (skip auth dialog)

Get a token from [Microsoft Graph Explorer](https://developer.microsoft.com/graph/graph-explorer):

1. Sign in, run `GET https://graph.microsoft.com/v1.0/me`
2. Click **Modify Permissions**, search for `Notes`, check `Notes.ReadWrite`
3. Click **Consent**, authorize, then copy the token from the **Access Token** tab
4. Add it to your MCP config (see below)

## MCP Configuration

All tools support the same JSON config format. The differences are where to put it.

### opencode

Add to `~/.config/opencode/opencode.json` (or `.opencode/opencode.json` in project):

```json
{
  "mcpServers": {
    "onenote": {
      "command": "node",
      "args": ["/absolute/path/to/onenote-mcp.mjs"],
      "env": {}
    }
  }
}
```

### Claude Desktop

Add to `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "onenote": {
      "command": "node",
      "args": ["/absolute/path/to/onenote-mcp.mjs"],
      "env": {}
    }
  }
}
```

### Cursor

1. Open Cursor Settings → MCP tab
2. Add a new server:
   - **Name:** `onenote`
   - **Command:** `node`
   - **Args:** `["/absolute/path/to/onenote-mcp.mjs"]`

### VS Code (via GitHub Copilot MCP)

Add to VS Code's MCP settings (`settings.json` or MCP config file):

```json
{
  "mcpServers": {
    "onenote": {
      "command": "node",
      "args": ["/absolute/path/to/onenote-mcp.mjs"],
      "env": {}
    }
  }
}
```

Requires [VS Code Insiders](https://code.visualstudio.com/insiders/) with GitHub Copilot MCP support enabled.

### Gemini CLI

Add to your Gemini CLI config file:

```json
{
  "mcpServers": {
    "onenote": {
      "command": "node",
      "args": ["/absolute/path/to/onenote-mcp.mjs"],
      "env": {}
    }
  }
}
```

### Skip Authentication on Startup

If you already have a token, add it to the config to skip the auth dialog:

```json
{
  "mcpServers": {
    "onenote": {
      "command": "node",
      "args": ["/absolute/path/to/onenote-mcp.mjs", "--token", "YOUR_TOKEN_HERE"],
      "env": {}
    }
  }
}
```

Or use the `GRAPH_ACCESS_TOKEN` environment variable:

```json
{
  "mcpServers": {
    "onenote": {
      "command": "node",
      "args": ["/absolute/path/to/onenote-mcp.mjs"],
      "env": {
        "GRAPH_ACCESS_TOKEN": "YOUR_TOKEN_HERE"
      }
    }
  }
}
```

## Available MCP Tools

| Tool | Description |
|------|-------------|
| `authenticate` | Start the Microsoft device code authentication flow |
| `saveAccessToken` | Save a manually obtained Graph API token |
| `listNotebooks` | List all your OneNote notebooks |
| `getNotebook` | Get details of a specific notebook (optional `notebookId`) |
| `getSharedNotebook` | Access a notebook shared by another user via its Web URL |
| `listSections` | List sections in a notebook (optional `notebookId`) |
| `listPages` | List pages in a section (optional `sectionId`) |
| `getPage` | Get the full HTML content of a page |
| `createPage` | Create a new page with HTML content |
| `updatePage` | Update an existing page's content and title |
| `searchPages` | Search pages by title |

## Example Interactions

```
Show me my OneNote notebooks
What sections are in my Projects notebook?
Create a new page titled "Meeting Notes" in my Work notebook
Find all notes about machine learning
Read and summarize my "Project Requirements" page
Access the shared notebook at https://1drv.ms/o/s/...
```

## Troubleshooting

- **Authentication fails** — try a different browser, clear cookies/cache
- **"expired_token"** — re-authenticate
- **Server won't start** — verify Node.js 16+ and `npm install` completed
- **AI can't connect** — make sure the server is running and the config path is correct

## Security Notes

- Tokens are stored in `.access-token.txt` (gitignored)
- Tokens grant access to your OneNote data — keep them secure
- Tokens expire after some time, requiring re-authentication
- No Azure setup or API keys required

## License

MIT