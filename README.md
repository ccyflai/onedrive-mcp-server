# OneDrive for Business MCP Server

A multi-user MCP (Model Context Protocol) server that provides tools to list, read, write, and search files in OneDrive for Business. It authenticates each user individually using the **OAuth 2.0 device code flow** (delegated permissions) — no application-level secrets are required.

The server uses the **Streamable HTTP** transport (stateless, JSON responses), so multiple MCP clients can connect concurrently over the network.

## Prerequisites

- Python 3.10+
- A Microsoft Entra ID (Azure AD) app registration with **delegated** permissions

## Azure AD App Registration

1. Go to the [Microsoft Entra admin center](https://entra.microsoft.com/) > **App registrations** > **New registration**.
2. Set a name (e.g. "OneDrive MCP Server").
3. Under **Supported account types**, choose "Accounts in this organizational directory only" (single-tenant) or "Accounts in any organizational directory" (multi-tenant) as appropriate.
4. No redirect URI is needed for the device code flow.
5. Click **Register**.
6. Copy the **Application (client) ID** — this is your `ONEDRIVE_CLIENT_ID`.
7. Copy the **Directory (tenant) ID** — this is your `ONEDRIVE_TENANT_ID`.
8. Go to **Authentication** > under **Advanced settings**, set **Allow public client flows** to **Yes**, then **Save**.
9. Go to **API permissions** > **Add a permission** > **Microsoft Graph** > **Delegated permissions**, and add:
   - `Files.ReadWrite.All`
   - `User.Read`
   - `offline_access`
10. Click **Grant admin consent** if you are a tenant admin (optional — users will be prompted on first sign-in).

## Installation

```bash
pip install -r requirements.txt
```

Or with uv:

```bash
uv pip install -r requirements.txt
```

## Configuration

Set the following environment variables:

| Variable | Required | Default | Description |
|---|---|---|---|
| `ONEDRIVE_CLIENT_ID` | Yes | — | Application (client) ID from your Azure AD app registration |
| `ONEDRIVE_TENANT_ID` | No | `organizations` | Directory (tenant) ID, or `organizations` for any work/school account |
| `ONEDRIVE_MCP_HOST` | No | `0.0.0.0` | Host/IP the HTTP server listens on |
| `ONEDRIVE_MCP_PORT` | No | `8000` | Port the HTTP server listens on |
| `ONEDRIVE_TOKEN_CACHE` | No | `~/.onedrive-mcp-token-cache.json` | Path to the persistent token cache file |

## Running the Server

```bash
export ONEDRIVE_CLIENT_ID="your-client-id-here"
export ONEDRIVE_TENANT_ID="your-tenant-id-here"

python server.py
```

The server starts listening on `http://0.0.0.0:8000/mcp` (Streamable HTTP).

### Connecting MCP Clients

The MCP endpoint URL is:

```
http://<host>:<port>/mcp
```

#### Claude Code

```bash
claude mcp add --transport http onedrive http://localhost:8000/mcp
```

#### OpenCode (opencode.toml)

```toml
[mcp.onedrive]
type = "http"
url = "http://localhost:8000/mcp"
```

#### MCP Inspector (for testing)

Start the server, then in a separate terminal:

```bash
npx -y @modelcontextprotocol/inspector
```

Connect to `http://localhost:8000/mcp` in the inspector UI.

## Multi-User Authentication Workflow

Authentication is a two-step process designed for HTTP where long-polling is impractical:

```
Step 1:  Call  start_auth
         → returns { session_id, verification_uri, user_code, message }
         → present the URL and code to the user

Step 2:  User opens the URL in a browser and enters the code.

Step 3:  Call  complete_auth(session_id="...")
         → polls Microsoft until sign-in completes
         → returns { user_id: "alice@contoso.com", ... }

Step 4:  Pass user_id to file-operation tools:
         list_files(user_id="alice@contoso.com", path="/")
         read_file(user_id="alice@contoso.com", path="/Documents/report.txt")
```

Multiple users can authenticate independently. Previously-authenticated users are silently re-authenticated on server restart via cached refresh tokens.

## Available Tools

### Authentication & Session Management

| Tool | Description |
|---|---|
| `start_auth` | Begin device-code sign-in. Returns `session_id`, `verification_uri`, and `user_code`. |
| `complete_auth(session_id)` | Finish sign-in. Polls until the user completes the browser flow. Returns `user_id`. |
| `list_users` | List all authenticated users. |
| `remove_user(user_id)` | Sign out a user (remove cached tokens). |

### File Operations

All file-operation tools require a `user_id` parameter (the email returned by `complete_auth`).

| Tool | Parameters | Description |
|---|---|---|
| `list_files` | `user_id`, `path="/"`*, `page_size=50`* | List files and folders in a directory |
| `read_file` | `user_id`, `path` | Download a file (text returned as-is, binary as base64) |
| `write_file` | `user_id`, `path`, `content`, `encoding="utf-8"`* | Upload/create/overwrite a file (up to 250 MB) |
| `search_files` | `user_id`, `query`, `page_size=25`* | Full-text search across file names and contents |
| `get_file_info` | `user_id`, `path` | Get detailed file/folder metadata |

*Parameters marked with * are optional with the shown defaults.*

## Security Notes

- **Delegated permissions only** — the server acts as the user, not as an application. Each user can only access their own OneDrive files.
- **No client secret** needed. The app registration is a public client.
- **Token cache** at `~/.onedrive-mcp-token-cache.json` contains refresh tokens for all users. Protect this file — anyone with access can impersonate cached users.
- **`offline_access`** scope provides refresh tokens so users don't need to re-authenticate across restarts.
- **Stateless HTTP** — the server holds no per-session state on the HTTP side. Pending device-code flows are held in memory and expire after 15 minutes.
- Call `remove_user` to explicitly revoke a user's cached session.
