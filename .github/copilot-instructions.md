# Copilot Cloud Agent Instructions — OneDrive for Business MCP Server

## Repository Overview

This is a **single-file Python MCP server** (`server.py`) that exposes OneDrive for Business as a set of MCP tools. All logic lives in `server.py`; there are no sub-packages, no test suite, no linting configuration, and no CI workflows. The only other files are `README.md`, `AGENTS.md`, `requirements.txt`, and this instructions file.

## Project Layout

```
server.py           # Entire server implementation (~700 lines)
requirements.txt    # mcp>=1.0.0, msal>=1.28.0, httpx>=0.27.0
README.md           # User-facing documentation
AGENTS.md           # Agent-facing summary (keep in sync with changes)
.github/
  copilot-instructions.md  # This file
```

## Running and Testing Locally

**There is no test suite and no linter configured.** Do not attempt to run `pytest`, `ruff`, or `mypy` — none are installed or configured.

To verify the server starts without errors:

```bash
pip install -r requirements.txt
export ONEDRIVE_CLIENT_ID="fake-client-id"
python server.py
```

> **Known issue / workaround:** Starting the server with a fake `ONEDRIVE_CLIENT_ID` will make it bind on port 8000 successfully. The `_get_msal_app()` function is only called lazily (on the first tool invocation), so the server process itself starts cleanly. If `ONEDRIVE_CLIENT_ID` is empty the server still starts — the `RuntimeError` is only raised when a tool is actually called.

To do a quick syntax/import check without starting the HTTP listener, run:

```bash
python -c "import server; print('OK')"
```

## Code Conventions (must follow for all changes)

1. **Type annotations** are required on every function signature. Use `X | Y` union syntax (PEP 604), not `Optional[X]` or `Union[X, Y]`.
2. **Tool return values** are `str`:
   - All tools return `json.dumps(result, indent=2)` except `read_file`.
   - `read_file` returns raw UTF-8 text for small (<10 MB) text files, or a JSON metadata line followed by `"base64:<b64data>"` for binary/large files.
3. **Error handling**:
   - Raise `RuntimeError` for configuration/auth failures (missing env var, expired token, bad session ID).
   - Return `"Error: ..."` strings for user-input validation errors (bad path, wrong type, etc.).
4. **Async patterns**:
   - Blocking MSAL calls must be wrapped with `asyncio.to_thread()`.
   - HTTP calls use `httpx.AsyncClient` instantiated per-request with `timeout=60.0`.
5. **No new top-level modules** — keep all logic in `server.py`.
6. **No new dependencies** unless absolutely necessary; if added, update `requirements.txt`.

## Architecture Details

### Authentication (MSAL)

- Uses `msal.PublicClientApplication` — no client secret, device-code flow only.
- A single global `_msal_app` is lazily created by `_get_msal_app()`.
- Token cache is a `msal.SerializableTokenCache` serialized to `~/.onedrive-mcp-token-cache.json` (path overridable via `ONEDRIVE_TOKEN_CACHE`).
- Pending device-code flows are stored in the in-memory dict `_pending_flows`, keyed by a random `session_id` (UUID hex). They expire after 15 minutes (`_FLOW_TTL = 900`). `_gc_pending_flows()` purges expired entries.
- `_acquire_token_for_user(user_id)` does a silent token refresh; raises `RuntimeError` if the account is unknown or the refresh token has expired.

### Graph API

- Base URL: `https://graph.microsoft.com/v1.0`
- All requests go through `_graph_request(user_id, method, path, ...)`.
- Paths starting with `/` are appended to `GRAPH_BASE`; absolute URLs are used as-is (needed for `@odata.nextLink` pagination).
- `resp.raise_for_status()` is called after every request — Graph errors become `httpx.HTTPStatusError`.

### MCP Server

- Built with `mcp.server.fastmcp.FastMCP`.
- Configured as `stateless_http=True, json_response=True` — each HTTP request is fully independent.
- Entrypoint: `main()` → `mcp.run(transport="streamable-http")`.
- Listens on `http://<ONEDRIVE_MCP_HOST>:<ONEDRIVE_MCP_PORT>/mcp` (defaults: `0.0.0.0:8000`).

### MCP Tools (all in `server.py`)

| Tool | Purpose |
|---|---|
| `start_auth()` | Initiate device-code flow; returns `session_id`, `verification_uri`, `user_code` |
| `complete_auth(session_id)` | Poll until sign-in completes; returns `user_id` (email) |
| `list_users()` | List all cached authenticated users |
| `remove_user(user_id)` | Remove a user's cached tokens |
| `list_files(user_id, path, page_size)` | List directory contents; supports pagination via `nextLink`/`hasMore` |
| `read_file(user_id, path)` | Download a file; text returned raw, binary as base64 |
| `write_file(user_id, path, content, encoding)` | Upload/overwrite a file via Graph simple upload |
| `search_files(user_id, query, page_size)` | Full-text search; supports pagination via `hasMore` |
| `get_file_info(user_id, path)` | Get detailed file/folder metadata |

## Known Limitations and Gotchas

1. **`write_file` is limited to ~4 MB.** It uses the Graph simple upload endpoint (`PUT /me/drive/root:/{path}:/content`). Despite the README saying 250 MB, there is no resumable-upload implementation. Attempting to upload files larger than ~4 MB will receive an HTTP error from Graph.

2. **Pagination is not automatic.** `list_files` and `search_files` return a `nextLink` URL and `hasMore: true` when more results exist. Callers must pass the `nextLink` URL as the `path` argument in a subsequent `_graph_request` call — there is no built-in auto-pagination.

3. **`_pending_flows` is in-memory only.** Restarting the server loses all pending auth sessions. Users must call `start_auth` again after a restart.

4. **Token cache file is sensitive.** It contains refresh tokens for all users. Never commit it. It lives at `~/.onedrive-mcp-token-cache.json` by default.

5. **`complete_auth` blocks the event loop thread** (via `asyncio.to_thread`) while polling Microsoft's token endpoint — this is intentional and correct.

6. **`_graph_request` does not handle `nextLink` pagination automatically** — the caller must pass the full `nextLink` URL and call `_graph_request` directly with `user_id, "GET", next_link_url`.

## Environment Variables

| Variable | Required | Default | Notes |
|---|---|---|---|
| `ONEDRIVE_CLIENT_ID` | **Yes** | — | Azure AD Application (client) ID |
| `ONEDRIVE_TENANT_ID` | No | `organizations` | Tenant ID or `organizations` for multi-tenant |
| `ONEDRIVE_MCP_HOST` | No | `0.0.0.0` | Bind address |
| `ONEDRIVE_MCP_PORT` | No | `8000` | Bind port (integer) |
| `ONEDRIVE_TOKEN_CACHE` | No | `~/.onedrive-mcp-token-cache.json` | Token cache file path |

## When Making Changes

- **Edit only `server.py`** for all functional changes.
- **Update `README.md`** if the public-facing behavior of any tool changes (parameters, return shape, limits).
- **Update `AGENTS.md`** if the architecture, conventions, or known gotchas change.
- After editing, verify there are no syntax errors: `python -c "import server; print('OK')"`.
- Do not add a test framework, linter, or CI unless explicitly asked — the project intentionally has none.
