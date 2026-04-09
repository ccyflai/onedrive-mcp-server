# AGENTS.md — OneDrive for Business MCP Server

## Setup

```bash
pip install -r requirements.txt   # mcp, msal, httpx
export ONEDRIVE_CLIENT_ID="<client-id>"
export ONEDRIVE_TENANT_ID="<tenant-id>"   # optional, defaults to "organizations"
python server.py
```

Requires Python 3.10+. Server listens on `http://0.0.0.0:8000/mcp` (Streamable HTTP).

**Auth flow:** `start_auth` → present URL/code to user → `complete_auth` → use returned `user_id` in all file tools.

## Environment Variables

| Variable | Required | Default |
|---|---|---|
| `ONEDRIVE_CLIENT_ID` | Yes | — |
| `ONEDRIVE_TENANT_ID` | No | `organizations` |
| `ONEDRIVE_MCP_HOST` | No | `0.0.0.0` |
| `ONEDRIVE_MCP_PORT` | No | `8000` |
| `ONEDRIVE_TOKEN_CACHE` | No | `~/.onedrive-mcp-token-cache.json` |

## Architecture

- **Single-file server** — all logic in `server.py`, entrypoint `main()` → `mcp.run(transport="streamable-http")`
- **Multi-user** — every file op requires `user_id` (email returned by `complete_auth`)
- **Token cache on disk** — contains refresh tokens; treat as sensitive
- **No client secret** — `PublicClientApplication` (device-code flow only)
- **Required Graph scopes:** `Files.ReadWrite.All`, `User.Read`, `offline_access`
- **Stateless HTTP** — pending auth flows in-memory, expire after 15 min

## Code Conventions

- Type annotations required on all signatures (use `X | Y` union syntax, not `Optional[X]`)
- Public tool functions return `json.dumps(result, indent=2)` as `str` — **exception:** `read_file` returns raw text for small text files (< 10 MB, text MIME or known text extension), or a JSON header line + `"base64:"` prefix + base64 payload for binary/large files
- Error handling: raise `RuntimeError` for config/auth failures; return `"Error: ..."` strings for user input validation
- Wrap blocking MSAL calls with `asyncio.to_thread()`; use `httpx.AsyncClient` per request with `timeout=60.0`
- `write_file` with binary data: caller passes `encoding="base64"` and base64-encoded content

## Gotchas

- `write_file` uses Graph simple upload (`PUT .../content`), which is limited to **4 MB**. Despite the README claiming 250 MB, there is no resumable-upload implementation. Files larger than ~4 MB will fail.
- `list_files` and `search_files` return `nextLink`/`hasMore` for pagination — callers should check these to page through large result sets.

## Not Configured

No test suite, no pytest, no linting (ruff), no type-checking (mypy).
