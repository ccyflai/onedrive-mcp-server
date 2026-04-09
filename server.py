"""
MCP Server for OneDrive for Business — multi-user, HTTP transport.

Provides tools to list, read, and write files in OneDrive for Business
using delegated user identity via the OAuth 2.0 device code flow.

Runs as a Streamable-HTTP MCP server (stateless, JSON responses) so that
multiple MCP clients can connect concurrently over the network.

Multiple users can authenticate concurrently; every file-operation tool
requires a ``user_id`` parameter (the user's email or "username" returned
by the authentication tools) so the server knows whose OneDrive to access.
"""

import asyncio
import base64
import json
import logging
import os
import sys
import time
import uuid
from pathlib import Path

import httpx
import msal

from mcp.server.fastmcp import FastMCP

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
CLIENT_ID = os.environ.get("ONEDRIVE_CLIENT_ID", "")
TENANT_ID = os.environ.get("ONEDRIVE_TENANT_ID", "organizations")
HOST = os.environ.get("ONEDRIVE_MCP_HOST", "0.0.0.0")
PORT = int(os.environ.get("ONEDRIVE_MCP_PORT", "8000"))
SCOPES = ["Files.ReadWrite.All", "User.Read", "offline_access"]

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TOKEN_CACHE_PATH = Path(
    os.environ.get(
        "ONEDRIVE_TOKEN_CACHE",
        Path.home() / ".onedrive-mcp-token-cache.json",
    )
)

logger = logging.getLogger("onedrive-mcp")

# ---------------------------------------------------------------------------
# MSAL helpers — single app, multi-account token cache
# ---------------------------------------------------------------------------
_msal_app: msal.PublicClientApplication | None = None
_token_cache = msal.SerializableTokenCache()

# Pending device-code flows, keyed by a server-generated session_id.
# Each entry stores the MSAL flow dict and a timestamp.
_pending_flows: dict[str, dict] = {}

# How long a pending flow is kept before being garbage-collected (seconds).
_FLOW_TTL = 900  # 15 min (matches the device-code expiry)


def _load_token_cache() -> None:
    """Load the MSAL token cache from disk if it exists."""
    if TOKEN_CACHE_PATH.exists():
        _token_cache.deserialize(TOKEN_CACHE_PATH.read_text())


def _save_token_cache() -> None:
    """Persist the MSAL token cache to disk when it has changed."""
    if _token_cache.has_state_changed:
        TOKEN_CACHE_PATH.write_text(_token_cache.serialize())


def _get_msal_app() -> msal.PublicClientApplication:
    """Return (and lazily create) the shared MSAL public-client application."""
    global _msal_app
    if _msal_app is None:
        if not CLIENT_ID:
            raise RuntimeError(
                "ONEDRIVE_CLIENT_ID environment variable is required. "
                "Register an app in Microsoft Entra (Azure AD) and set "
                "this variable to its Application (client) ID."
            )
        _load_token_cache()
        authority = f"https://login.microsoftonline.com/{TENANT_ID}"
        _msal_app = msal.PublicClientApplication(
            CLIENT_ID, authority=authority, token_cache=_token_cache
        )
    return _msal_app


def _find_account(user_id: str) -> dict | None:
    """Look up a cached MSAL account by username (case-insensitive)."""
    app = _get_msal_app()
    user_lower = user_id.strip().lower()
    for acct in app.get_accounts():
        if acct.get("username", "").lower() == user_lower:
            return acct
    return None


def _gc_pending_flows() -> None:
    """Remove expired pending device-code flows."""
    now = time.monotonic()
    expired = [
        sid
        for sid, entry in _pending_flows.items()
        if now - entry["created_at"] > _FLOW_TTL
    ]
    for sid in expired:
        del _pending_flows[sid]


async def _acquire_token_for_user(user_id: str) -> str:
    """
    Silently acquire an access token for a previously-authenticated user.

    Raises RuntimeError if the user has not authenticated or the refresh
    token has expired.
    """
    app = _get_msal_app()
    account = _find_account(user_id)
    if account is None:
        raise RuntimeError(
            f"No authenticated session found for user '{user_id}'. "
            "Call 'start_auth' then 'complete_auth' to sign this user in."
        )

    result = app.acquire_token_silent(SCOPES, account=account)
    if result and "access_token" in result:
        _save_token_cache()
        return result["access_token"]

    raise RuntimeError(
        f"Token for user '{user_id}' has expired and could not be refreshed. "
        "Call 'start_auth' then 'complete_auth' to re-authenticate."
    )


async def _graph_request(
    user_id: str,
    method: str,
    path: str,
    *,
    params: dict | None = None,
    json_body: dict | None = None,
    content: bytes | None = None,
    content_type: str | None = None,
    follow_redirects: bool = True,
) -> httpx.Response:
    """Make an authenticated Graph API request on behalf of *user_id*."""
    token = await _acquire_token_for_user(user_id)
    headers: dict[str, str] = {"Authorization": f"Bearer {token}"}
    if content_type:
        headers["Content-Type"] = content_type

    url = f"{GRAPH_BASE}{path}" if path.startswith("/") else path

    async with httpx.AsyncClient(follow_redirects=follow_redirects) as client:
        resp = await client.request(
            method,
            url,
            headers=headers,
            params=params,
            json=json_body,
            content=content,
            timeout=60.0,
        )
    resp.raise_for_status()
    return resp


# ---------------------------------------------------------------------------
# MCP server definition — Streamable HTTP, stateless, JSON responses
# ---------------------------------------------------------------------------
mcp = FastMCP(
    "OneDrive for Business",
    stateless_http=True,
    json_response=True,
    host=HOST,
    port=PORT,
    instructions=(
        "This server provides tools to interact with OneDrive for Business "
        "via the Microsoft Graph API. It supports multiple users concurrently "
        "and is served over HTTP (Streamable HTTP transport).\n\n"
        "Authentication workflow:\n"
        "1. Call 'start_auth' — returns a verification URL, user code, and "
        "   a session_id.\n"
        "2. The user opens the URL in a browser and enters the code.\n"
        "3. Call 'complete_auth' with the session_id — polls Microsoft's "
        "   token endpoint and returns the user_id on success.\n"
        "4. Pass that user_id to every file-operation tool.\n\n"
        "Session management:\n"
        "- 'list_users'  — see all authenticated users.\n"
        "- 'remove_user' — sign out a user.\n"
    ),
)


# ------------------------------------------------------------------
# Tool: start_auth
# ------------------------------------------------------------------
@mcp.tool()
async def start_auth() -> str:
    """Start the device-code sign-in flow for a new user.

    Returns immediately with a verification URL, a user_code to enter
    in the browser, and a session_id.  The caller should present the
    URL and code to the user, then call 'complete_auth' with the
    session_id to finish sign-in.

    Returns:
        JSON object: { session_id, verification_uri, user_code, message,
                       expires_in }.
    """
    _gc_pending_flows()

    app = _get_msal_app()
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(
            f"Failed to initiate device code flow: "
            f"{flow.get('error_description', flow)}"
        )

    session_id = uuid.uuid4().hex
    _pending_flows[session_id] = {
        "flow": flow,
        "created_at": time.monotonic(),
    }

    return json.dumps(
        {
            "session_id": session_id,
            "verification_uri": flow.get("verification_uri", ""),
            "user_code": flow["user_code"],
            "expires_in": flow.get("expires_in", 900),
            "message": flow.get("message", ""),
        },
        indent=2,
    )


# ------------------------------------------------------------------
# Tool: complete_auth
# ------------------------------------------------------------------
@mcp.tool()
async def complete_auth(session_id: str) -> str:
    """Complete the device-code sign-in flow.

    Call this after the user has entered the code at the verification URL
    provided by 'start_auth'.

    This tool polls Microsoft's token endpoint (in a background thread)
    until the user completes sign-in or the code expires.

    Args:
        session_id: The session_id returned by 'start_auth'.

    Returns:
        JSON object with user_id, display_name, and status on success.
    """
    entry = _pending_flows.get(session_id)
    if entry is None:
        raise RuntimeError(
            f"No pending authentication flow for session_id '{session_id}'. "
            "It may have expired. Call 'start_auth' to begin a new flow."
        )

    flow = entry["flow"]
    app = _get_msal_app()

    # acquire_token_by_device_flow blocks while polling — run in a thread
    result = await asyncio.to_thread(app.acquire_token_by_device_flow, flow)

    # Clean up regardless of outcome
    _pending_flows.pop(session_id, None)

    if "access_token" not in result:
        raise RuntimeError(
            f"Authentication failed: {result.get('error_description', result)}"
        )

    _save_token_cache()

    claims = result.get("id_token_claims", {})
    username = claims.get("preferred_username", "")
    display_name = claims.get("name", username)

    if not username:
        accounts = app.get_accounts()
        if accounts:
            username = accounts[-1].get("username", "unknown")

    return json.dumps(
        {
            "user_id": username,
            "display_name": display_name,
            "status": "authenticated",
            "message": (
                f"User '{display_name}' ({username}) is now authenticated. "
                f'Pass user_id="{username}" to subsequent tool calls.'
            ),
        },
        indent=2,
    )


# ------------------------------------------------------------------
# Tool: list_users
# ------------------------------------------------------------------
@mcp.tool()
async def list_users() -> str:
    """List all currently authenticated users.

    Returns:
        JSON object with a 'users' array (user_id, display_name) and count.
    """
    app = _get_msal_app()
    accounts = app.get_accounts()

    users = [
        {
            "user_id": acct.get("username", "unknown"),
            "display_name": acct.get("name", acct.get("username", "unknown")),
            "home_account_id": acct.get("home_account_id"),
        }
        for acct in accounts
    ]
    return json.dumps({"users": users, "count": len(users)}, indent=2)


# ------------------------------------------------------------------
# Tool: remove_user
# ------------------------------------------------------------------
@mcp.tool()
async def remove_user(user_id: str) -> str:
    """Sign out a user by removing their cached tokens.

    Args:
        user_id: The user's email / username as returned by 'complete_auth'
                 or 'list_users'.

    Returns:
        Confirmation message.
    """
    app = _get_msal_app()
    account = _find_account(user_id)
    if account is None:
        return json.dumps(
            {"status": "not_found", "message": f"No session found for '{user_id}'."}
        )

    app.remove_account(account)
    _save_token_cache()
    return json.dumps(
        {"status": "removed", "message": f"User '{user_id}' has been signed out."}
    )


# ------------------------------------------------------------------
# Tool: list_files
# ------------------------------------------------------------------
@mcp.tool()
async def list_files(user_id: str, path: str = "/", page_size: int = 50) -> str:
    """List files and folders in a OneDrive for Business directory.

    Args:
        user_id: Email of the authenticated user (from 'complete_auth').
        path: The folder path relative to the drive root.
              Use "/" for the root folder, or paths like "/Documents"
              or "/Documents/Reports".
        page_size: Maximum number of items to return (1-200, default 50).

    Returns:
        JSON object with an 'items' array of file/folder metadata and a
        'count' field.  Each item contains name, id, size,
        lastModifiedDateTime, and type ("file" or "folder").
    """
    page_size = max(1, min(page_size, 200))

    if path in ("", "/"):
        endpoint = "/me/drive/root/children"
    else:
        clean = path.strip("/")
        endpoint = f"/me/drive/root:/{clean}:/children"

    resp = await _graph_request(
        user_id, "GET", endpoint, params={"$top": str(page_size)}
    )
    data = resp.json()
    items = data.get("value", [])

    results = []
    for item in items:
        entry: dict = {
            "name": item.get("name"),
            "id": item.get("id"),
            "size": item.get("size"),
            "lastModifiedDateTime": item.get("lastModifiedDateTime"),
            "webUrl": item.get("webUrl"),
        }
        if "folder" in item:
            entry["type"] = "folder"
            entry["childCount"] = item["folder"].get("childCount")
        else:
            entry["type"] = "file"
            mime = item.get("file", {}).get("mimeType")
            if mime:
                entry["mimeType"] = mime
        results.append(entry)

    next_link = data.get("@odata.nextLink")
    output: dict = {"items": results, "count": len(results)}
    if next_link:
        output["nextLink"] = next_link
        output["hasMore"] = True

    return json.dumps(output, indent=2)


# ------------------------------------------------------------------
# Tool: read_file
# ------------------------------------------------------------------
@mcp.tool()
async def read_file(user_id: str, path: str) -> str:
    """Read (download) a file from OneDrive for Business.

    Args:
        user_id: Email of the authenticated user (from 'complete_auth').
        path: The full file path relative to the drive root,
              e.g. "/Documents/report.txt" or "/myfile.csv".

    Returns:
        For text-based files (< 10 MB): the raw text content.
        For binary files or large text files: a base64-encoded string
        prefixed with "base64:" and preceded by a JSON metadata line.
    """
    clean = path.strip("/")
    if not clean:
        return "Error: Please provide a file path, not a directory."

    meta_resp = await _graph_request(user_id, "GET", f"/me/drive/root:/{clean}")
    meta = meta_resp.json()

    if "folder" in meta:
        return (
            f"Error: '{path}' is a folder, not a file. "
            "Use list_files to browse its contents."
        )

    size = meta.get("size", 0)
    mime = meta.get("file", {}).get("mimeType", "application/octet-stream")
    name = meta.get("name", clean.split("/")[-1])

    download_resp = await _graph_request(
        user_id, "GET", f"/me/drive/root:/{clean}:/content"
    )
    raw = download_resp.content

    text_mimes = (
        "text/",
        "application/json",
        "application/xml",
        "application/javascript",
        "application/x-yaml",
        "application/yaml",
        "application/csv",
        "application/sql",
    )
    is_text = any(mime.startswith(t) for t in text_mimes)

    text_extensions = {
        ".txt",
        ".md",
        ".csv",
        ".json",
        ".xml",
        ".yaml",
        ".yml",
        ".html",
        ".htm",
        ".css",
        ".js",
        ".ts",
        ".py",
        ".sh",
        ".bat",
        ".ps1",
        ".sql",
        ".log",
        ".ini",
        ".cfg",
        ".conf",
        ".toml",
        ".env",
        ".gitignore",
        ".dockerfile",
    }
    ext = Path(name).suffix.lower()
    if ext in text_extensions:
        is_text = True

    TEN_MB = 10 * 1024 * 1024
    if is_text and size < TEN_MB:
        try:
            return raw.decode("utf-8")
        except UnicodeDecodeError:
            return raw.decode("latin-1")

    b64 = base64.b64encode(raw).decode("ascii")
    header = json.dumps(
        {"name": name, "size": size, "mimeType": mime, "encoding": "base64"}
    )
    return f"{header}\nbase64:{b64}"


# ------------------------------------------------------------------
# Tool: write_file
# ------------------------------------------------------------------
@mcp.tool()
async def write_file(
    user_id: str,
    path: str,
    content: str,
    encoding: str = "utf-8",
) -> str:
    """Write (upload) a file to OneDrive for Business.

    Creates a new file or overwrites an existing one.

    Args:
        user_id: Email of the authenticated user (from 'complete_auth').
        path: Destination file path relative to the drive root,
              e.g. "/Documents/notes.txt" or "/data/export.csv".
              Parent folders are created automatically by OneDrive.
        content: The file content as a string.
                 For binary data, pass a base64-encoded string and set
                 encoding="base64".
        encoding: How to interpret 'content'.
                  "utf-8" (default) = plain text.
                  "base64" = base64-encoded binary data.

    Returns:
        JSON metadata of the created/updated file.
    """
    clean = path.strip("/")
    if not clean:
        return "Error: Please provide a file path."

    if encoding == "base64":
        raw = base64.b64decode(content)
    else:
        raw = content.encode("utf-8")

    endpoint = f"/me/drive/root:/{clean}:/content"
    resp = await _graph_request(
        user_id,
        "PUT",
        endpoint,
        content=raw,
        content_type="application/octet-stream",
    )

    result = resp.json()
    return json.dumps(
        {
            "name": result.get("name"),
            "id": result.get("id"),
            "size": result.get("size"),
            "webUrl": result.get("webUrl"),
            "lastModifiedDateTime": result.get("lastModifiedDateTime"),
            "status": "created" if resp.status_code == 201 else "updated",
        },
        indent=2,
    )


# ------------------------------------------------------------------
# Tool: search_files
# ------------------------------------------------------------------
@mcp.tool()
async def search_files(user_id: str, query: str, page_size: int = 25) -> str:
    """Search for files in OneDrive for Business by name or content.

    Args:
        user_id: Email of the authenticated user (from 'complete_auth').
        query: The search query string.  OneDrive searches file names,
               metadata, and — for supported file types — file contents.
        page_size: Maximum number of results (1-200, default 25).

    Returns:
        JSON object with an 'items' array of matching file/folder metadata.
    """
    page_size = max(1, min(page_size, 200))
    endpoint = f"/me/drive/root/search(q='{query}')"
    resp = await _graph_request(
        user_id, "GET", endpoint, params={"$top": str(page_size)}
    )
    data = resp.json()
    items = data.get("value", [])

    results = []
    for item in items:
        entry: dict = {
            "name": item.get("name"),
            "id": item.get("id"),
            "size": item.get("size"),
            "lastModifiedDateTime": item.get("lastModifiedDateTime"),
            "webUrl": item.get("webUrl"),
            "path": (
                item.get("parentReference", {})
                .get("path", "")
                .replace("/drive/root:", "", 1)
                + "/"
                + item.get("name", "")
            ),
        }
        if "folder" in item:
            entry["type"] = "folder"
        else:
            entry["type"] = "file"
            mime = item.get("file", {}).get("mimeType")
            if mime:
                entry["mimeType"] = mime
        results.append(entry)

    output: dict = {"items": results, "count": len(results)}
    next_link = data.get("@odata.nextLink")
    if next_link:
        output["hasMore"] = True

    return json.dumps(output, indent=2)


# ------------------------------------------------------------------
# Tool: get_file_info
# ------------------------------------------------------------------
@mcp.tool()
async def get_file_info(user_id: str, path: str) -> str:
    """Get detailed metadata about a file or folder in OneDrive for Business.

    Args:
        user_id: Email of the authenticated user (from 'complete_auth').
        path: File or folder path relative to the drive root,
              e.g. "/Documents/report.docx" or "/Photos".

    Returns:
        JSON object with detailed metadata: name, size, creation date,
        last modified date, web URL, download URL (for files), etc.
    """
    clean = path.strip("/")
    if not clean:
        endpoint = "/me/drive/root"
    else:
        endpoint = f"/me/drive/root:/{clean}"

    resp = await _graph_request(user_id, "GET", endpoint)
    item = resp.json()

    info: dict = {
        "name": item.get("name"),
        "id": item.get("id"),
        "size": item.get("size"),
        "createdDateTime": item.get("createdDateTime"),
        "lastModifiedDateTime": item.get("lastModifiedDateTime"),
        "webUrl": item.get("webUrl"),
    }

    if "folder" in item:
        info["type"] = "folder"
        info["childCount"] = item["folder"].get("childCount")
    else:
        info["type"] = "file"
        mime = item.get("file", {}).get("mimeType")
        if mime:
            info["mimeType"] = mime
        download_url = item.get("@microsoft.graph.downloadUrl")
        if download_url:
            info["downloadUrl"] = download_url

    created_by = item.get("createdBy", {}).get("user", {})
    if created_by:
        info["createdBy"] = created_by.get("displayName")

    modified_by = item.get("lastModifiedBy", {}).get("user", {})
    if modified_by:
        info["lastModifiedBy"] = modified_by.get("displayName")

    parent = item.get("parentReference", {})
    if parent.get("path"):
        info["parentPath"] = parent["path"].replace("/drive/root:", "", 1) or "/"

    return json.dumps(info, indent=2)


# ---------------------------------------------------------------------------
# Entrypoint
# ---------------------------------------------------------------------------
def main() -> None:
    """Run the MCP server with Streamable HTTP transport."""
    mcp.run(transport="streamable-http")


if __name__ == "__main__":
    main()
