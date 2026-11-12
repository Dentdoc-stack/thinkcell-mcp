"""MCP 2026-07-28 ("modern", stateless) on the connector's HTTP endpoint.

The MCP Python SDK 1.x speaks the legacy protocol (``initialize`` handshake,
up to 2025-11-25). The 2026-07-28 revision drops the handshake: every request
carries its protocol version in ``params._meta`` and in the
``MCP-Protocol-Version`` header, and servers must answer ``server/discover``.
The spec lets one endpoint serve both eras ("dual-era"), so this ASGI wrapper
takes requests that carry modern ``_meta`` and passes everything else to the
SDK unchanged. Tool calls still go through the SDK's own handlers, so the input
and output schema checks and error results are the same in both eras.

Supported here: ``server/discover``, ``tools/list`` and ``tools/call``. This
server has no prompts, resources or subscriptions. Other methods get 404 and
-32601, as the Streamable HTTP binding says.
"""
from __future__ import annotations

import base64
import json
from typing import Any

from mcp import types
from mcp.shared.version import SUPPORTED_PROTOCOL_VERSIONS

MODERN_VERSIONS = ["2026-07-28"]
SUPPORTED_VERSIONS = [*MODERN_VERSIONS, *reversed(SUPPORTED_PROTOCOL_VERSIONS)]
META_VERSION = "io.modelcontextprotocol/protocolVersion"
META_SERVER_INFO = "io.modelcontextprotocol/serverInfo"
LIST_TTL_MS = 3_600_000  # the tool list only changes when the server is updated
MAX_BODY = 8_000_000  # bytes per request; chart data is far smaller

HEADER_MISMATCH, UNSUPPORTED_VERSION, METHOD_NOT_FOUND, INVALID_PARAMS = -32020, -32022, -32601, -32602


def _header_value(raw: str | None) -> str | None:
    """Undo the spec's ``=?base64?...?=`` encoding of Mcp-Name / Mcp-Param-* values."""
    if raw and raw.startswith("=?base64?") and raw.endswith("?="):
        try:
            return base64.b64decode(raw[9:-2]).decode()
        except ValueError:
            return None
    return raw


def modern_era(app, server_mcp, verifier, resource_metadata_url: str, server_info: dict[str, str],
               instructions: str | None = None):
    """Wrap the SDK app so modern requests to ``/mcp`` are answered here."""
    from starlette.responses import JSONResponse, Response

    low = server_mcp._mcp_server  # shortcut: the SDK's own request handlers, reused for both eras

    def rpc_error(status, id_, code, message, data=None):
        error = {"code": code, "message": message, **({"data": data} if data is not None else {})}
        return JSONResponse({"jsonrpc": "2.0", "id": id_, "error": error}, status_code=status)

    def complete(id_, result: dict[str, Any]):
        result = {"resultType": "complete", **result}
        result["_meta"] = {**result.get("_meta", {}), META_SERVER_INFO: server_info}
        return JSONResponse({"jsonrpc": "2.0", "id": id_, "result": result})

    async def authorized(headers) -> bool:
        auth = headers.get("authorization", "")
        return auth.lower().startswith("bearer ") and await verifier.verify_token(auth[7:].strip()) is not None

    async def handle(msg: dict, headers: dict[str, str]):
        id_, method = msg.get("id"), msg.get("method")
        params = msg.get("params") or {}
        version = params["_meta"][META_VERSION]
        if not await authorized(headers):
            return Response(status_code=401, headers={"WWW-Authenticate": (
                f'Bearer error="invalid_token", error_description="Authentication required", '
                f'resource_metadata="{resource_metadata_url}"')})
        if headers.get("mcp-protocol-version") != version:
            return rpc_error(400, id_, HEADER_MISMATCH,
                             f"Header mismatch: MCP-Protocol-Version {headers.get('mcp-protocol-version')!r} "
                             f"does not match _meta protocolVersion {version!r}")
        if version not in MODERN_VERSIONS:
            return rpc_error(400, id_, UNSUPPORTED_VERSION, "Unsupported protocol version",
                             {"supported": SUPPORTED_VERSIONS, "requested": version})
        if headers.get("mcp-method") != method:
            return rpc_error(400, id_, HEADER_MISMATCH,
                             f"Header mismatch: Mcp-Method {headers.get('mcp-method')!r} does not match {method!r}")
        if method == "tools/call" and _header_value(headers.get("mcp-name")) != params.get("name"):
            return rpc_error(400, id_, HEADER_MISMATCH,
                             f"Header mismatch: Mcp-Name {headers.get('mcp-name')!r} does not match "
                             f"{params.get('name')!r}")
        if id_ is None:  # a notification: nothing to answer
            return Response(status_code=202)

        if method == "server/discover":
            return complete(id_, {"supportedVersions": SUPPORTED_VERSIONS, "capabilities": {"tools": {}},
                                  **({"instructions": instructions} if instructions else {}),
                                  "ttlMs": LIST_TTL_MS, "cacheScope": "private"})
        if method == "tools/list":
            result = await low.request_handlers[types.ListToolsRequest](types.ListToolsRequest(method="tools/list"))
            tools = sorted(result.root.tools, key=lambda t: t.name)  # deterministic order (spec SHOULD)
            return complete(id_, {"tools": [t.model_dump(by_alias=True, exclude_none=True) for t in tools],
                                  "ttlMs": LIST_TTL_MS, "cacheScope": "private"})
        if method == "tools/call":
            if not isinstance(params.get("name"), str):
                return rpc_error(400, id_, INVALID_PARAMS, "tools/call needs a tool name")
            request = types.CallToolRequest(method="tools/call", params=types.CallToolRequestParams(
                name=params["name"], arguments=params.get("arguments") or {}))
            if params["name"] not in {t.name for t in await server_mcp.list_tools()}:
                return rpc_error(400, id_, INVALID_PARAMS, f"Unknown tool: {params['name']}")
            result = await low.request_handlers[types.CallToolRequest](request)
            return complete(id_, result.root.model_dump(by_alias=True, exclude_none=True))
        return rpc_error(404, id_, METHOD_NOT_FOUND, f"Method not found: {method}")

    async def wrapped(scope, receive, send):
        if scope["type"] != "http" or scope["method"] != "POST" or scope["path"].rstrip("/") != "/mcp":
            return await app(scope, receive, send)
        body, more = b"", True
        while more:
            message = await receive()
            body += message.get("body", b"")
            more = message.get("more_body", False)
            if len(body) > MAX_BODY:  # read before sign-in, so anyone could fill memory
                return await Response(status_code=413)(scope, receive, send)
        try:
            msg = json.loads(body)
            modern = isinstance(msg, dict) and META_VERSION in ((msg.get("params") or {}).get("_meta") or {})
        except ValueError:
            modern = False
        if not modern:  # legacy (initialize-era) traffic: hand the body back to the SDK unchanged
            sent = False

            async def replay():
                nonlocal sent
                if sent:
                    return await receive()
                sent = True
                return {"type": "http.request", "body": body, "more_body": False}
            return await app(scope, replay, send)
        headers = {k.decode().lower(): v.decode() for k, v in scope.get("headers", [])}
        response = await handle(msg, headers)
        await response(scope, receive, send)

    return wrapped
