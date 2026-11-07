"""Sign-in for the remote connector: the OAuth 2.1 flow Claude's connectors use.

Claude registers itself (dynamic client registration) and sends the user to
``/connect`` on this server. The user approves by typing the *connect code*
that ``python connector.py code`` shows on the PC running think-cell. Claude
then gets an access token (1 hour) and a refresh token (90 days, rotated on
use), so it stays connected without asking again.

Tokens are stored only as SHA-256 hashes in a JSON file next to the connector
config, so a restart does not sign Claude out and the file holds no usable
token. Wrong connect codes are limited per sign-in and per 10 minutes overall.
"""
from __future__ import annotations

import hashlib
import hmac
import json
import secrets
import time
from collections import deque
from pathlib import Path

from mcp.server.auth.provider import (
    AccessToken,
    AuthorizationCode,
    AuthorizationParams,
    RefreshToken,
    construct_redirect_uri,
)
from mcp.shared.auth import OAuthClientInformationFull, OAuthToken

ACCESS_TTL = 3600
REFRESH_TTL = 90 * 24 * 3600
CODE_TTL = 300  # authorization code, between /connect and Claude's token call
REQUEST_TTL = 900  # time the user has to type the connect code
TRIES_PER_REQUEST = 5
FAILURES_PER_WINDOW, FAILURE_WINDOW = 10, 600

# No 0/O, 1/I/L: the code is read off one screen and typed into another.
CODE_ALPHABET = "ABCDEFGHJKMNPQRSTUVWXYZ23456789"


def new_connect_code() -> str:
    """A fresh 8-character connect code, e.g. ``K7MP-Q2XD``."""
    raw = "".join(secrets.choice(CODE_ALPHABET) for _ in range(8))
    return f"{raw[:4]}-{raw[4:]}"


def _norm(code: str) -> str:
    return "".join(ch for ch in code.upper() if ch.isalnum())


def _hash(token: str) -> str:
    return hashlib.sha256(token.encode()).hexdigest()


class ConnectorAuth:
    """OAuthAuthorizationServerProvider for one PC: approval by connect code."""

    def __init__(self, state_file: Path, issuer_url: str, connect_code: str,
                 api_key: str | None = None) -> None:
        self.state_file = state_file
        self.issuer_url = issuer_url.rstrip("/")
        self.connect_code = _norm(connect_code)
        # Fixed key for clients that cannot do OAuth (think-cell Assist sends
        # "Authorization: Bearer <key>" from its MCP server settings).
        self.api_key_hash = _hash(api_key) if api_key else None
        self.pending: dict[str, dict] = {}  # sign-in request id -> client, params, expiry, tries
        self.codes: dict[str, AuthorizationCode] = {}
        self.failures: deque[float] = deque()
        state = json.loads(state_file.read_text(encoding="utf-8")) if state_file.is_file() else {}
        self.clients = {k: OAuthClientInformationFull.model_validate(v) for k, v in state.get("clients", {}).items()}
        self.access = {k: AccessToken.model_validate(v) for k, v in state.get("access", {}).items()}
        self.refresh = {k: RefreshToken.model_validate(v) for k, v in state.get("refresh", {}).items()}

    def _save(self) -> None:
        now = time.time()
        self.access = {k: v for k, v in self.access.items() if not v.expires_at or v.expires_at > now}
        self.refresh = {k: v for k, v in self.refresh.items() if not v.expires_at or v.expires_at > now}
        state = {
            "clients": {k: v.model_dump(mode="json") for k, v in self.clients.items()},
            "access": {k: v.model_dump(mode="json") for k, v in self.access.items()},
            "refresh": {k: v.model_dump(mode="json") for k, v in self.refresh.items()},
        }
        self.state_file.parent.mkdir(parents=True, exist_ok=True)
        self.state_file.write_text(json.dumps(state, indent=1), encoding="utf-8")

    # ------------------------------------------------------------ clients
    async def get_client(self, client_id: str) -> OAuthClientInformationFull | None:
        return self.clients.get(client_id)

    async def register_client(self, client_info: OAuthClientInformationFull) -> None:
        self.clients[client_info.client_id] = client_info
        self._save()

    # ------------------------------------------------------------ approval
    async def authorize(self, client: OAuthClientInformationFull, params: AuthorizationParams) -> str:
        request_id = secrets.token_urlsafe(16)
        self.pending[request_id] = {"client_id": client.client_id, "params": params,
                                    "expires": time.time() + REQUEST_TTL, "tries": 0}
        return f"{self.issuer_url}/connect?request={request_id}"

    def client_name(self, request_id: str) -> str | None:
        """Name Claude registered under, for the /connect page; None if the request is gone."""
        request = self.pending.get(request_id)
        if not request or request["expires"] < time.time():
            return None
        client = self.clients.get(request["client_id"])
        return (client.client_name if client else None) or "Claude"

    def approve(self, request_id: str, typed_code: str) -> tuple[str | None, str | None]:
        """Check the typed connect code: ``(redirect URL, None)`` or ``(None, error)``."""
        now = time.time()
        while self.failures and self.failures[0] < now - FAILURE_WINDOW:
            self.failures.popleft()
        if len(self.failures) >= FAILURES_PER_WINDOW:
            return None, "Too many wrong codes. Wait 10 minutes, then connect again from Claude."
        request = self.pending.get(request_id)
        if not request or request["expires"] < now:
            self.pending.pop(request_id, None)
            return None, "This sign-in has expired. Start again from Claude (Connect)."
        if not hmac.compare_digest(_norm(typed_code), self.connect_code):
            self.failures.append(now)
            request["tries"] += 1
            if request["tries"] >= TRIES_PER_REQUEST:
                del self.pending[request_id]
                return None, "Wrong code too many times. Start again from Claude (Connect)."
            return None, "That code is not right. Run  python connector.py code  on the PC to see it."
        del self.pending[request_id]
        params: AuthorizationParams = request["params"]
        code = secrets.token_urlsafe(32)
        self.codes[code] = AuthorizationCode(
            code=code, scopes=params.scopes or [], expires_at=now + CODE_TTL,
            client_id=request["client_id"], code_challenge=params.code_challenge,
            redirect_uri=params.redirect_uri,
            redirect_uri_provided_explicitly=params.redirect_uri_provided_explicitly,
            resource=params.resource,
        )
        return construct_redirect_uri(str(params.redirect_uri), code=code, state=params.state), None

    # ------------------------------------------------------------ tokens
    async def load_authorization_code(self, client: OAuthClientInformationFull,
                                      authorization_code: str) -> AuthorizationCode | None:
        code = self.codes.get(authorization_code)
        if code and code.client_id == client.client_id and code.expires_at > time.time():
            return code
        return None

    async def exchange_authorization_code(self, client: OAuthClientInformationFull,
                                          authorization_code: AuthorizationCode) -> OAuthToken:
        self.codes.pop(authorization_code.code, None)
        return self._issue(client.client_id, authorization_code.scopes, authorization_code.resource)

    async def load_refresh_token(self, client: OAuthClientInformationFull,
                                 refresh_token: str) -> RefreshToken | None:
        token = self.refresh.get(_hash(refresh_token))
        if token and token.client_id == client.client_id and (not token.expires_at or token.expires_at > time.time()):
            return token
        return None

    async def exchange_refresh_token(self, client: OAuthClientInformationFull,
                                     refresh_token: RefreshToken, scopes: list[str]) -> OAuthToken:
        self.refresh.pop(refresh_token.token, None)  # rotate: each refresh token works once
        return self._issue(client.client_id, scopes or refresh_token.scopes, None)

    async def load_access_token(self, token: str) -> AccessToken | None:
        if self.api_key_hash and hmac.compare_digest(_hash(token), self.api_key_hash):
            return AccessToken(token=self.api_key_hash, client_id="api-key", scopes=[])
        access = self.access.get(_hash(token))
        if access and (not access.expires_at or access.expires_at > time.time()):
            return access
        return None

    async def revoke_token(self, token: AccessToken | RefreshToken) -> None:
        self.access.pop(token.token, None)
        self.refresh.pop(token.token, None)
        self._save()

    def _issue(self, client_id: str, scopes: list[str], resource: str | None) -> OAuthToken:
        access, refresh = secrets.token_urlsafe(32), secrets.token_urlsafe(32)
        now = int(time.time())
        # Stored under (and as) their hashes: revoke_token gets the stored object back.
        self.access[_hash(access)] = AccessToken(token=_hash(access), client_id=client_id, scopes=scopes,
                                                 expires_at=now + ACCESS_TTL, resource=resource)
        self.refresh[_hash(refresh)] = RefreshToken(token=_hash(refresh), client_id=client_id, scopes=scopes,
                                                    expires_at=now + REFRESH_TTL)
        self._save()
        return OAuthToken(access_token=access, expires_in=ACCESS_TTL,
                          scope=" ".join(scopes) or None, refresh_token=refresh)
