"""
Secretless Microsoft Entra credentials for the http (multi-user) transport.

In http mode this server is a confidential client: it redeems authorization
codes and refresh tokens at Entra's token endpoint (fastmcp's OAuth proxy) and
exchanges each caller's token for a Microsoft Graph token (On-Behalf-Of). By
default fastmcp authenticates both with a client secret. Microsoft's guidance
for production workloads is to avoid client secrets in favour of certificate
credentials or federated identity credentials (managed identity / workload
identity), which this module provides as RFC 7523 client assertions:

* ``certificate`` -- a JWT signed with the app's certificate private key
  (``PS256``, ``x5t#S256`` header), minted per request.
* ``managed_identity`` -- an Azure managed identity token for the Entra token
  exchange audience, used as a federated identity credential.
* ``workload_identity`` -- a projected federated token file (for example AKS
  workload identity), re-read on every use because the platform rotates it.

:class:`EntraProvider` wires an assertion source into fastmcp's
``AzureProvider`` through its two documented extension points:
``_create_upstream_oauth_client`` (code/refresh redemption) and
``get_obo_credential`` (On-Behalf-Of).

References:
- https://learn.microsoft.com/entra/identity-platform/certificate-credentials
- https://learn.microsoft.com/entra/workload-id/workload-identity-federation-config-app-trust-managed-identity
- https://learn.microsoft.com/entra/identity-platform/v2-oauth2-on-behalf-of-flow
"""

from __future__ import annotations

import base64
import hashlib
import time
import uuid
from collections import OrderedDict
from pathlib import Path
from typing import TYPE_CHECKING, Any, Protocol

from fastmcp.server.auth.oauth_proxy.upstream import AsyncOAuth2Client
from fastmcp.server.auth.providers.azure import AzureProvider

if TYPE_CHECKING:
    from azure.identity.aio import OnBehalfOfCredential

    from mcp_microsoft.config import AppConfig

# RFC 7523 client assertion type for JWT bearer client authentication.
CLIENT_ASSERTION_TYPE = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"

# Lifetime of a certificate-signed assertion. Entra accepts up to 10 minutes;
# keep it short because a new one is minted for every token request.
_CERT_ASSERTION_LIFETIME_S = 300

# Refresh a cached managed-identity assertion this long before it expires.
_MI_REFRESH_MARGIN_S = 300

# Federated identity credential audiences per Entra authority host.
_FIC_AUDIENCES = {
    "login.microsoftonline.us": "api://AzureADTokenExchangeUSGov",
    "login.chinacloudapi.cn": "api://AzureADTokenExchangeChina",
}
_FIC_DEFAULT_AUDIENCE = "api://AzureADTokenExchange"

# Upper bound on cached per-user OBO credentials (matches fastmcp's default).
_OBO_CACHE_MAX = 128


class AssertionSource(Protocol):
    """Supplies client assertions proving this app's identity to Entra."""

    async def get(self) -> str:
        """Return a valid assertion, refreshing any cached one if needed."""
        ...

    def current(self) -> str:
        """Return a valid assertion without awaiting (for sync callbacks)."""
        ...

    async def aclose(self) -> None:
        """Release any underlying credential."""
        ...


def fic_audience(authority_host: str) -> str:
    """Return the federated-identity-credential audience for *authority_host*."""
    return _FIC_AUDIENCES.get(authority_host.lower(), _FIC_DEFAULT_AUDIENCE)


def token_endpoint(authority_host: str, tenant_id: str) -> str:
    return f"https://{authority_host}/{tenant_id}/oauth2/v2.0/token"


def _b64url(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode("ascii")


class CertificateAssertionSource:
    """Mints certificate-signed client assertions (Entra certificate credentials).

    The PEM file must contain the private key and the matching certificate that
    was uploaded to the app registration. Nothing is cached: every call mints a
    fresh, short-lived assertion with a unique ``jti``.
    """

    def __init__(
        self,
        *,
        client_id: str,
        audience: str,
        pem: bytes,
        password: str | None = None,
    ) -> None:
        from cryptography import x509
        from cryptography.hazmat.primitives import hashes, serialization
        from cryptography.hazmat.primitives.asymmetric import rsa

        try:
            key = serialization.load_pem_private_key(
                pem, password=password.encode() if password else None
            )
        except (TypeError, ValueError) as exc:
            raise ValueError(
                "MCP_AUTH_CLIENT_CERTIFICATE_PATH does not hold a readable PEM "
                "private key (wrong file, or wrong/missing "
                "MCP_AUTH_CLIENT_CERTIFICATE_PASSWORD)."
            ) from exc
        if not isinstance(key, rsa.RSAPrivateKey):
            raise ValueError(
                "Entra certificate credentials must use an RSA key; the PEM "
                "private key in MCP_AUTH_CLIENT_CERTIFICATE_PATH is not RSA."
            )
        try:
            certificate = x509.load_pem_x509_certificate(pem)
        except ValueError as exc:
            raise ValueError(
                "MCP_AUTH_CLIENT_CERTIFICATE_PATH must also contain the PEM "
                "certificate uploaded to the app registration."
            ) from exc
        if certificate.public_key().public_numbers() != key.public_key().public_numbers():
            raise ValueError(
                "The certificate in MCP_AUTH_CLIENT_CERTIFICATE_PATH does not "
                "match its private key."
            )

        self._client_id = client_id
        self._audience = audience
        self._key = key
        self._thumbprint = _b64url(certificate.fingerprint(hashes.SHA256()))

    @classmethod
    def from_file(
        cls, *, client_id: str, audience: str, path: str, password: str | None
    ) -> CertificateAssertionSource:
        try:
            pem = Path(path).read_bytes()
        except OSError as exc:
            raise ValueError(
                f"Cannot read MCP_AUTH_CLIENT_CERTIFICATE_PATH ({exc.strerror})."
            ) from exc
        return cls(client_id=client_id, audience=audience, pem=pem, password=password)

    def current(self) -> str:
        import jwt

        now = int(time.time())
        claims = {
            "aud": self._audience,
            "iss": self._client_id,
            "sub": self._client_id,
            "jti": str(uuid.uuid4()),
            "nbf": now,
            "iat": now,
            "exp": now + _CERT_ASSERTION_LIFETIME_S,
        }
        return jwt.encode(
            claims,
            self._key,
            algorithm="PS256",
            headers={"x5t#S256": self._thumbprint, "typ": "JWT"},
        )

    async def get(self) -> str:
        return self.current()

    async def aclose(self) -> None:
        return None


class FederatedTokenFileSource:
    """Reads a platform-projected federated token (workload identity).

    The platform rotates the file before the token expires, so it is re-read on
    every use rather than cached.
    """

    def __init__(self, path: str) -> None:
        self._path = Path(path)

    def current(self) -> str:
        try:
            token = self._path.read_text(encoding="utf-8").strip()
        except OSError as exc:
            raise RuntimeError(
                f"Cannot read the federated token file ({exc.strerror}); check "
                "MCP_AUTH_FEDERATED_TOKEN_FILE / AZURE_FEDERATED_TOKEN_FILE."
            ) from exc
        if not token:
            raise RuntimeError("The federated token file is empty.")
        return token

    async def get(self) -> str:
        return self.current()

    async def aclose(self) -> None:
        return None


class ManagedIdentityAssertionSource:
    """Uses an Azure managed identity token as a federated identity credential.

    The managed identity must be registered as a federated identity credential
    on the app registration (issuer ``https://login.microsoftonline.com/<tenant>/v2.0``,
    subject = the identity's object ID). The token is cached and refreshed
    asynchronously; :meth:`current` only returns the cached value so the sync
    callback azure-identity invokes never blocks the event loop on the network.
    """

    def __init__(self, *, audience: str, client_id: str | None = None) -> None:
        self._scope = f"{audience}/.default"
        self._client_id = client_id or None
        self._credential: Any = None
        self._token: str | None = None
        self._expires_on = 0

    async def get(self) -> str:
        if self._token is None or self._expires_on - time.time() < _MI_REFRESH_MARGIN_S:
            if self._credential is None:
                from azure.identity.aio import ManagedIdentityCredential

                self._credential = ManagedIdentityCredential(client_id=self._client_id)
            access = await self._credential.get_token(self._scope)
            self._token = access.token
            self._expires_on = access.expires_on
        return self._token

    def current(self) -> str:
        if self._token is None or self._expires_on <= time.time():
            raise RuntimeError(
                "Managed identity assertion is not available; await get() first."
            )
        return self._token

    async def aclose(self) -> None:
        if self._credential is not None:
            await self._credential.close()
            self._credential = None


def build_assertion_source(config: AppConfig) -> AssertionSource | None:
    """Return the assertion source for ``config.auth_credential_type``.

    ``None`` means the client-secret flow (fastmcp's default) is in use.
    Construction reads the certificate file, so configuration errors surface at
    startup instead of on the first sign-in.
    """
    kind = config.auth_credential_type
    if kind == "secret":
        return None
    if kind == "certificate":
        return CertificateAssertionSource.from_file(
            client_id=config.auth_client_id,
            audience=token_endpoint(config.auth_authority_host, config.auth_tenant_id),
            path=config.auth_client_certificate_path,
            password=config.auth_client_certificate_password or None,
        )
    if kind == "managed_identity":
        return ManagedIdentityAssertionSource(
            audience=fic_audience(config.auth_authority_host),
            client_id=config.auth_managed_identity_client_id or None,
        )
    if kind == "workload_identity":
        return FederatedTokenFileSource(config.auth_federated_token_file)
    raise ValueError(f"Unsupported MCP_AUTH_CREDENTIAL_TYPE {kind!r}")


class AssertionOAuth2Client(AsyncOAuth2Client):
    """fastmcp's upstream token client, authenticating with a client assertion.

    Adds ``client_assertion_type``/``client_assertion`` (RFC 7523 section 2.2)
    to every token request instead of a client secret. ``client_secret`` stays
    ``None``, so fastmcp's best-effort upstream revocation identifies the app
    by ``client_id`` only, as for a public client.
    """

    def __init__(
        self, *, client_id: str, source: AssertionSource, timeout: float | None
    ) -> None:
        super().__init__(
            client_id=client_id,
            client_secret=None,
            token_endpoint_auth_method="none",
            timeout=timeout,
        )
        self._source = source

    async def _assertion_params(self) -> dict[str, str]:
        return {
            "client_assertion_type": CLIENT_ASSERTION_TYPE,
            "client_assertion": await self._source.get(),
        }

    async def fetch_token(
        self, url: str, *, grant_type: str = "authorization_code", **params: Any
    ) -> dict[str, Any]:
        params.update(await self._assertion_params())
        return await super().fetch_token(url, grant_type=grant_type, **params)

    async def refresh_token(
        self, url: str, *, refresh_token: str | None = None, **params: Any
    ) -> dict[str, Any]:
        params.update(await self._assertion_params())
        return await super().refresh_token(url, refresh_token=refresh_token, **params)


class EntraProvider(AzureProvider):
    """``AzureProvider`` that can authenticate to Entra without a client secret.

    With ``assertion_source=None`` it behaves exactly like ``AzureProvider``.
    """

    def __init__(
        self, *, assertion_source: AssertionSource | None = None, **kwargs: Any
    ) -> None:
        super().__init__(**kwargs)
        self._assertion_source = assertion_source
        self._entra_client_id: str = kwargs["client_id"]
        self._entra_tenant_id: str = kwargs["tenant_id"]
        self._entra_authority: str = kwargs.get(
            "base_authority", "login.microsoftonline.com"
        )
        self._assertion_obo: OrderedDict[str, OnBehalfOfCredential] = OrderedDict()

    @property
    def uses_client_assertion(self) -> bool:
        return self._assertion_source is not None

    def _create_upstream_oauth_client(self) -> AsyncOAuth2Client:
        if self._assertion_source is None:
            return super()._create_upstream_oauth_client()
        from fastmcp.server.auth.oauth_proxy.models import HTTP_TIMEOUT_SECONDS

        return AssertionOAuth2Client(
            client_id=self._entra_client_id,
            source=self._assertion_source,
            timeout=HTTP_TIMEOUT_SECONDS,
        )

    async def get_obo_credential(self, user_assertion: str) -> OnBehalfOfCredential:
        if self._assertion_source is None:
            return await super().get_obo_credential(user_assertion)

        from azure.identity.aio import OnBehalfOfCredential

        # Make sure a fresh assertion is cached before azure-identity calls the
        # sync ``current`` callback during the exchange.
        await self._assertion_source.get()

        key = hashlib.sha256(user_assertion.encode()).hexdigest()
        cached = self._assertion_obo.get(key)
        if cached is not None:
            self._assertion_obo.move_to_end(key)
            return cached

        credential = OnBehalfOfCredential(
            tenant_id=self._entra_tenant_id,
            client_id=self._entra_client_id,
            user_assertion=user_assertion,
            authority=f"https://{self._entra_authority}",
            client_assertion_func=self._assertion_source.current,
        )
        self._assertion_obo[key] = credential
        while len(self._assertion_obo) > _OBO_CACHE_MAX:
            _, evicted = self._assertion_obo.popitem(last=False)
            await evicted.close()
        return credential

    async def close_obo_credentials(self) -> None:
        await super().close_obo_credentials()
        credentials = list(self._assertion_obo.values())
        self._assertion_obo.clear()
        for credential in credentials:
            try:
                await credential.close()
            except Exception:  # best-effort shutdown; never mask the original
                pass
        if self._assertion_source is not None:
            await self._assertion_source.aclose()
