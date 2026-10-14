"""Enterprise Entra authentication for http mode.

Covers the secretless credential types (certificate, managed identity,
workload identity), the RFC 7523 client-assertion wiring into fastmcp's
AzureProvider, and the fail-closed validation of every new auth setting.
"""

from __future__ import annotations

import base64
import datetime as dt
import time
from types import SimpleNamespace

import httpx2
import jwt
import pytest
from cryptography import x509
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import ec, rsa
from cryptography.x509.oid import NameOID
from fastmcp.server.auth.providers.azure import AzureProvider

from mcp_microsoft import entra
from mcp_microsoft.config import AppConfig, validate_http_config
from mcp_microsoft.entra import (
    CLIENT_ASSERTION_TYPE,
    AssertionOAuth2Client,
    CertificateAssertionSource,
    EntraProvider,
    FederatedTokenFileSource,
    ManagedIdentityAssertionSource,
    fic_audience,
)

_TENANT = "11111111-2222-3333-4444-555555555555"
_SIGNING_KEY = "a" * 64
_TOKEN_URL = f"https://login.microsoftonline.com/{_TENANT}/oauth2/v2.0/token"


def _config(**overrides: object) -> AppConfig:
    base: dict[str, object] = dict(
        transport="http",
        base_url="https://mcp.example.com",
        auth_client_id="cid",
        auth_client_secret="secret",
        auth_tenant_id=_TENANT,
    )
    base.update(overrides)
    return AppConfig(**base)  # type: ignore[arg-type]


def _secretless(kind: str, **overrides: object) -> AppConfig:
    values: dict[str, object] = dict(
        auth_credential_type=kind,
        auth_client_secret="",
        auth_jwt_signing_key=_SIGNING_KEY,
    )
    values.update(overrides)
    return _config(**values)


def _key_and_cert(key=None) -> tuple[object, x509.Certificate, bytes]:
    key = key or rsa.generate_private_key(public_exponent=65537, key_size=2048)
    name = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, "mcp-microsoft-test")])
    now = dt.datetime.now(dt.timezone.utc)
    cert = (
        x509.CertificateBuilder()
        .subject_name(name)
        .issuer_name(name)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now - dt.timedelta(minutes=1))
        .not_valid_after(now + dt.timedelta(days=1))
        .sign(key, hashes.SHA256())
    )
    pem = key.private_bytes(
        serialization.Encoding.PEM,
        serialization.PrivateFormat.PKCS8,
        serialization.NoEncryption(),
    ) + cert.public_bytes(serialization.Encoding.PEM)
    return key, cert, pem


# ---------------------------------------------------------------------------
# validate_http_config
# ---------------------------------------------------------------------------


def test_secret_mode_remains_the_compatible_default() -> None:
    assert validate_http_config(_config()) == []
    problems = validate_http_config(_config(auth_client_secret=""))
    assert any("MCP_AUTH_CLIENT_SECRET is required" in p for p in problems)


@pytest.mark.parametrize(
    ("kind", "extra"),
    [
        ("certificate", {"auth_client_certificate_path": "/run/secrets/app.pem"}),
        ("managed_identity", {}),
        ("workload_identity", {"auth_federated_token_file": "/var/run/token"}),
    ],
)
def test_secretless_modes_validate(kind: str, extra: dict) -> None:
    assert validate_http_config(_secretless(kind, **extra)) == []


def test_secretless_mode_rejects_a_lingering_client_secret() -> None:
    problems = validate_http_config(
        _secretless("managed_identity", auth_client_secret="still-here")
    )
    assert any("must be unset" in p for p in problems)


def test_secretless_mode_requires_signing_key() -> None:
    problems = validate_http_config(
        _secretless("managed_identity", auth_jwt_signing_key="")
    )
    assert any("MCP_AUTH_JWT_SIGNING_KEY is required" in p for p in problems)


def test_certificate_and_workload_modes_require_their_inputs() -> None:
    cert = validate_http_config(_secretless("certificate"))
    assert any("MCP_AUTH_CLIENT_CERTIFICATE_PATH" in p for p in cert)
    workload = validate_http_config(_secretless("workload_identity"))
    assert any("FEDERATED_TOKEN_FILE" in p for p in workload)


def test_unknown_credential_type_is_rejected() -> None:
    problems = validate_http_config(_config(auth_credential_type="password"))
    assert any("MCP_AUTH_CREDENTIAL_TYPE must be one of" in p for p in problems)


def test_short_keys_are_rejected() -> None:
    problems = validate_http_config(
        _config(auth_jwt_signing_key="short", request_state_keys=("tiny",))
    )
    assert any("MCP_AUTH_JWT_SIGNING_KEY must carry" in p for p in problems)
    assert any("MCP_REQUEST_STATE_KEYS entry 1" in p for p in problems)


def test_redis_requires_a_fernet_key_and_key_requires_redis() -> None:
    missing = validate_http_config(_config(auth_redis_url="rediss://cache:6380/0"))
    assert any("MCP_AUTH_STORAGE_ENCRYPTION_KEY must be a Fernet key" in p for p in missing)

    ok = _config(
        auth_redis_url="rediss://cache:6380/0",
        auth_storage_encryption_key=Fernet.generate_key().decode(),
    )
    assert validate_http_config(ok) == []

    orphan = validate_http_config(
        _config(auth_storage_encryption_key=Fernet.generate_key().decode())
    )
    assert any("only applies with MCP_AUTH_REDIS_URL" in p for p in orphan)


@pytest.mark.parametrize(
    "url",
    ["http://localhost:8000", "http://127.0.0.1:8000", "http://[::1]:8000"],
)
def test_plain_http_is_allowed_only_for_loopback(url: str) -> None:
    assert validate_http_config(_config(base_url=url)) == []


def test_plain_http_public_base_url_is_rejected() -> None:
    problems = validate_http_config(_config(base_url="http://mcp.example.com"))
    assert any("must use https://" in p for p in problems)


def test_authority_host_must_be_a_bare_host() -> None:
    problems = validate_http_config(
        _config(auth_authority_host="https://login.microsoftonline.us")
    )
    assert any("MCP_AUTH_AUTHORITY_HOST" in p for p in problems)


def test_from_env_parses_enterprise_settings(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("MCP_AUTH_CREDENTIAL_TYPE", "Workload_Identity")
    monkeypatch.delenv("MCP_AUTH_FEDERATED_TOKEN_FILE", raising=False)
    monkeypatch.setenv("AZURE_FEDERATED_TOKEN_FILE", "/var/run/azure/token")
    monkeypatch.setenv(
        "MCP_AUTH_ALLOWED_REDIRECT_URIS",
        " https://claude.ai/api/mcp/auth_callback , http://localhost:* ,",
    )
    monkeypatch.setenv("MCP_REQUEST_STATE_KEYS", "k1" * 20 + ",k2" + "x" * 40)

    config = AppConfig.from_env()

    assert config.auth_credential_type == "workload_identity"
    assert config.auth_federated_token_file == "/var/run/azure/token"
    assert config.auth_allowed_redirect_uris == (
        "https://claude.ai/api/mcp/auth_callback",
        "http://localhost:*",
    )
    assert len(config.request_state_keys) == 2


# ---------------------------------------------------------------------------
# Assertion sources
# ---------------------------------------------------------------------------


def test_certificate_assertion_has_entra_shape() -> None:
    key, cert, pem = _key_and_cert()
    source = CertificateAssertionSource(client_id="cid", audience=_TOKEN_URL, pem=pem)

    first = source.current()
    second = source.current()

    header = jwt.get_unverified_header(first)
    assert header["alg"] == "PS256"
    expected_thumbprint = (
        base64.urlsafe_b64encode(cert.fingerprint(hashes.SHA256())).rstrip(b"=").decode()
    )
    assert header["x5t#S256"] == expected_thumbprint
    claims = jwt.decode(first, key.public_key(), algorithms=["PS256"], audience=_TOKEN_URL)
    assert claims["iss"] == claims["sub"] == "cid"
    assert claims["exp"] - claims["iat"] == 300
    assert jwt.decode(second, key.public_key(), algorithms=["PS256"], audience=_TOKEN_URL)[
        "jti"
    ] != claims["jti"]


def test_certificate_source_rejects_bad_material(tmp_path) -> None:
    _, _, pem = _key_and_cert()
    other_key, other_cert, _ = _key_and_cert()
    mismatched = pem.split(b"-----BEGIN CERTIFICATE-----")[0] + other_cert.public_bytes(
        serialization.Encoding.PEM
    )
    with pytest.raises(ValueError, match="does not match"):
        CertificateAssertionSource(client_id="c", audience="a", pem=mismatched)

    key_only = pem.split(b"-----BEGIN CERTIFICATE-----")[0]
    with pytest.raises(ValueError, match="must also contain"):
        CertificateAssertionSource(client_id="c", audience="a", pem=key_only)

    _, _, ec_pem = _key_and_cert(ec.generate_private_key(ec.SECP256R1()))
    with pytest.raises(ValueError, match="RSA"):
        CertificateAssertionSource(client_id="c", audience="a", pem=ec_pem)

    with pytest.raises(ValueError, match="Cannot read"):
        CertificateAssertionSource.from_file(
            client_id="c", audience="a", path=str(tmp_path / "missing.pem"), password=None
        )


@pytest.mark.asyncio
async def test_federated_token_file_is_reread_each_time(tmp_path) -> None:
    token_file = tmp_path / "token"
    token_file.write_text("first\n")
    source = FederatedTokenFileSource(str(token_file))
    assert await source.get() == "first"
    token_file.write_text("rotated")
    assert source.current() == "rotated"
    token_file.write_text("")
    with pytest.raises(RuntimeError, match="empty"):
        source.current()


@pytest.mark.asyncio
async def test_managed_identity_source_caches_and_refreshes() -> None:
    calls: list[str] = []

    class FakeCredential:
        def __init__(self) -> None:
            self.closed = False

        async def get_token(self, scope: str):
            calls.append(scope)
            return SimpleNamespace(token=f"mi-{len(calls)}", expires_on=int(time.time()) + 3600)

        async def close(self) -> None:
            self.closed = True

    source = ManagedIdentityAssertionSource(audience="api://AzureADTokenExchange")
    fake = FakeCredential()
    source._credential = fake

    with pytest.raises(RuntimeError, match="await get"):
        source.current()
    assert await source.get() == "mi-1"
    assert await source.get() == "mi-1"
    assert source.current() == "mi-1"
    assert calls == ["api://AzureADTokenExchange/.default"]

    source._expires_on = int(time.time()) + 10  # inside the refresh margin
    assert await source.get() == "mi-2"
    await source.aclose()
    assert fake.closed is True


def test_fic_audience_follows_the_cloud() -> None:
    assert fic_audience("login.microsoftonline.com") == "api://AzureADTokenExchange"
    assert fic_audience("login.microsoftonline.us") == "api://AzureADTokenExchangeUSGov"
    assert fic_audience("login.chinacloudapi.cn") == "api://AzureADTokenExchangeChina"


# ---------------------------------------------------------------------------
# Upstream token client and provider wiring
# ---------------------------------------------------------------------------


class _StaticSource:
    def __init__(self, value: str = "signed-assertion") -> None:
        self.value = value
        self.gets = 0

    async def get(self) -> str:
        self.gets += 1
        return self.value

    def current(self) -> str:
        return self.value

    async def aclose(self) -> None:
        return None


@pytest.mark.asyncio
@pytest.mark.parametrize("grant", ["authorization_code", "refresh_token"])
async def test_assertion_client_sends_rfc7523_credentials(grant: str) -> None:
    seen: dict[str, object] = {}

    def handler(request: httpx2.Request) -> httpx2.Response:
        seen["auth"] = request.headers.get("Authorization")
        seen["form"] = dict(
            pair.split("=", 1) for pair in request.content.decode().split("&")
        )
        return httpx2.Response(200, json={"access_token": "at", "expires_in": 3600})

    client = AssertionOAuth2Client(client_id="cid", source=_StaticSource(), timeout=5)
    await client._client.aclose()
    client._client = httpx2.AsyncClient(transport=httpx2.MockTransport(handler))

    if grant == "authorization_code":
        await client.fetch_token(_TOKEN_URL, code="abc", redirect_uri="https://x/cb")
    else:
        await client.refresh_token(_TOKEN_URL, refresh_token="rt")
    await client.aclose()

    form = seen["form"]
    assert seen["auth"] is None
    assert "client_secret" not in form
    assert form["client_id"] == "cid"
    assert form["client_assertion"] == "signed-assertion"
    assert form["client_assertion_type"] == CLIENT_ASSERTION_TYPE.replace(":", "%3A")
    assert client.client_secret is None


def _entra_provider(source) -> EntraProvider:
    return EntraProvider(
        assertion_source=source,
        client_id="cid",
        client_secret=None,
        tenant_id=_TENANT,
        base_url="https://mcp.example.com",
        required_scopes=["mcp-access"],
        jwt_signing_key=_SIGNING_KEY,
    )


@pytest.mark.asyncio
async def test_provider_uses_assertions_for_upstream_and_obo() -> None:
    source = _StaticSource()
    provider = _entra_provider(source)

    assert isinstance(provider, AzureProvider)  # identity._find_azure_provider
    assert provider.uses_client_assertion is True
    assert isinstance(provider._create_upstream_oauth_client(), AssertionOAuth2Client)

    credential = await provider.get_obo_credential("user-token-a")
    assert await provider.get_obo_credential("user-token-a") is credential
    assert await provider.get_obo_credential("user-token-b") is not credential
    assert source.gets == 3  # refreshed before every exchange
    await provider.close_obo_credentials()


def test_server_builds_secretless_provider_from_config(tmp_path) -> None:
    from mcp_microsoft.server import _build_azure_provider

    _, _, pem = _key_and_cert()
    pem_path = tmp_path / "app.pem"
    pem_path.write_bytes(pem)
    config = _secretless(
        "certificate",
        auth_client_certificate_path=str(pem_path),
        auth_allowed_redirect_uris=("https://claude.ai/api/mcp/auth_callback",),
    )

    provider = _build_azure_provider(config)

    assert isinstance(provider, EntraProvider)
    assert provider.uses_client_assertion is True
    assert provider._allowed_client_redirect_uris == [
        "https://claude.ai/api/mcp/auth_callback"
    ]
    assertion = provider._assertion_source.current()
    assert jwt.decode(assertion, options={"verify_signature": False})["aud"] == _TOKEN_URL


def test_server_keeps_secret_mode_unchanged() -> None:
    from mcp_microsoft.server import _build_azure_provider

    provider = _build_azure_provider(_config())

    assert provider.uses_client_assertion is False
    assert not isinstance(provider._create_upstream_oauth_client(), AssertionOAuth2Client)


def test_server_builds_shared_state_backends() -> None:
    from key_value.aio.wrappers.encryption import FernetEncryptionWrapper
    from mcp.server.request_state import RequestStateSecurity

    from mcp_microsoft.server import (
        _build_oauth_storage,
        _build_request_state_security,
    )

    assert _build_oauth_storage(_config()) is None
    assert _build_request_state_security(_config()) is None

    security = _build_request_state_security(_config(request_state_keys=("k" * 64,)))
    assert isinstance(security, RequestStateSecurity)

    pytest.importorskip("redis", reason="needs the [redis] extra")
    storage = _build_oauth_storage(
        _config(
            auth_redis_url="redis://localhost:6379/0",
            auth_storage_encryption_key=Fernet.generate_key().decode(),
        )
    )
    assert isinstance(storage, FernetEncryptionWrapper)


def test_build_assertion_source_dispatch(tmp_path) -> None:
    assert entra.build_assertion_source(_config()) is None
    mi = entra.build_assertion_source(
        _secretless("managed_identity", auth_authority_host="login.microsoftonline.us")
    )
    assert isinstance(mi, ManagedIdentityAssertionSource)
    assert mi._scope == "api://AzureADTokenExchangeUSGov/.default"
    workload = entra.build_assertion_source(
        _secretless("workload_identity", auth_federated_token_file=str(tmp_path / "t"))
    )
    assert isinstance(workload, FederatedTokenFileSource)
