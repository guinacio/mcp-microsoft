# Multi-user Streamable HTTP server — implementation plan

**Date:** 2026-07-12 · **Target version:** 0.8.0 · **Status:** approved for implementation

## Goal

Add a production-ready **multi-user remote server mode** over MCP **Streamable HTTP** (spec 2025-11-25),
where each connecting user authenticates with their own Microsoft Entra ID account and every Graph call
runs under that user's delegated identity — while keeping the existing single-user stdio mode (MCPB /
Claude Desktop) fully backward compatible.

## Architecture decision

Two mutually exclusive server modes, selected at startup:

| | `stdio` (default, unchanged) | `http` (new) |
|---|---|---|
| Transport | stdio | Streamable HTTP (`/mcp/`) |
| Identity | `profile` string → ProfileManager singleton → MSAL `PublicClientApplication` (interactive/device-code), encrypted disk cache | Per-request bearer token → FastMCP `AzureProvider` (OAuth proxy) → **On-Behalf-Of** exchange for a Graph token |
| Azure app type | Public client (Mobile & desktop) | **Confidential client** (Web, client secret) |
| Multi-account | Multiple named profiles | Any number of concurrent users, one Entra identity each |
| Profile tools | Registered | **Not registered** (server-operator concern, dangerous remotely) |
| `profile` param on tools | Honored | **Inert** (ignored; identity always from token) |

### Why AzureProvider + OBO (not RemoteAuthProvider/JWT-only)

- Entra ID does not support Dynamic Client Registration; FastMCP's `AzureProvider` implements the
  OAuth-proxy pattern that bridges this for MCP clients (Claude, VS Code, etc.), including
  `/.well-known/oauth-protected-resource` (RFC 9728) required by MCP 2025-11-25 authorization.
- `AzureProvider.get_obo_credential(user_assertion)` (fastmcp ≥3.2, mature in 3.4.x) performs the
  OBO exchange via `azure.identity.aio.OnBehalfOfCredential`, with per-user LRU credential caching.
- The `_EntraOBOToken` dependency in fastmcp 3.4.4 proves the exchange is possible with public APIs
  only: `get_access_token()` → `get_server()` → `_find_azure_provider(server.auth)` →
  `await provider.get_obo_credential(access_token.token)` → `await credential.get_token(*scopes)`.
  We replicate this **inside the graph layer** instead of adding a dependency parameter to 89 tools.

### The single seam

`GraphClient._get_headers()` (graph.py:91-95) is the only place identity enters an outbound Graph
request, re-evaluated on every retry attempt. All 89 Graph-backed tools funnel through
`get_graph(params.profile)`. Strategy:

1. Introduce a **token-provider abstraction** consumed by `GraphClient`.
2. stdio mode → `ProfileTokenProvider` (wraps today's ProfileManager path; behavior identical).
3. http mode → `OboTokenProvider` (reads ambient request auth context; `profile` arg ignored).
4. Tool bodies, input models, and tests: **no changes**.

Because the OBO exchange is async, `_get_headers()` becomes `async`; the stdio MSAL path runs the
blocking MSAL call via `anyio.to_thread.run_sync` (bonus: stops blocking the event loop).

## Graph token scope strategy (http mode)

Request `https://graph.microsoft.com/.default` in the OBO exchange — resolves to all delegated Graph
permissions configured on the app registration. Service enablement (Teams/SharePoint) stays governed
by feature-flag env vars; in http mode flags must be **explicit** (no corporate-account fallback —
that heuristic depends on a configured profile which http mode doesn't use).

`AzureProvider.additional_authorize_scopes` = the Graph scopes actually needed (from
`build_default_scopes()` semantics) + `offline_access`, expressed as full
`https://graph.microsoft.com/<Scope>` URIs.

**Audience:** http mode targets a single work/school tenant (`tenant_id` = the concrete tenant
GUID; pseudo-tenants such as `organizations`/`common`/`consumers` and verified domains are rejected
at startup — fastmcp's `AzureProvider` pins the accepted token issuer to a literal URL built from
this value, so only the concrete GUID validates real tokens). Personal Microsoft accounts remain
served by stdio mode (OBO + custom API scopes are not reliably supported for consumer accounts).

## Azure App Registration (http mode) — required setup

New or updated registration (documented in docs/azure-setup.md):
1. Platform **Web**, redirect URI `{BASE_URL}/auth/callback`.
2. **Expose an API**: Application ID URI `api://{client_id}`; add scope `mcp-access`.
3. Manifest: `"requestedAccessTokenVersion": 2`.
4. **Client secret** (confidential client).
5. Delegated Graph permissions: existing DEFAULT_SCOPES (+ Teams/SharePoint scopes if enabled).
6. Admin consent as required by tenant policy.

## Configuration additions (AppConfig, env-driven, frozen dataclass)

| Env var | Purpose | Default |
|---|---|---|
| `MCP_TRANSPORT` | `stdio` \| `http` | `stdio` |
| `MCP_HTTP_HOST` | bind host | `127.0.0.1` |
| `MCP_HTTP_PORT` | bind port | `8000` |
| `MCP_BASE_URL` | public base URL (behind proxy: the external HTTPS URL) | required in http mode |
| `MCP_AUTH_CLIENT_ID` | confidential app client id | required in http mode |
| `MCP_AUTH_CLIENT_SECRET` | client secret | required in http mode |
| `MCP_AUTH_TENANT_ID` | tenant GUID (concrete; pseudo-tenants/domains rejected) | required in http mode |
| `MCP_AUTH_REQUIRED_SCOPE` | custom API scope name | `mcp-access` |
| `MCP_HTTP_STATELESS` | stateless streamable HTTP (horizontal scaling) | `false` |
| `MCP_RATE_LIMIT_RPS` | per-user rate limit | `10` |

Fail-fast validation: http mode with missing/incomplete auth config aborts startup with a clear error.
stdio mode ignores all `MCP_AUTH_*` / `MCP_HTTP_*` vars.

## Phases

### Phase 0 — branch + dependency upgrade
- Branch `feature/http-multi-user` off master.
- `pyproject.toml`: `fastmcp[azure]>=3.4.4` (metapackage → `fastmcp-slim[azure]` → `azure-identity`);
  keep floor honest. `uv lock` + `uv sync`; note installed venv had 3.2.0 vs lock 3.2.0/pip-visible 3.0.2 —
  re-sync cleanly.
- Run full test suite; fix any 3.2→3.4 API breakage (expected: none or minor).

### Phase 1 — token-provider abstraction (no behavior change)
- New `src/mcp_microsoft/identity.py`:
  - `TokenProvider` protocol: `async def get_token(self) -> str`.
  - `ProfileTokenProvider(profile: str | None)` — wraps `ProfileManager.get_token` via
    `anyio.to_thread.run_sync`.
- `graph.py`: `GraphClient(token_provider)` (keep `profile` ctor arg for compat, mapped to
  `ProfileTokenProvider`); `_get_headers()` → `async`; `_send_with_retry` awaits it.
- `ProfileManager.get_graph` unchanged externally.
- Existing tests must pass untouched (they monkeypatch module-local `get_graph` — seam preserved).
- New unit tests for identity.py.

### Phase 2 — config + server wiring for dual transport
- `config.py`: new fields per table above + validation helper `validate_http_config()`.
- `server.py`:
  - `create_mcp_server(config)` gains mode awareness: in http mode, construct `AzureProvider`
    (client id/secret/tenant, `base_url`, `required_scopes=[required_scope]`,
    `additional_authorize_scopes=[Graph scopes + offline_access]`) and pass `auth=` to `FastMCP`.
  - **Do not register** `tools/profiles.py` tools in http mode; keep `services.py`.
  - Feature flags: http mode resolves flags from env only (never profile fallback).
- `main()`: transport dispatch — stdio → `.run()`; http → `.run(transport="http", host, port)`
  (+ stateless flag). CLI passthrough args optional.
- Tests: config validation, tool-registration matrix per mode (extend
  test_security_hardening.py pattern).

### Phase 3 — per-user OBO identity path
- `identity.py`: `OboTokenProvider` — replicates `_EntraOBOToken` logic:
  `get_access_token()` (raise clean auth error if absent) → `get_server()` → find AzureProvider →
  `get_obo_credential(token)` → `credential.get_token("https://graph.microsoft.com/.default")`.
- `get_graph(profile)` becomes mode-aware: http mode returns the shared OBO-backed `GraphClient`
  (headers computed per request from ambient context → naturally per-user); `profile` ignored.
- `tools/sharepoint.py:_get_sharepoint_graph`: http-mode branch — derive tenant type from token
  claims (`tid`) instead of `resolve_profile()`.
- Per-user audit fields: log `oid`/`preferred_username` claim (never the token) at tool-call level.
- Tests: fake auth context / fake provider; assert per-user header derivation and profile-arg inertness.

### Phase 4 — production hardening
- Middleware (http mode only): bounded per-user rate-limiting middleware (see As-built deviations);
  structured logging middleware
  (tool name, user oid, duration, outcome); error masking (`mask_error_details=True`).
- **Local-disk tool audit**: any tool that reads/writes server-local paths (onedrive upload/download
  `save_path`-style args, attachment save) must be gated or content-returning in http mode —
  server disk is not the user's disk. Inventory & gate.
- `@mcp.custom_route("/health")` → 200 JSON (no auth) for load balancers.
- Deletion kill-switch, Teams/SharePoint gating: verified functional in http mode.
- Security notes: TLS termination via reverse proxy, `MCP_BASE_URL` must match public URL,
  bind guidance, secrets handling (env only; Key Vault pointer), single-worker constraint
  (in-memory OAuth-proxy client storage + OBO cache) with scaling escape hatch (`client_storage`).

### Phase 5 — tests, Docker, docs, release prep
- Integration tests: in-memory FastMCP client against http-mode server construction; auth-context
  fakes; regression run of full suite.
- `Dockerfile` (uv-based, non-root) + `docker-compose.yml` example (env-driven) + `.env.example`.
- README: new "Remote server (multi-user)" section; docs/azure-setup.md: confidential-client
  walkthrough (Web platform, Expose an API, secret, v2 tokens); CHANGELOG; bump 0.8.0.
- manifest.json untouched (MCPB remains stdio).

### Phase 6 — orchestrator review + verification
- Full-diff review by supervisor; run suite; smoke-run stdio and http modes; code-review pass;
  final report.

## Risks / notes

- **fastmcp 3.2→3.4 upgrade** may carry subtle breaking changes (3.x line is moving fast) — Phase 0
  runs the suite first, isolated.
- `AccessToken.token` in the OBO call is the user assertion accepted by Entra — this mirrors
  fastmcp's own `_EntraOBOToken`; if the proxy hands us a FastMCP-issued JWT instead, fall back to
  requesting Azure tokens directly (AzureProvider validates Azure JWTs; verify at Phase 3 with a
  real tenant if possible, else against fastmcp's own tests).
- OBO does not work for personal MSA in the general case — explicitly out of scope for http mode.
- Multi-worker deployments need external `client_storage` — out of scope; documented.
- MCPB/stdio users: zero behavior change; `authenticate_ms_profile` and disk caches untouched.

## As-built deviations

Minor points where the shipped implementation differs from the plan body above;
recorded here rather than by rewriting the historical text.

- **TokenProvider method name.** The `TokenProvider` protocol method is
  `get_access_token()` (not `get_token()` as sketched in Phase 1). All providers
  (`ProfileTokenProvider`, `OboTokenProvider`) and `GraphClient._get_headers`
  use that name.
- **Thread offload.** The blocking MSAL call in stdio mode runs via
  `asyncio.to_thread(...)` (not `anyio.to_thread.run_sync`). Functionally
  equivalent; one less indirection.
- **Endpoint path.** The Streamable HTTP endpoint is `/mcp` (no trailing slash);
  `/mcp/` 307-redirects to it. RFC 9728 metadata is served at
  `/.well-known/oauth-protected-resource/mcp`.
- **Rate limiting is a bounded per-user middleware (post-review fix).** fastmcp's
  `RateLimitingMiddleware` is replaced by `middleware.UserRateLimitMiddleware`.
  Two reasons: (1) fastmcp keys every request under a single literal `"global"`
  bucket unless a `get_client_id` callable is supplied, letting one user throttle
  all others; and (2) its per-client `defaultdict` of limiters never evicts, so
  entries accumulate for the whole process lifetime. Ours keys the token bucket
  on `f"{tid}:{oid}"` (falling back to `oid`, then `sub`, then a shared
  `"unauthenticated"` bucket; never raises), and stores buckets in an
  `OrderedDict` capped at `_LIMITER_CAP` (10,000, LRU eviction) with a lazy
  idle-TTL sweep (`_LIMITER_IDLE_TTL`, 900 s). User-facing semantics are
  unchanged: same token-bucket algorithm, `burst_capacity` still defaults to
  `2 * max_requests_per_second`, applied in `on_request`, and over-limit raises
  the same `RateLimitError` (McpError code `-32000`).
- **Single-tenant GUID required for `MCP_AUTH_TENANT_ID` (post-review fix).**
  `validate_http_config` now requires a concrete tenant GUID and rejects
  pseudo-tenants (`organizations`/`common`/`consumers`) and verified domains
  with an actionable error. fastmcp's `AzureProvider` pins the accepted token
  issuer to a literal `https://{authority}/{tenant_id}/v2.0` (unlike
  `AzureJWTVerifier`, it has no pseudo-tenant special-casing), and real Entra
  tokens carry the concrete tenant GUID in `iss`, so any non-GUID value makes
  every request fail authentication. Multi-tenant support is future work: it
  needs issuer-validation skipping plus per-tenant OBO authority selection.

## Phase 7 — observability (added 2026-07-12)

DevOps traffic/usage metrics for http mode, **off by default** and enabled
only when `MCP_STATS_TOKEN` is set. stdio mode is untouched — no registry, no
middleware, no routes.

- **`src/mcp_microsoft/metrics.py` (new).** `MetricsRegistry` aggregates in
  memory on the event loop (documented no-await/no-lock invariant): global
  totals, a rolling 60-bucket per-minute traffic timeline (lazily
  zero-filled), per-tool latency (a `deque(maxlen=256)` of ms → p50/p95/avg at
  snapshot time via linear-interpolation percentile), and per-user activity in
  an `OrderedDict` capped at 1000 with least-recently-seen eviction
  (`users_evicted` counter). `record()` is the single write API; `snapshot()`
  returns a JSON-able view; `render_prometheus()` hand-rolls the 0.0.4 text
  exposition with escaped label values and **deliberately no per-user label
  series** (unbounded-cardinality anti-pattern). Module-level
  `get_metrics_registry()` / `reset_metrics_registry()` singleton, with the
  reset wired into `runtime.reset_runtime_state()`.
- **`MetricsMiddleware` (`middleware.py`).** Sibling of
  `AuditLoggingMiddleware`; the shared identity helper `_caller_identity` was
  lifted to module level so both use it (the old staticmethod remains as a
  thin alias for existing callers/tests). Registered in http mode **after** the
  audit middleware (innermost, so it times the tool itself), records in both
  success and error paths, and wraps the registry write so it can never raise
  into or mask the observed call.
- **Routes (`server.py`), http mode + non-empty `stats_token` only.**
  `config.stats_token` (env `MCP_STATS_TOKEN`, stripped, optional — not in
  `validate_http_config`'s required set). A timing-safe auth helper accepts
  either `Authorization: Bearer <token>` or HTTP Basic (any user, password ==
  token) so browsers can open `/dashboard` natively; failure → 401 with
  `WWW-Authenticate: Basic realm="mcp-microsoft stats", Bearer` and no body.
  `GET /metrics` (Prometheus text), `GET /stats` (JSON snapshot), and
  `GET /dashboard` (one self-contained, dependency-free HTML page polling
  `/stats` every 10s). When `stats_token` is empty none of the three are
  registered (info log). `/health` is unchanged.
- **Tests.** `tests/test_metrics.py` (registry math + middleware) and
  `tests/test_observability.py` (ASGI route gating/content), plus two
  `stats_token` config assertions in `tests/test_http_transport.py`.
- **No new dependencies** (stdlib + starlette responses, already transitive).
  No tool payloads/args/results and no tokens are ever recorded or logged —
  only `oid`/`username`, the same exposure as the audit log.

## Phase 8 — file upload app (2026-07-12)

- **Goal.** Kill the "base64 as a tool argument" upload path that pushes a
  file's whole content through the model context window. Adopt FastMCP's
  `FileUpload` app (`fastmcp[apps]` → prefab-ui 0.20.2): users drag-drop files
  into an interactive UI, the bytes go straight to the server (bypassing the
  model), and the Graph upload tools consume them by name.
- **Dependency.** `fastmcp[azure]>=3.4.4` → `fastmcp[azure,apps]>=3.4.4`;
  `uv lock`/`uv sync` pull prefab-ui 0.20.2. NB the class imports as
  `from fastmcp.apps.file_upload import FileUpload` — `fastmcp.apps.__init__`
  only lazily re-exports `FastMCPApp`, not `FileUpload`.
- **`src/mcp_microsoft/uploads.py` — `ScopedFileUpload(FileUpload)`.**
  - *Scoping* (`_get_scope_key`): http mode keys on the caller's Entra `oid`
    claim (via `get_access_token`, same as `middleware._caller_identity`) —
    stable across reconnects and stateless HTTP where the stock `session_id`
    keying is not — falling back to session id then `"__default__"`; stdio
    keeps the per-session default. Never raises.
  - *Bounds* over the otherwise-unbounded in-memory store: per-scope max files
    (20) and max bytes (100 MB); global whole-scope LRU cap (1000, mirroring
    `metrics` user-cap); idle-scope TTL prune (2h, lazy sweep like the rate
    limiter). Over-quota `on_store` raises `ValueError`; `scopes_evicted` /
    `scopes_pruned` counters exposed as attributes. Byte accounting uses true
    decoded size; name length capped at 255. Content never logged.
  - *Accessors*: module singleton (`get`/`set`/`reset_upload_provider`, reset
    wired into `runtime.reset_runtime_state`) plus
    `resolve_uploaded_file(name) -> (bytes, content_type)` reading the current
    request's scope; raises `ValueError` (feature-off message, or missing with
    available names).
- **Config** (`config.py`): `enable_file_upload: bool | None`
  (`MCP_ENABLE_FILE_UPLOAD`, tri-state; default on in http, off in stdio via
  `feature_flags.is_file_upload_enabled`) and `upload_max_mb: int = 10`
  (`MCP_UPLOAD_MAX_MB`, positive; `feature_flags.resolve_upload_max_bytes`
  rejects ≤ 0).
- **Wiring** (`server.py`): when enabled, construct
  `ScopedFileUpload(max_file_size=upload_max_mb*1MB)`, `set_upload_provider`,
  `mcp.add_provider(provider)` under the default (empty) namespace — its tool
  names (`store_files`/`list_files`/`read_file`/`file_manager`) don't collide
  with our 95 tools. Info-logged either way. Provider tools go through the
  http middleware stack (rate limit / audit / metrics) — proven by test.
  Note: `store_files` is not model-visible (absent from model tool listing), but
  it IS callable programmatically via its hashed backend name (the same name the
  UI's CallTool uses) — it is not UI-only or unreachable. That is safe because
  such calls flow through the same middleware (per-user rate limit, audit,
  metrics), the per-file `max_file_size` check, the per-scope file/byte quotas,
  and the global encoded-byte budget.
- **Tool integration**: `upload_file` (onedrive) and `upload_to_site`
  (sharepoint) gain `uploaded_file: str | None` — mutually exclusive with
  `local_path`/`content_base64`, filename defaults to the stored name, bytes
  reused through the EXISTING small-PUT / chunked-session paths via a temp
  file. Feature-off / not-found raise a clear `ValueError`.
- **Tests** (`tests/test_uploads.py` new, `test_http_transport.py` +
  `test_observability.py` extended): store quotas + atomicity, oid scope
  isolation, LRU eviction + TTL prune (fake clock), name cap, resolve
  happy/missing/feature-off, registration matrix (http default on / stdio
  default off / explicit both ways), middleware-coverage proof, and
  tool-integration (uploaded_file → Graph PUT, mutual exclusion, feature off).
  Suite: 299 → 323 green.
