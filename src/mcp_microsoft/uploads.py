"""Context-free file uploads via fastmcp's FileUpload app (http + stdio).

Wraps :class:`fastmcp.apps.file_upload.FileUpload` (the ``fastmcp[apps]`` /
prefab-ui drag-and-drop provider) with the two things this repo requires before
it can be exposed on a long-running multi-user server:

* **Per-caller scoping.** The stock store partitions by ``ctx.session_id``,
  which is *not* stable in stateless http mode (every request is a new session).
  :meth:`ScopedFileUpload._get_scope_key` instead keys http-mode storage on the
  caller's validated Entra ``oid`` claim (falling back to ``sub``) — resolved the
  same way ``middleware._rate_limit_key`` does, via the ambient bearer token — so
  a user's uploads survive reconnects and stateless requests, and no two users
  can ever see each other's files. When http mode can supply *neither* ``oid``
  nor ``sub`` the resolver **refuses** (raises :class:`ToolError`) rather than
  fall back to a shared bucket — mirroring the rate limiter's identity
  philosophy. stdio mode keeps the default per-session behavior (session id, then
  a shared ``"__default__"`` bucket for the single local user).
* **Bounds.** The stock in-memory store is unbounded (files, bytes, and distinct
  scopes all grow without limit). This subclass caps per-scope file count and
  total *decoded* bytes, caps total *encoded* bytes across ALL scopes with a
  global budget, LRU-evicts whole idle scopes past a global scope-count cap, and
  lazily prunes scopes idle past a TTL — mirroring the bounded-store patterns
  already used by ``metrics.MetricsRegistry`` (user cap) and
  ``middleware.UserRateLimitMiddleware`` (LRU + idle-TTL sweep). Over-quota
  :meth:`~ScopedFileUpload.on_store` raises a clear :class:`ToolError` naming the
  offending limit (a ``ToolError`` survives http-mode error masking, so the
  actionable message reaches the caller).

**Encoded vs decoded accounting.** Two different byte totals are tracked on
purpose. The *per-scope* quota is measured in **decoded** bytes — the
user-meaningful "how big is my file" number. The *global* budget is measured in
**encoded** bytes — the length of the base64 string actually held in RAM (~4/3
of the decoded size), which is what genuinely constrains process memory.

**Thread-safety.** fastmcp runs sync tools (``store_files``/``list_files``/
``read_file``) in a threadpool (``call_sync_fn_in_threadpool``), so these
storage callbacks execute on worker threads concurrently. A single
``threading.Lock`` (:attr:`ScopedFileUpload._lock`) guards every read and write
of the shared store; critical sections stay small and never do I/O (base64
decode happens *after* the lock is released, on a copied-out reference).

The Graph upload tools (``upload_file``, ``upload_to_site``) consume stored files
*by name* via :func:`resolve_uploaded_file`, so the raw bytes flow from the UI
straight into a Graph upload without ever passing through the model's context
window.

**store_files reachability.** The drag-and-drop backend tool ``store_files`` is
registered (by the base class) under a hashed backend name and is therefore
callable programmatically via that name — it is *not* UI-only or unreachable.
That is safe: such calls flow through the exact same middleware stack as every
other tool (per-user rate limit, audit log, metrics), the per-file
``max_file_size`` check, the per-scope file/byte quotas, and the global budget
enforced here.

**Privacy.** The raw base64 ``data`` of every uploaded file sits in this
process's memory for the lifetime of its scope. File *content* is never logged
(only names/sizes appear in summaries), matching the audit middleware's rule.
"""

from __future__ import annotations

import base64
import threading
import time
from collections import OrderedDict
from datetime import datetime, timezone
from typing import Any

from fastmcp.apps.file_upload import (
    FileUpload,
    _b64_decoded_size,
    _format_size,
    _make_summary,
    _TEXT_EXTENSIONS,
)
from fastmcp.exceptions import ToolError

# ---------------------------------------------------------------------------
# Default bounds. Chosen to be generous for interactive use yet firmly capped
# so a hostile or careless caller cannot exhaust process memory.
# ---------------------------------------------------------------------------

_DEFAULT_MAX_FILES_PER_SCOPE = 20
_DEFAULT_MAX_BYTES_PER_SCOPE = 100 * 1024 * 1024  # 100 MB per user (decoded)
_DEFAULT_MAX_SCOPES = 1000  # distinct users tracked before whole-scope LRU eviction
_DEFAULT_IDLE_TTL = 2 * 60 * 60.0  # 2h; scopes idle longer are lazily pruned
_DEFAULT_MAX_NAME_LENGTH = 255  # reject absurd client-controlled dict keys
_DEFAULT_GLOBAL_BUDGET_BYTES = 1024 * 1024 * 1024  # 1 GB of base64 across all scopes

# read_file returns decoded text straight into the model's context window, so the
# returned ``content`` is capped here (the full bytes still flow to Graph uploads
# via resolve()). A file larger than this is truncated with a flag + note.
_READ_CONTENT_CAP = 48 * 1024  # 48 KB of returned text


def _ambient_oid_sub() -> tuple[str | None, str | None]:
    """Return the caller's ``(oid, sub)`` claims from the ambient bearer token.

    Mirrors ``middleware._caller_tid_oid_sub``: reads the validated access token
    via fastmcp's request-scoped dependency. Never raises — an absent token, a
    missing claim, or a raising dependency all collapse to ``None`` for the
    affected field. The bearer token itself is never read into the return value.
    """
    try:
        from fastmcp.server.dependencies import get_access_token

        token = get_access_token()
    except Exception:
        return None, None
    if token is None:
        return None, None
    claims = getattr(token, "claims", None) or {}
    oid = claims.get("oid")
    sub = claims.get("sub")
    return (str(oid) if oid else None, str(sub) if sub else None)


def _ctx_session_id(ctx: Any) -> str | None:
    """Return ``ctx.session_id`` if available, else None. Never raises."""
    if ctx is None:
        return None
    try:
        sid = ctx.session_id
    except Exception:
        return None
    return str(sid) if sid else None


def _is_http_mode() -> bool:
    """Return True when the process is configured for http (multi-user) mode."""
    try:
        from mcp_microsoft.config import get_app_config

        return get_app_config().transport == "http"
    except Exception:
        return False


def _current_context() -> Any:
    """Return the ambient fastmcp Context for the current request, else None."""
    try:
        from fastmcp.server.dependencies import get_context

        return get_context()
    except Exception:
        return None


class ScopedFileUpload(FileUpload):
    """Per-caller-scoped, bounded in-memory FileUpload provider.

    See the module docstring for the scoping and bounds rationale. All storage
    lives in ``self._store`` — an ``OrderedDict`` mapping a scope key to a
    ``{"files": OrderedDict[name -> entry], "bytes": int, "encoded_bytes": int,
    "last_seen": float}`` record, ordered by recency of access so the
    least-recently-used scope sits at the head for O(1) eviction/prune. ``bytes``
    is the scope's decoded footprint (per-scope quota); ``encoded_bytes`` is the
    base64 footprint that feeds :attr:`_global_encoded_bytes` (global budget).

    All access to ``self._store`` is serialized by :attr:`_lock`; every override
    runs on a threadpool worker (see the module docstring's thread-safety note).
    """

    def __init__(
        self,
        *,
        max_file_size: int = 10 * 1024 * 1024,
        max_files_per_scope: int = _DEFAULT_MAX_FILES_PER_SCOPE,
        max_bytes_per_scope: int = _DEFAULT_MAX_BYTES_PER_SCOPE,
        max_scopes: int = _DEFAULT_MAX_SCOPES,
        idle_ttl: float = _DEFAULT_IDLE_TTL,
        max_name_length: int = _DEFAULT_MAX_NAME_LENGTH,
        global_budget_bytes: int = _DEFAULT_GLOBAL_BUDGET_BYTES,
        clock: Any = time.monotonic,
        name: str = "Files",
        **kwargs: Any,
    ) -> None:
        super().__init__(name, max_file_size=max_file_size, **kwargs)
        self._max_files_per_scope = max_files_per_scope
        self._max_bytes_per_scope = max_bytes_per_scope
        self._max_scopes = max_scopes
        self._idle_ttl = idle_ttl
        self._max_name_length = max_name_length
        self._global_budget_bytes = global_budget_bytes
        self._clock = clock
        # Serializes every read/write of self._store. fastmcp dispatches the sync
        # storage tools on threadpool workers, so concurrent on_store/on_list/
        # on_read/resolve calls race on the shared OrderedDict without this.
        self._lock = threading.Lock()
        # Replace the base's unbounded dict with an LRU-ordered store. Every
        # method below (on_store/on_list/on_read/resolve) is overridden, so the
        # base never touches this attribute with its own shape assumptions.
        self._store: "OrderedDict[str, dict[str, Any]]" = OrderedDict()
        # Running total of base64-encoded bytes held across ALL scopes — the true
        # RAM footprint the global budget bounds. Adjusted on every store,
        # overwrite, eviction, and prune. Guarded by self._lock.
        self._global_encoded_bytes = 0
        # Observability counters (plain attributes; no metrics-registry coupling).
        self.scopes_evicted = 0  # whole scopes dropped by the global LRU cap
        self.scopes_pruned = 0  # whole scopes dropped by the idle-TTL sweep

    def __repr__(self) -> str:
        return f"ScopedFileUpload({self.name!r})"

    # ------------------------------------------------------------------
    # Scoping
    # ------------------------------------------------------------------

    def _get_scope_key(self, ctx: Any) -> str:
        """Return the storage partition key for the current request.

        http mode: the caller's ``oid`` claim (stable across reconnects and
        stateless requests), falling back to ``sub``. When neither is present the
        request has no stable per-user identity, so rather than dump uploads into
        a shared bucket (a cross-user leak) this **refuses** by raising
        :class:`ToolError` — mirroring ``middleware._rate_limit_key``'s
        oid → sub philosophy. stdio mode: the per-session default (the base class
        behavior), falling back to a shared ``"__default__"`` bucket for the
        single local user.

        Raises:
            ToolError: in http mode when neither ``oid`` nor ``sub`` is available.
        """
        if _is_http_mode():
            oid, sub = _ambient_oid_sub()
            if oid:
                return f"oid:{oid}"
            if sub:
                return f"sub:{sub}"
            raise ToolError(
                "cannot determine a stable user identity for upload storage"
            )

        sid = _ctx_session_id(ctx)
        return f"sid:{sid}" if sid else "__default__"

    # ------------------------------------------------------------------
    # Bounds housekeeping
    # ------------------------------------------------------------------

    def _prune_idle(self, now: float) -> None:
        """Drop whole scopes idle longer than the TTL, sweeping from the LRU head.

        Scopes are ordered by last access (``move_to_end`` on every touch), so the
        head is the least-recently-seen: once it is within the TTL every later
        scope is too, and the sweep stops. O(1) amortized per request. Decrements
        the global encoded-byte total for each dropped scope.

        Caller must hold :attr:`_lock` (the lock is not re-entrant).
        """
        while self._store:
            _key, entry = next(iter(self._store.items()))
            if now - entry["last_seen"] <= self._idle_ttl:
                break
            _key, dropped = self._store.popitem(last=False)
            self._global_encoded_bytes -= dropped.get("encoded_bytes", 0)
            self.scopes_pruned += 1

    def _enforce_scope_cap(self) -> None:
        """Evict least-recently-used whole scopes until under the global cap.

        Decrements the global encoded-byte total for each evicted scope. Caller
        must hold :attr:`_lock` (the lock is not re-entrant).
        """
        while len(self._store) > self._max_scopes:
            _key, dropped = self._store.popitem(last=False)
            self._global_encoded_bytes -= dropped.get("encoded_bytes", 0)
            self.scopes_evicted += 1

    def _touch(self, scope: str, entry: dict[str, Any], now: float) -> None:
        """Mark *scope* most-recently-used and stamp its idle clock.

        Caller must hold :attr:`_lock`.
        """
        entry["last_seen"] = now
        self._store.move_to_end(scope)

    # ------------------------------------------------------------------
    # Storage interface (overrides FileUpload's in-memory defaults)
    # ------------------------------------------------------------------

    def on_store(
        self,
        files: list[dict[str, Any]],
        ctx: Any,
    ) -> list[dict[str, Any]]:
        """Store uploaded files under the caller's scope, enforcing quotas.

        Validates name length and the resulting per-scope file-count, per-scope
        decoded-byte quota, and global encoded-byte budget *before* committing
        anything (over-quota is atomic — no partial store). Files with an existing
        name overwrite in place. Per-scope accounting uses the true decoded length
        of each base64 payload; the global budget uses the encoded (base64)
        length actually held in RAM.

        Raises:
            ToolError: if any name exceeds the length cap, or the store would
                exceed the per-scope file/byte quota or the global budget. The
                message names the offending limit.
        """
        now = self._clock()
        scope = self._get_scope_key(ctx)  # may raise ToolError (http, no identity)

        # Name-length cap (reject, don't silently truncate). Pure input
        # validation — no shared state, so it runs before taking the lock.
        for f in files:
            name = f.get("name", "")
            if len(name) > self._max_name_length:
                raise ToolError(
                    f"Uploaded file name is too long ({len(name)} chars; the "
                    f"limit is {self._max_name_length}). Rename the file and "
                    "upload it again."
                )

        # Decoded (per-scope) and encoded (global) sizes of the incoming batch.
        incoming_sizes = {f["name"]: _b64_decoded_size(f.get("data", "")) for f in files}
        incoming_encoded = {f["name"]: len(f.get("data", "")) for f in files}

        with self._lock:
            self._prune_idle(now)
            entry = self._store.get(scope)
            files_map: "OrderedDict[str, dict[str, Any]]" = (
                entry["files"] if entry is not None else OrderedDict()
            )
            base_bytes = entry["bytes"] if entry is not None else 0
            base_encoded = entry["encoded_bytes"] if entry is not None else 0

            # Prospective post-store shape (overwrites reuse a slot).
            prospective_count = len(set(files_map) | set(incoming_sizes))
            if prospective_count > self._max_files_per_scope:
                raise ToolError(
                    f"Upload rejected: it would leave {prospective_count} files in "
                    f"your upload area, over the per-user limit of "
                    f"{self._max_files_per_scope}. Remove or overwrite existing "
                    "uploads first."
                )
            replaced = sum(files_map[n]["size"] for n in incoming_sizes if n in files_map)
            prospective_bytes = base_bytes - replaced + sum(incoming_sizes.values())
            if prospective_bytes > self._max_bytes_per_scope:
                raise ToolError(
                    f"Upload rejected: it would use {_format_size(prospective_bytes)} "
                    f"of upload storage, over the per-user limit of "
                    f"{_format_size(self._max_bytes_per_scope)}. Remove existing "
                    "uploads first."
                )
            replaced_encoded = sum(
                files_map[n]["encoded_size"] for n in incoming_encoded if n in files_map
            )
            prospective_scope_encoded = (
                base_encoded - replaced_encoded + sum(incoming_encoded.values())
            )
            prospective_global = (
                self._global_encoded_bytes - replaced_encoded + sum(incoming_encoded.values())
            )
            if prospective_global > self._global_budget_bytes:
                raise ToolError(
                    "Upload rejected: server upload storage is full "
                    f"({_format_size(self._global_encoded_bytes)} of "
                    f"{_format_size(self._global_budget_bytes)} in use across all "
                    "users). Try again later."
                )

            # Commit (only now do we materialize a new scope).
            if entry is None:
                entry = {
                    "files": OrderedDict(),
                    "bytes": 0,
                    "encoded_bytes": 0,
                    "last_seen": now,
                }
                self._store[scope] = entry
                files_map = entry["files"]
            for f in files:
                name = f["name"]
                files_map[name] = {
                    "name": name,
                    "size": incoming_sizes[name],
                    "encoded_size": incoming_encoded[name],
                    "type": f.get("type") or "application/octet-stream",
                    "data": f["data"],
                    "uploaded_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
                }
            entry["bytes"] = prospective_bytes
            entry["encoded_bytes"] = prospective_scope_encoded
            self._global_encoded_bytes = prospective_global
            self._touch(scope, entry, now)
            self._enforce_scope_cap()

            return [_make_summary(e) for e in files_map.values()]

    def on_list(self, ctx: Any) -> list[dict[str, Any]]:
        """List the current scope's stored files (summaries only)."""
        now = self._clock()
        scope = self._get_scope_key(ctx)  # may raise ToolError (http, no identity)
        with self._lock:
            self._prune_idle(now)
            entry = self._store.get(scope)
            if entry is None:
                return []
            self._touch(scope, entry, now)
            return [_make_summary(e) for e in entry["files"].values()]

    def on_read(self, name: str, ctx: Any) -> dict[str, Any]:
        """Read a stored file's preview by name (the model-visible ``read_file``).

        Text files are decoded to a ``content`` string, **capped at**
        :data:`_READ_CONTENT_CAP` so a large file cannot flood the model's
        context window; a truncated result carries ``"truncated": True`` and a
        ``"note"`` explaining that the full file can be sent whole to
        OneDrive/SharePoint via ``upload_file``/``upload_to_site`` with
        ``uploaded_file``. Non-text files return a short ``content_base64``
        preview (unchanged). The full bytes are exposed only to the Graph upload
        tools via :meth:`resolve`, never to the model here.

        The signature stays ``(name, ctx)`` — the base class registers the
        ``read_file`` tool with that fixed signature and reusing that registration
        (rather than copying its prefab-ui ``file_manager`` UI wholesale) is the
        maintainable choice — so paging via an ``offset`` parameter is not offered;
        the note steers oversized reads to a whole-file upload instead.

        Raises:
            ToolError: if the file is not found (message lists available names).
        """
        now = self._clock()
        scope = self._get_scope_key(ctx)  # may raise ToolError (http, no identity)
        with self._lock:
            entry = self._store.get(scope)
            files_map = entry["files"] if entry else {}
            if name not in files_map:
                raise ToolError(
                    f"File {name!r} not found. Available: {list(files_map)}"
                )
            self._touch(scope, entry, now)  # type: ignore[arg-type]
            fe = files_map[name]
            result: dict[str, Any] = {
                "name": fe["name"],
                "size": fe["size"],
                "type": fe["type"],
                "uploaded_at": fe["uploaded_at"],
            }
            is_text = fe["type"].startswith("text/") or any(
                fe["name"].endswith(ext) for ext in _TEXT_EXTENSIONS
            )
            data_b64 = fe["data"]  # copy the reference out; decode after unlock

        # Base64 decode is CPU work on potentially megabytes — do it off the lock.
        if is_text:
            try:
                text = base64.b64decode(data_b64).decode("utf-8")
            except UnicodeDecodeError:
                result["content_base64"] = data_b64[:200] + "..."
            else:
                if len(text) > _READ_CONTENT_CAP:
                    result["content"] = text[:_READ_CONTENT_CAP]
                    result["truncated"] = True
                    result["note"] = (
                        f"Showing the first {_READ_CONTENT_CAP // 1024} KB of a "
                        f"{_format_size(result['size'])} file. To use the whole "
                        "file, send it directly to OneDrive/SharePoint via "
                        "upload_file or upload_to_site with uploaded_file="
                        f"{name!r} (its bytes never pass through this preview)."
                    )
                else:
                    result["content"] = text
        else:
            result["content_base64"] = data_b64[:200] + "..."
        return result

    # ------------------------------------------------------------------
    # Tool integration
    # ------------------------------------------------------------------

    def resolve(self, name: str) -> tuple[bytes, str]:
        """Return ``(decoded_bytes, content_type)`` for *name* in the current scope.

        Reads from the CURRENT request's scope (resolved from the ambient
        context/token), decodes the stored base64, and returns the raw bytes so a
        Graph upload tool can stream them without the model ever seeing the
        content.

        Raises:
            ToolError: if the named file is not in the caller's scope (message
                lists the available upload names), or — in http mode — the caller
                has no stable identity.
            ValueError: if the stored data is not valid base64 (an internal
                corruption; masked in http mode, which is correct).
        """
        ctx = _current_context()
        scope = self._get_scope_key(ctx)  # may raise ToolError (http, no identity)
        with self._lock:
            entry = self._store.get(scope)
            files_map = entry["files"] if entry else {}
            if name not in files_map:
                raise ToolError(
                    f"No uploaded file named {name!r}. Available uploads: "
                    f"{sorted(files_map)}. Upload it via the file-upload UI first "
                    "(see list_files)."
                )
            fe = files_map[name]
            data_b64 = fe["data"]  # copy out; decode after releasing the lock
            content_type = fe.get("type") or "application/octet-stream"

        try:
            data = base64.b64decode(data_b64, validate=True)
        except Exception as exc:  # malformed stored payload
            raise ValueError(
                f"Uploaded file {name!r} could not be decoded: {exc}"
            ) from exc
        return data, content_type


# ---------------------------------------------------------------------------
# Process-wide singleton (wired in server.create_mcp_server; reset in runtime)
# ---------------------------------------------------------------------------

_provider: ScopedFileUpload | None = None


def get_upload_provider() -> ScopedFileUpload | None:
    """Return the process-wide upload provider, or None when the feature is off."""
    return _provider


def set_upload_provider(provider: ScopedFileUpload | None) -> None:
    """Install (or clear) the process-wide upload provider."""
    global _provider
    _provider = provider


def reset_upload_provider() -> None:
    """Drop the cached provider so a rebuild re-wires it. Wired into runtime."""
    global _provider
    _provider = None


def resolve_uploaded_file(name: str) -> tuple[bytes, str]:
    """Resolve a previously uploaded file *name* to ``(bytes, content_type)``.

    The single entry point the Graph upload tools call. Raises :class:`ToolError`
    with an actionable message (safe to surface past http-mode masking) when the
    feature is disabled or the file is not found in the caller's scope.
    """
    provider = get_upload_provider()
    if provider is None:
        raise ToolError(
            "file upload is not enabled on this server (MCP_ENABLE_FILE_UPLOAD)"
        )
    return provider.resolve(name)
