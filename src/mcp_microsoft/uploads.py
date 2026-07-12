"""Context-free file uploads via fastmcp's FileUpload app (http + stdio).

Wraps :class:`fastmcp.apps.file_upload.FileUpload` (the ``fastmcp[apps]`` /
prefab-ui drag-and-drop provider) with the two things this repo requires before
it can be exposed on a long-running multi-user server:

* **Per-caller scoping.** The stock store partitions by ``ctx.session_id``,
  which is *not* stable in stateless http mode (every request is a new session).
  :meth:`ScopedFileUpload._get_scope_key` instead keys http-mode storage on the
  caller's validated Entra ``oid`` claim — resolved the same way
  ``middleware._caller_identity`` does, via the ambient bearer token — so a
  user's uploads survive reconnects and stateless requests, and no two users can
  ever see each other's files. stdio mode keeps the default per-session
  behavior. The key resolver never raises.
* **Bounds.** The stock in-memory store is unbounded (files, bytes, and distinct
  scopes all grow without limit). This subclass caps per-scope file count and
  total bytes, LRU-evicts whole idle scopes past a global cap, and lazily prunes
  scopes idle past a TTL — mirroring the bounded-store patterns already used by
  ``metrics.MetricsRegistry`` (user cap) and ``middleware.UserRateLimitMiddleware``
  (LRU + idle-TTL sweep). Over-quota :meth:`~ScopedFileUpload.on_store` raises a
  clear ``ValueError`` naming the offending limit.

The Graph upload tools (``upload_file``, ``upload_to_site``) consume stored files
*by name* via :func:`resolve_uploaded_file`, so the raw bytes flow from the UI
straight into a Graph upload without ever passing through the model's context
window.

**Privacy.** The raw base64 ``data`` of every uploaded file sits in this
process's memory for the lifetime of its scope. File *content* is never logged
(only names/sizes appear in summaries), matching the audit middleware's rule.
"""

from __future__ import annotations

import base64
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

# ---------------------------------------------------------------------------
# Default bounds. Chosen to be generous for interactive use yet firmly capped
# so a hostile or careless caller cannot exhaust process memory.
# ---------------------------------------------------------------------------

_DEFAULT_MAX_FILES_PER_SCOPE = 20
_DEFAULT_MAX_BYTES_PER_SCOPE = 100 * 1024 * 1024  # 100 MB per user
_DEFAULT_MAX_SCOPES = 1000  # distinct users tracked before whole-scope LRU eviction
_DEFAULT_IDLE_TTL = 2 * 60 * 60.0  # 2h; scopes idle longer are lazily pruned
_DEFAULT_MAX_NAME_LENGTH = 255  # reject absurd client-controlled dict keys


def _ambient_oid() -> str | None:
    """Return the caller's ``oid`` claim from the ambient bearer token, else None.

    Mirrors ``middleware._caller_identity``: reads the validated access token via
    fastmcp's request-scoped dependency. Never raises — an absent token, a
    missing claim, or a raising dependency all collapse to ``None``. The bearer
    token itself is never read into the return value.
    """
    try:
        from fastmcp.server.dependencies import get_access_token

        token = get_access_token()
    except Exception:
        return None
    if token is None:
        return None
    claims = getattr(token, "claims", None) or {}
    oid = claims.get("oid")
    return str(oid) if oid else None


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
    ``{"files": OrderedDict[name -> entry], "bytes": int, "last_seen": float}``
    record, ordered by recency of access so the least-recently-used scope sits at
    the head for O(1) eviction/prune.
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
        self._clock = clock
        # Replace the base's unbounded dict with an LRU-ordered store. Every
        # method below (on_store/on_list/on_read/resolve) is overridden, so the
        # base never touches this attribute with its own shape assumptions.
        self._store: "OrderedDict[str, dict[str, Any]]" = OrderedDict()
        # Observability counters (plain attributes; no metrics-registry coupling).
        self.scopes_evicted = 0  # whole scopes dropped by the global LRU cap
        self.scopes_pruned = 0  # whole scopes dropped by the idle-TTL sweep

    def __repr__(self) -> str:
        return f"ScopedFileUpload({self.name!r})"

    # ------------------------------------------------------------------
    # Scoping
    # ------------------------------------------------------------------

    def _get_scope_key(self, ctx: Any) -> str:
        """Return the storage partition key for the current request. Never raises.

        http mode: the caller's ``oid`` claim (stable across reconnects and
        stateless requests), falling back to the session id, then a shared
        ``"__default__"`` bucket. stdio mode: the per-session default (the base
        class behavior), falling back to ``"__default__"``.
        """
        if _is_http_mode():
            oid = _ambient_oid()
            if oid:
                return f"oid:{oid}"
            sid = _ctx_session_id(ctx)
            if sid:
                return f"sid:{sid}"
            return "__default__"

        sid = _ctx_session_id(ctx)
        return f"sid:{sid}" if sid else "__default__"

    # ------------------------------------------------------------------
    # Bounds housekeeping
    # ------------------------------------------------------------------

    def _prune_idle(self, now: float) -> None:
        """Drop whole scopes idle longer than the TTL, sweeping from the LRU head.

        Scopes are ordered by last access (``move_to_end`` on every touch), so the
        head is the least-recently-seen: once it is within the TTL every later
        scope is too, and the sweep stops. O(1) amortized per request.
        """
        while self._store:
            _key, entry = next(iter(self._store.items()))
            if now - entry["last_seen"] <= self._idle_ttl:
                break
            self._store.popitem(last=False)
            self.scopes_pruned += 1

    def _enforce_scope_cap(self) -> None:
        """Evict least-recently-used whole scopes until under the global cap."""
        while len(self._store) > self._max_scopes:
            self._store.popitem(last=False)
            self.scopes_evicted += 1

    def _touch(self, scope: str, entry: dict[str, Any], now: float) -> None:
        """Mark *scope* most-recently-used and stamp its idle clock."""
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

        Validates name length and the resulting per-scope file-count and
        total-byte quotas *before* committing anything (over-quota is atomic — no
        partial store). Files with an existing name overwrite in place. Byte
        accounting uses the true decoded length of each base64 payload, not the
        client-reported ``size``.

        Raises:
            ValueError: if any name exceeds the length cap, or the store would
                exceed the per-scope file or byte quota. The message names the
                offending limit.
        """
        now = self._clock()
        self._prune_idle(now)
        scope = self._get_scope_key(ctx)
        entry = self._store.get(scope)
        if entry is None:
            entry = {"files": OrderedDict(), "bytes": 0, "last_seen": now}
            self._store[scope] = entry
        else:
            self._store.move_to_end(scope)
        files_map: "OrderedDict[str, dict[str, Any]]" = entry["files"]

        # 1. Name-length cap (reject, don't silently truncate).
        for f in files:
            name = f.get("name", "")
            if len(name) > self._max_name_length:
                raise ValueError(
                    f"Uploaded file name is too long ({len(name)} chars; the "
                    f"limit is {self._max_name_length}). Rename the file and "
                    "upload it again."
                )

        # 2. Compute the prospective post-store shape (overwrites reuse a slot).
        incoming_sizes = {f["name"]: _b64_decoded_size(f.get("data", "")) for f in files}
        prospective_count = len(set(files_map) | set(incoming_sizes))
        if prospective_count > self._max_files_per_scope:
            raise ValueError(
                f"Upload rejected: it would leave {prospective_count} files in "
                f"your upload area, over the per-user limit of "
                f"{self._max_files_per_scope}. Remove or overwrite existing "
                "uploads first."
            )
        replaced = sum(files_map[n]["size"] for n in incoming_sizes if n in files_map)
        prospective_bytes = entry["bytes"] - replaced + sum(incoming_sizes.values())
        if prospective_bytes > self._max_bytes_per_scope:
            raise ValueError(
                f"Upload rejected: it would use {_format_size(prospective_bytes)} "
                f"of upload storage, over the per-user limit of "
                f"{_format_size(self._max_bytes_per_scope)}. Remove existing "
                "uploads first."
            )

        # 3. Commit.
        for f in files:
            name = f["name"]
            files_map[name] = {
                "name": name,
                "size": incoming_sizes[name],
                "type": f.get("type") or "application/octet-stream",
                "data": f["data"],
                "uploaded_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            }
        entry["bytes"] = prospective_bytes
        self._touch(scope, entry, now)
        self._enforce_scope_cap()

        return [_make_summary(e) for e in files_map.values()]

    def on_list(self, ctx: Any) -> list[dict[str, Any]]:
        """List the current scope's stored files (summaries only)."""
        now = self._clock()
        self._prune_idle(now)
        scope = self._get_scope_key(ctx)
        entry = self._store.get(scope)
        if entry is None:
            return []
        self._touch(scope, entry, now)
        return [_make_summary(e) for e in entry["files"].values()]

    def on_read(self, name: str, ctx: Any) -> dict[str, Any]:
        """Read a stored file's preview by name (the model-visible ``read_file``).

        Replicates the base class's text/binary preview behavior: text files are
        decoded to a ``content`` string; everything else returns a truncated
        ``content_base64`` preview. The full bytes are exposed only to the Graph
        upload tools via :meth:`resolve`, never to the model here.

        Raises:
            ValueError: if the file is not found (message lists available names).
        """
        now = self._clock()
        scope = self._get_scope_key(ctx)
        entry = self._store.get(scope)
        files_map = entry["files"] if entry else {}
        if name not in files_map:
            raise ValueError(f"File {name!r} not found. Available: {list(files_map)}")
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
        if is_text:
            try:
                result["content"] = base64.b64decode(fe["data"]).decode("utf-8")
            except UnicodeDecodeError:
                result["content_base64"] = fe["data"][:200] + "..."
        else:
            result["content_base64"] = fe["data"][:200] + "..."
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
            ValueError: if the named file is not in the caller's scope (message
                lists the available upload names) or its stored data is not valid
                base64.
        """
        ctx = _current_context()
        scope = self._get_scope_key(ctx)
        entry = self._store.get(scope)
        files_map = entry["files"] if entry else {}
        if name not in files_map:
            raise ValueError(
                f"No uploaded file named {name!r}. Available uploads: "
                f"{sorted(files_map)}. Upload it via the file-upload UI first "
                "(see list_files)."
            )
        fe = files_map[name]
        try:
            data = base64.b64decode(fe["data"], validate=True)
        except Exception as exc:  # malformed stored payload
            raise ValueError(
                f"Uploaded file {name!r} could not be decoded: {exc}"
            ) from exc
        return data, fe.get("type") or "application/octet-stream"


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

    The single entry point the Graph upload tools call. Raises ``ValueError``
    with an actionable message when the feature is disabled or the file is not
    found in the caller's scope.
    """
    provider = get_upload_provider()
    if provider is None:
        raise ValueError(
            "file upload is not enabled on this server (MCP_ENABLE_FILE_UPLOAD)"
        )
    return provider.resolve(name)
