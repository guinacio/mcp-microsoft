# syntax=docker/dockerfile:1
#
# mcp-microsoft — container image for the multi-user Streamable HTTP server.
#
# stdio mode (the default MCPB / Claude Desktop experience) is a local
# subprocess launched by the client and has no reason to run in a container.
# This image targets *http* transport only: it sets MCP_TRANSPORT=http and
# exposes the Streamable HTTP endpoint on :8000/mcp.
#
# You still must supply the auth/base-url configuration below at `docker run`
# / compose time — see .env.template's "Remote server (http) mode" section
# and docs/azure-setup.md.
#
# Pull:   docker pull ghcr.io/guinacio/mcp-microsoft:latest
#         (published for linux/amd64 + linux/arm64 on every GitHub release
#         by .github/workflows/release.yml)
# Build:  docker build -t mcp-microsoft:0.10.0 .
# Run:    docker run --rm -p 8000:8000 --env-file .env ghcr.io/guinacio/mcp-microsoft:latest
# Or:     docker compose up -d   (see docker-compose.yml)

# ---------------------------------------------------------------------------
# Stage 1: builder — resolve and install dependencies with uv.
# Kept separate from the runtime stage so build-only tooling (the uv binary
# itself, pip caches, etc.) never ends up in the image that actually runs.
# ---------------------------------------------------------------------------
FROM python:3.12-slim AS builder

# Apply all pending OS security patches before anything else.
# Addresses Debian-layer CVEs reported by image scans (perl, glibc, sqlite3,
# util-linux, etc.) that are not fixable via Python package upgrades alone.
RUN apt-get update \
    && apt-get upgrade -y --no-install-recommends \
    && rm -rf /var/lib/apt/lists/*

# Official static uv binary — avoids a `pip install uv` bootstrap and keeps
# uv's own version pinned to upstream's "latest" tag rather than PyPI's.
COPY --from=ghcr.io/astral-sh/uv:latest /uv /usr/local/bin/uv

ENV \
    # Install the project venv at a fixed, known path so the runtime stage
    # can copy it verbatim.
    UV_PROJECT_ENVIRONMENT=/app/.venv \
    # Copy packages into the venv instead of hardlinking — hardlinks don't
    # survive the COPY --from= into the runtime stage below.
    UV_LINK_MODE=copy \
    # This is a container build: never try to download/manage a Python
    # interpreter, just use the one already in the base image.
    UV_PYTHON_DOWNLOADS=never

WORKDIR /app

# Copy only dependency-resolution inputs first so this (slow) layer stays
# cached across source-only changes. README.md is copied too because
# pyproject.toml declares it as the package readme and hatchling reads it
# during the build.
COPY pyproject.toml uv.lock README.md ./
RUN uv sync --frozen --no-dev --no-install-project

# Now copy the actual source and install the project itself.
COPY src/ ./src/
RUN uv sync --frozen --no-dev

# ---------------------------------------------------------------------------
# Stage 2: runtime — slim image with just the venv, source, and a non-root user.
# ---------------------------------------------------------------------------
FROM python:3.12-slim AS runtime

# Apply OS security patches in the runtime image — this is the layer that
# gets scanned and shipped, so patching here is what clears the findings.
RUN apt-get update \
    && apt-get upgrade -y --no-install-recommends \
    && rm -rf /var/lib/apt/lists/*

COPY --from=ghcr.io/astral-sh/uv:latest /uv /usr/local/bin/uv

ENV \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    UV_PROJECT_ENVIRONMENT=/app/.venv \
    # The venv baked into the image during the builder stage is already
    # exactly what uv.lock specifies — never let `uv run` try to re-resolve
    # or hit the network for it at container start.
    UV_NO_SYNC=1 \
    PATH="/app/.venv/bin:$PATH"

WORKDIR /app
COPY --from=builder /app /app

# Non-root: the server never needs root, and local-disk tools are already
# disabled in http mode (see README's "Remote server" section), so there's
# no legitimate reason for this process to write outside its venv/tmp.
RUN groupadd --system mcp \
    && useradd --system --gid mcp --home-dir /app --no-create-home mcp \
    && chown -R mcp:mcp /app
USER mcp

# Streamable HTTP mode, bound to all interfaces *inside* the container — the
# host/reverse-proxy in front of it controls what's actually reachable from
# outside. Auth/base-url values are intentionally NOT set here: they're
# secrets and deployment-specific, and must come from --env-file/-e at run
# time (see .env.template and docker-compose.yml).
ENV MCP_TRANSPORT=http \
    MCP_HTTP_HOST=0.0.0.0 \
    MCP_HTTP_PORT=8000

EXPOSE 8000

# Unauthenticated GET /health (see server.py's _register_health_route) — no
# bearer token needed, safe for a container-level healthcheck or LB probe.
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD python -c "import urllib.request,sys; sys.exit(0 if urllib.request.urlopen('http://127.0.0.1:8000/health', timeout=3).status == 200 else 1)"

CMD ["uv", "run", "mcp-microsoft"]
