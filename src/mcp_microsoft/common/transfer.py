from __future__ import annotations

from pathlib import Path
from typing import Any

import httpx
from fastmcp.server.context import Context

from mcp_microsoft.graph import get_transfer_http_client


async def upload_large_file_via_session(
    upload_url: str,
    file_path: Path,
    total_size: int,
    ctx: Context | None = None,
    *,
    chunk_size: int = 10 * 1024 * 1024,
) -> dict[str, Any]:
    """Upload a large file in chunks using a Microsoft Graph upload session."""
    result: dict[str, Any] = {}
    shared_client = get_transfer_http_client()

    if ctx is not None:
        await ctx.report_progress(progress=0, total=total_size)

    async def _send_chunks(client: httpx.AsyncClient) -> dict[str, Any]:
        offset = 0
        with file_path.open("rb") as stream:
            while offset < total_size:
                chunk = stream.read(chunk_size)
                if not chunk:
                    break

                end = offset + len(chunk)
                content_range = f"bytes {offset}-{end - 1}/{total_size}"

                response = await client.put(
                    upload_url,
                    content=chunk,
                    headers={
                        "Content-Range": content_range,
                        "Content-Length": str(len(chunk)),
                    },
                )
                response.raise_for_status()

                local_result = response.json() if response.status_code in (200, 201) else {}

                if ctx is not None:
                    await ctx.report_progress(progress=end, total=total_size)

                offset = end

        return local_result

    if shared_client is None:
        async with httpx.AsyncClient(timeout=120.0, follow_redirects=True) as ephemeral_client:
            result = await _send_chunks(ephemeral_client)
    else:
        result = await _send_chunks(shared_client)

    return result
