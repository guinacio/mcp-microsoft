from __future__ import annotations

from typing import Optional, Union

from mcp_microsoft.common.formatting import format_datetime_display
from mcp_microsoft.graph_types import GraphRecipient
from mcp_microsoft.models import Address


def parse_recipients(
    value: Optional[Union[str, list[str]]],
) -> list[dict]:
    """Convert text/list recipients into Graph recipient objects."""
    if not value:
        return []
    if isinstance(value, str):
        addresses = [addr.strip() for addr in value.split(",") if addr.strip()]
    else:
        addresses = [addr.strip() for addr in value if addr.strip()]
    return [{"emailAddress": {"address": addr}} for addr in addresses]


def recipient_values(recipients: list[GraphRecipient]) -> list[Address]:
    """Normalize Graph recipients into simple address models."""
    values: list[Address] = []
    for recipient in recipients or []:
        email_address = recipient.email_address
        values.append(Address(name=email_address.name, address=email_address.address))
    return values


def format_mail_datetime(iso: str | None) -> str:
    """Mail-friendly wrapper around the shared datetime formatter."""
    return format_datetime_display(iso)
