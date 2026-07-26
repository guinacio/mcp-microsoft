"""Tests for the CLI logging configuration (server._configure_logging).

Without configuration, Python's root logger sits at WARNING with only the
last-resort handler — which silently swallows the http-mode audit trail
(AuditLoggingMiddleware logs at INFO). These tests pin the level-resolution
rules; each one restores the root logger it mutates.
"""

from __future__ import annotations

import logging

import pytest

from mcp_microsoft.config import AppConfig
from mcp_microsoft.server import _configure_logging


@pytest.fixture()
def restore_root_logger():
    root = logging.getLogger()
    old_level = root.level
    old_handlers = root.handlers[:]
    yield
    root.handlers[:] = old_handlers
    root.setLevel(old_level)


def test_http_mode_defaults_to_info_so_audit_log_is_emitted(
    monkeypatch: pytest.MonkeyPatch, restore_root_logger
) -> None:
    monkeypatch.delenv("MCP_LOG_LEVEL", raising=False)
    _configure_logging(AppConfig(transport="http"))
    root = logging.getLogger()
    assert root.level == logging.INFO
    # The audit logger (mcp_microsoft.middleware) inherits and emits INFO.
    assert logging.getLogger("mcp_microsoft.middleware").isEnabledFor(logging.INFO)


def test_stdio_mode_defaults_to_warning(
    monkeypatch: pytest.MonkeyPatch, restore_root_logger
) -> None:
    monkeypatch.delenv("MCP_LOG_LEVEL", raising=False)
    _configure_logging(AppConfig(transport="stdio"))
    assert logging.getLogger().level == logging.WARNING


def test_mcp_log_level_env_overrides_either_default(
    monkeypatch: pytest.MonkeyPatch, restore_root_logger
) -> None:
    monkeypatch.setenv("MCP_LOG_LEVEL", "debug")
    _configure_logging(AppConfig(transport="stdio"))
    assert logging.getLogger().level == logging.DEBUG

    monkeypatch.setenv("MCP_LOG_LEVEL", "ERROR")
    _configure_logging(AppConfig(transport="http"))
    assert logging.getLogger().level == logging.ERROR


def test_invalid_mcp_log_level_falls_back_to_transport_default(
    monkeypatch: pytest.MonkeyPatch, restore_root_logger
) -> None:
    monkeypatch.setenv("MCP_LOG_LEVEL", "bogus")
    _configure_logging(AppConfig(transport="http"))
    assert logging.getLogger().level == logging.INFO
