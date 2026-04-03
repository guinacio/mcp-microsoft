"""Shared request model base class for mcp-microsoft tools.

Provides ToolRequestModel — a Pydantic BaseModel subclass that:
  - Strips leading/trailing whitespace from all string fields.
  - Accepts a raw JSON string in place of a dict (useful when an LLM
    serialises the entire payload as a string).
  - Converts camelCase field names to snake_case so callers can supply
    either convention.
  - Coerces comma-separated strings into list[str] fields, and parses
    JSON text for list[...] / dict / nested model fields.
  - Exposes a ``from_tool_params`` class-method as a convenience wrapper
    around ``model_validate``.
"""

from __future__ import annotations

import json
import re
from typing import Any, get_args, get_origin

from pydantic import BaseModel, ConfigDict, ValidationInfo, field_validator, model_validator


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------


def _camel_to_snake(key: str) -> str:
    """Convert a camelCase or kebab-case key to snake_case."""
    if not key:
        return key
    key = key.replace("-", "_")
    return re.sub(r"(?<!^)(?=[A-Z])", "_", key).lower()


def _annotation_contains(annotation: Any, target: type[Any]) -> bool:
    """Return True if *target* appears anywhere in the type annotation tree."""
    origin = get_origin(annotation)
    if origin is None:
        return annotation is target
    if origin is target:
        return True
    return any(_annotation_contains(arg, target) for arg in get_args(annotation))


def _annotation_contains_basemodel(annotation: Any) -> bool:
    """Return True if any leaf type in the annotation is a Pydantic BaseModel."""
    origin = get_origin(annotation)
    if origin is None:
        return isinstance(annotation, type) and issubclass(annotation, BaseModel)
    return any(_annotation_contains_basemodel(arg) for arg in get_args(annotation))


# ---------------------------------------------------------------------------
# Base class
# ---------------------------------------------------------------------------


class ToolRequestModel(BaseModel):
    """Base input model for mcp-microsoft tool functions.

    Usage::

        class SendEmailInput(ToolRequestModel):
            to: str | list[str]
            subject: str
            body: str
            ...

        # In the tool function:
        params = SendEmailInput.from_tool_params(raw_input)
    """

    model_config = ConfigDict(
        populate_by_name=True,
        str_strip_whitespace=True,
    )

    # ------------------------------------------------------------------
    # Validators
    # ------------------------------------------------------------------

    @model_validator(mode="before")
    @classmethod
    def _normalize_payload_shape(cls, data: Any) -> Any:
        """Accept a raw JSON string or camelCase dict in addition to a plain dict."""
        # If the entire payload was serialised as a JSON string, parse it first.
        if isinstance(data, str):
            raw = data.strip()
            if not raw:
                raise ValueError("Tool request payload cannot be an empty string.")
            try:
                data = json.loads(raw)
            except json.JSONDecodeError as exc:
                raise ValueError(
                    "Tool request payload must be a JSON object, not a plain string."
                ) from exc

        # Normalise camelCase / kebab-case keys to snake_case.
        if isinstance(data, dict):
            return {
                (_camel_to_snake(k) if isinstance(k, str) else k): v
                for k, v in data.items()
            }

        return data

    @field_validator("*", mode="before")
    @classmethod
    def _coerce_json_text_for_collection_fields(cls, value: Any, info: ValidationInfo) -> Any:
        """Parse JSON text or split comma-separated strings for list/dict/model fields."""
        if not isinstance(value, str):
            return value

        field_name = info.field_name
        if not field_name:
            return value

        field = cls.model_fields.get(field_name)
        if field is None:
            return value

        expects_list = _annotation_contains(field.annotation, list)
        expects_dict = _annotation_contains(field.annotation, dict)
        expects_model = _annotation_contains_basemodel(field.annotation)

        if not (expects_list or expects_dict or expects_model):
            return value

        raw = value.strip()
        if not raw:
            return value

        # If it looks like a JSON array or object, parse it.
        if raw[0] in ("[", "{"):
            try:
                return json.loads(raw)
            except json.JSONDecodeError as exc:
                expected = "JSON array" if expects_list else "JSON object"
                raise ValueError(
                    f"Field '{field_name}' expects {expected} text when passed as a string."
                ) from exc

        # Comma-separated plain text → list[str] convenience coercion.
        if expects_list:
            return [item.strip() for item in raw.split(",") if item.strip()]

        if expects_dict or expects_model:
            raise ValueError(
                f"Field '{field_name}' expects a JSON object; plain text is not supported."
            )

        return value

    # ------------------------------------------------------------------
    # Convenience factory
    # ------------------------------------------------------------------

    @classmethod
    def from_tool_params(cls, data: Any) -> "ToolRequestModel":
        """Validate and construct the model from raw tool parameters.

        Equivalent to ``cls.model_validate(data)`` but provides a stable,
        intention-revealing name for call sites.
        """
        return cls.model_validate(data)
