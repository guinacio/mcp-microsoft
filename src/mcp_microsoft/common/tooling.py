from __future__ import annotations

import copy
import inspect
import re
from collections.abc import Callable
from typing import Annotated, Any, get_type_hints

from mcp.types import ToolAnnotations
from pydantic import BaseModel
from pydantic_core import PydanticUndefined

READ_ONLY_TOOL = ToolAnnotations(read_only_hint=True, open_world_hint=True)
WRITE_TOOL = ToolAnnotations(destructive_hint=False, open_world_hint=True)
IDEMPOTENT_WRITE_TOOL = ToolAnnotations(
    destructive_hint=False,
    idempotent_hint=True,
    open_world_hint=True,
)
DESTRUCTIVE_TOOL = ToolAnnotations(destructive_hint=True, open_world_hint=True)

LOCAL_READ_TOOL = ToolAnnotations(
    read_only_hint=True,
    idempotent_hint=True,
    open_world_hint=False,
)
LOCAL_WRITE_TOOL = ToolAnnotations(destructive_hint=False, open_world_hint=False)
LOCAL_IDEMPOTENT_TOOL = ToolAnnotations(
    destructive_hint=False,
    idempotent_hint=True,
    open_world_hint=False,
)
LOCAL_DESTRUCTIVE_TOOL = ToolAnnotations(destructive_hint=True, open_world_hint=False)

_TITLE_OVERRIDES = {
    "id": "ID",
    "onedrive": "OneDrive",
    "mcp": "MCP",
    "ms": "MS",
    "rsvp": "RSVP",
    "sharepoint": "SharePoint",
    "url": "URL",
}


def humanize_tool_title(name: str) -> str:
    """Convert a snake_case tool name into a display title for MCP clients."""
    words = []
    for part in name.split("_"):
        if not part:
            continue
        words.append(_TITLE_OVERRIDES.get(part.lower(), part.capitalize()))
    return " ".join(words)


_ARG_LINE = re.compile(r"^(\s*)([A-Za-z_][A-Za-z0-9_]*)\s*:\s*(.*)$")


def _docstring_arg_descriptions(doc: str | None) -> dict[str, str]:
    """Parse the Google-style ``Args:`` section of a tool docstring."""
    if not doc:
        return {}
    descriptions: dict[str, str] = {}
    in_args = False
    arg_indent: int | None = None
    current: str | None = None
    for line in inspect.cleandoc(doc).splitlines():
        stripped = line.strip()
        if not in_args:
            in_args = stripped == "Args:"
            continue
        if not stripped:
            current = None
            continue
        indent = len(line) - len(line.lstrip())
        if indent == 0:
            break  # next section (Returns:, Raises:, ...)
        match = _ARG_LINE.match(line)
        if match and (arg_indent is None or indent == arg_indent):
            arg_indent = indent
            current = match.group(2)
            descriptions[current] = match.group(3).strip()
        elif current is not None:
            descriptions[current] = f"{descriptions[current]} {stripped}".strip()
    return descriptions


def _strip_docstring_args(doc: str | None) -> str | None:
    """Drop the ``Args:`` section: each argument now carries its own description."""
    if not doc:
        return doc
    kept: list[str] = []
    skipping = False
    for line in inspect.cleandoc(doc).splitlines():
        if line.strip() == "Args:":
            skipping = True
            continue
        if skipping and line and not line[0].isspace():
            skipping = False
        if not skipping:
            kept.append(line)
    return "\n".join(kept).strip()


def _request_model_param(fn: Callable[..., Any]) -> tuple[str, type[BaseModel]] | None:
    """Return the ``(name, model)`` of a tool's request-model parameter, if any."""
    try:
        hints = get_type_hints(fn)
    except Exception:  # unresolvable annotations: register the function as-is
        return None
    for param in inspect.signature(fn).parameters.values():
        hint = hints.get(param.name)
        if isinstance(hint, type) and issubclass(hint, BaseModel):
            return param.name, hint
    return None


def flatten_request_model(fn: Callable[..., Any]) -> Callable[..., Any]:
    """Expose a tool's request-model fields as top-level tool arguments.

    Tool implementations take one validated ``ToolRequestModel`` (plus an
    optional ``ctx``). Publishing that model as a single ``params`` argument
    makes every client nest its arguments one level down, so the function
    registered with FastMCP takes each field as its own keyword argument and
    builds the model before calling the implementation. Validation, coercion
    and model validators are unchanged because the same model still runs.
    """
    found = _request_model_param(fn)
    if found is None:
        return fn
    model_param, model = found
    descriptions = _docstring_arg_descriptions(fn.__doc__)
    hints = get_type_hints(fn)

    params: list[inspect.Parameter] = []
    annotations: dict[str, Any] = {}
    for field_name, field in model.model_fields.items():
        arg_name = field.alias or field_name
        description = (
            field.description or descriptions.get(field_name) or descriptions.get(arg_name)
        )
        default = (
            inspect.Parameter.empty
            if field.is_required()
            else field.get_default(call_default_factory=True)
        )
        # Keep the field's constraints, deprecation and examples on the
        # argument; the default moves to the parameter and the alias is
        # already the argument name.
        arg_field = copy.copy(field)
        arg_field.default = PydanticUndefined
        arg_field.default_factory = None
        arg_field.alias = arg_field.validation_alias = arg_field.serialization_alias = None
        arg_field.description = description
        annotation = Annotated[field.annotation, arg_field]
        params.append(
            inspect.Parameter(
                arg_name, inspect.Parameter.KEYWORD_ONLY, default=default, annotation=annotation
            )
        )
        annotations[arg_name] = annotation

    passthrough: list[str] = []
    for param in inspect.signature(fn).parameters.values():
        if param.name == model_param:
            continue
        annotation = hints.get(param.name, param.annotation)
        passthrough.append(param.name)
        params.append(
            param.replace(kind=inspect.Parameter.KEYWORD_ONLY, annotation=annotation)
        )
        annotations[param.name] = annotation
    annotations["return"] = hints.get("return", inspect.Signature.empty)

    async def wrapper(**kwargs: Any) -> Any:
        extra = {name: kwargs.pop(name) for name in passthrough if name in kwargs}
        result = fn(model.model_validate(kwargs), **extra)
        if inspect.isawaitable(result):
            result = await result
        return result

    wrapper.__name__ = fn.__name__
    wrapper.__qualname__ = fn.__qualname__
    wrapper.__module__ = fn.__module__
    wrapper.__doc__ = _strip_docstring_args(fn.__doc__)
    wrapper.__annotations__ = annotations
    wrapper.__signature__ = inspect.Signature(  # type: ignore[attr-defined]
        params, return_annotation=annotations["return"]
    )
    return wrapper


def register_tool(
    server: Any,
    fn: Callable[..., Any],
    *,
    annotations: ToolAnnotations | dict[str, Any] | None = None,
    name: str | None = None,
    title: str | None = None,
    **kwargs: Any,
) -> Any:
    """Register a tool with a stable human-readable title and flat arguments."""
    tool_name = name or fn.__name__
    tool_title = title or humanize_tool_title(tool_name)
    return server.tool(
        name=tool_name,
        title=tool_title,
        annotations=annotations,
        **kwargs,
    )(flatten_request_model(fn))
