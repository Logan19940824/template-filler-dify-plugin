from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

TOKEN_RE = re.compile(r"\{\{\s*([^{}]+?)\s*\}\}")


def parse_values(raw: Any) -> dict[str, Any]:
    if isinstance(raw, dict):
        return raw
    try:
        result = json.loads(raw)
    except (TypeError, json.JSONDecodeError) as exc:
        raise ValueError("values must be a valid JSON object") from exc
    if not isinstance(result, dict):
        raise ValueError("values must be a JSON object")
    return result


def string_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    return str(value)


def _segments(expression: str) -> list[str]:
    return [part for part in expression.strip().split(".") if part]


def placeholder_names(text: str) -> list[str]:
    return list(dict.fromkeys(match.group(1).strip() for match in TOKEN_RE.finditer(text)))


def array_fields(text: str) -> list[str]:
    result = []
    for match in TOKEN_RE.finditer(text):
        path = []
        for part in _segments(match.group(1)):
            if part.endswith("[]"):
                path.append(part[:-2])
                result.append(".".join(path))
            else:
                path.append(part)
    return list(dict.fromkeys(result))


def value_defaults(names: list[str]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for name in names:
        key = name.split("[].", 1)[0] if "[]." in name else name
        result.setdefault(key, [] if "[]." in name else "")
    return result


def expand_contexts(values: dict[str, Any], paths: list[str]) -> list[dict[str, Any]]:
    contexts = [{}]
    for path in sorted(paths, key=lambda value: (value.count("."), value)):
        parent_path, _, field = path.rpartition(".")
        expanded = []
        for context in contexts:
            parent = values if not parent_path else context.get(parent_path, {})
            items = parent.get(field, []) if isinstance(parent, dict) else []
            if not isinstance(items, list) or not all(isinstance(item, dict) for item in items):
                raise ValueError(f"array value '{path}' must be a JSON array of objects")
            for index, item in enumerate(items, 1):
                child = dict(context)
                child[path] = item
                child[f"__index__:{path}"] = index
                expanded.append(child)
        contexts = expanded
    return contexts


def render_context(text: str, values: dict[str, Any], context: dict[str, Any]) -> str:
    def render(match):
        expression = match.group(1).strip()
        current: Any = values
        path = []
        last_array = None
        for part in _segments(expression):
            if part.endswith("[]"):
                path.append(part[:-2])
                last_array = ".".join(path)
                current = context.get(last_array, {})
            elif part == "_index":
                current = context.get(f"__index__:{last_array}", match.group(0))
            else:
                path.append(part)
                current = current.get(part, match.group(0)) if isinstance(current, dict) else match.group(0)
        return string_value(current)
    return TOKEN_RE.sub(render, text)


def replace_text(text: str, values: dict[str, Any]) -> str:
    return render_context(text, values, {})


def replace_item_text(text: str, values: dict[str, Any], item: dict[str, Any]) -> str:
    return render_context(text, values, item)


def ensure_extension(path: Path, suffix: str) -> None:
    if path.suffix.lower() != suffix:
        raise ValueError(f"expected a {suffix} file")
