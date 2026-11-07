"""Structural validation for think-cell ``.ppttc`` documents.

A ``.ppttc`` document is a JSON array of *template entries*. Each entry has a
``template`` (a ``.pptx`` filename) and a ``data`` array of *elements*. Each
element has a ``name`` and a ``table`` -- a 2D array of *cells*, where a cell
is either ``null`` or an object carrying exactly one of ``string``,
``number``, ``percentage`` or ``date`` (optionally plus ``fill``).

``validate_ppttc_data`` checks that structure and the per-element required
fields, returning ``{"valid": bool, "errors": list[str]}``.
"""
from __future__ import annotations

import math
from typing import Any

# Keys that identify a cell's value type. Exactly one must be present.
# Mirrors the official think-cell ppttc-schema.json, which accepts
# "string", "number", "percentage" and "date" (see sample.ppttc).
CELL_TYPE_KEYS: frozenset[str] = frozenset(
    {"string", "number", "percentage", "date"}
)

# Cell type keys whose value must be numeric.
NUMERIC_CELL_KEYS: frozenset[str] = frozenset({"number", "percentage"})

# Keys allowed on a cell object in addition to the single type key.
CELL_EXTRA_KEYS: frozenset[str] = frozenset({"fill"})


def validate_ppttc_data(document: Any) -> dict[str, Any]:
    """Validate a parsed ``.ppttc`` document.

    Args:
        document: The JSON-decoded ``.ppttc`` content (expected: a list).

    Returns:
        ``{"valid": bool, "errors": list[str]}`` -- ``valid`` is True only
        when ``errors`` is empty.
    """
    errors: list[str] = []

    if not isinstance(document, list):
        return {
            "valid": False,
            "errors": [
                "The top level of a .ppttc document must be a JSON array, "
                f"got {_typename(document)}"
            ],
        }

    if len(document) == 0:
        errors.append("Document is empty: it contains no template entries")

    for index, entry in enumerate(document):
        errors.extend(_validate_template_entry(entry, f"template[{index}]"))

    return {"valid": len(errors) == 0, "errors": errors}


def _validate_template_entry(entry: Any, where: str) -> list[str]:
    """Validate a single ``{"template": ..., "data": [...]}`` entry."""
    errors: list[str] = []

    if not isinstance(entry, dict):
        return [f"{where} must be an object, got {_typename(entry)}"]

    # -- template field --------------------------------------------------
    if "template" not in entry:
        errors.append(f"{where} is missing required field 'template'")
    else:
        template = entry["template"]
        if not isinstance(template, str) or not template:
            errors.append(f"{where}.template must be a non-empty string")
        elif not template.lower().endswith(".pptx"):
            errors.append(
                f"{where}.template must reference a .pptx file, "
                f"got '{template}'"
            )

    # -- data field ------------------------------------------------------
    if "data" not in entry:
        errors.append(f"{where} is missing required field 'data'")
        return errors

    data = entry["data"]
    if not isinstance(data, list):
        errors.append(f"{where}.data must be an array, got {_typename(data)}")
        return errors

    if len(data) == 0:
        errors.append(f"{where}.data has no charts or text fields")

    for index, element in enumerate(data):
        errors.extend(_validate_element(element, f"{where}.data[{index}]"))

    return errors


def _validate_element(element: Any, where: str) -> list[str]:
    """Validate a single chart / text-field element (``name`` + ``table``)."""
    errors: list[str] = []

    if not isinstance(element, dict):
        return [f"{where} must be an object, got {_typename(element)}"]

    if "name" not in element:
        errors.append(f"{where} is missing required field 'name'")
    elif not isinstance(element["name"], str) or not element["name"]:
        errors.append(f"{where}.name must be a non-empty string")

    if "table" not in element:
        errors.append(f"{where} is missing required field 'table'")
        return errors

    table = element["table"]
    if not isinstance(table, list):
        errors.append(f"{where}.table must be an array, got {_typename(table)}")
        return errors
    if len(table) == 0:
        errors.append(f"{where}.table must contain at least one row")
        return errors

    errors.extend(_validate_table(table, f"{where}.table"))
    return errors


def _validate_table(table: list[Any], where: str) -> list[str]:
    """Validate the rows/cells of a table and check row-width consistency."""
    errors: list[str] = []
    row_width: int | None = None

    for row_index, row in enumerate(table):
        if not isinstance(row, list):
            errors.append(
                f"{where}[{row_index}] must be an array, got {_typename(row)}"
            )
            continue

        for cell_index, cell in enumerate(row):
            errors.extend(
                _validate_cell(cell, f"{where}[{row_index}][{cell_index}]")
            )

        # An empty row is a legal separator; skip it for the width check.
        if len(row) == 0:
            continue
        if row_width is None:
            row_width = len(row)
        elif len(row) != row_width:
            errors.append(
                f"{where}[{row_index}] has {len(row)} cells but earlier rows "
                f"have {row_width}; all non-empty rows must be the same width"
            )

    return errors


def _validate_cell(cell: Any, where: str) -> list[str]:
    """Validate one cell: ``null`` or ``{type_key: value[, "fill": ...]}``."""
    if cell is None:
        return []

    if not isinstance(cell, dict):
        return [
            f"{where} must be null or an object such as "
            f'{{"string": ...}}, got {_typename(cell)}'
        ]

    keys = set(cell.keys())
    type_keys = keys & CELL_TYPE_KEYS
    if len(type_keys) == 0:
        return [
            f"{where} must contain exactly one of "
            f"{sorted(CELL_TYPE_KEYS)}"
        ]
    if len(type_keys) > 1:
        return [
            f"{where} must contain exactly one type key, "
            f"found {sorted(type_keys)}"
        ]

    errors: list[str] = []
    unsupported = keys - CELL_TYPE_KEYS - CELL_EXTRA_KEYS
    if unsupported:
        errors.append(f"{where} has unsupported key(s): {sorted(unsupported)}")

    type_key = next(iter(type_keys))
    value = cell[type_key]
    if type_key in NUMERIC_CELL_KEYS:
        if not isinstance(value, (int, float)) or isinstance(value, bool):
            errors.append(
                f"{where}.{type_key} must be numeric, got {_typename(value)}"
            )
        elif not math.isfinite(value):
            errors.append(
                f"{where}.{type_key} must be a finite number; NaN/Infinity "
                f"are not valid JSON and break think-cell"
            )
    else:  # "string" or "date"
        if not isinstance(value, str):
            errors.append(
                f"{where}.{type_key} must be a string, got {_typename(value)}"
            )
        elif type_key == "date" and not _is_iso_date(value):
            # think-cell manual, Cell data types: required format "YYYY-MM-DD".
            errors.append(f"{where}.date must be an ISO date YYYY-MM-DD, got {value!r}")

    return errors


def _is_iso_date(value: str) -> bool:
    import datetime

    try:
        return len(value) == 10 and datetime.date.fromisoformat(value) is not None
    except ValueError:
        return False


def _typename(value: Any) -> str:
    """Return a JSON-flavoured type name for clearer error messages."""
    if value is None:
        return "null"
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, str):
        return "string"
    if isinstance(value, (int, float)):
        return "number"
    if isinstance(value, list):
        return "array"
    if isinstance(value, dict):
        return "object"
    return type(value).__name__
