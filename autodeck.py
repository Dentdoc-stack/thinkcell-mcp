"""Zero-setup think-cell deck builder.

The rest of this server can only fill think-cell charts that already exist
**and are named** inside a user-supplied ``.pptx`` template. Naming those
elements is a manual step in PowerPoint. This module removes that step.

think-cell ships its own official automation template. A copy lives in this
project at ``templates/thinkcell_auto.pptx`` and already contains five named
think-cell elements:

==================  ==========================================================
Element name        Purpose
==================  ==========================================================
``SlideTitle``      Slide title text field.
``LeftChartTitle``  Title text field above the left chart.
``RightChartTitle`` Title text field above the right chart.
``LeftChart``       Left chart -- date or category axis, ``percentage`` series.
``RightChart``      Right chart -- date or category axis, ``number`` series.
==================  ==========================================================

A chart's ``categories`` may be either ISO dates (``YYYY-MM-DD`` -- a date
axis) or plain non-empty strings (a category axis); the axis kind is inferred
per chart from whether *every* category matches the ISO date pattern.

Because those names are fixed and known, the user never has to open
PowerPoint: they just provide data. :func:`build_auto_deck` turns a list of
slide definitions into a ``.ppttc`` document, writing **one ``.ppttc`` entry
per slide** (each top-level entry produces exactly one slide).

The ``.ppttc`` JSON is assembled directly here -- not via the ``thinkcell``
library -- so cells can use the full think-cell cell vocabulary
(``date``, ``percentage``, ``fill``), which that library does not expose.
"""
from __future__ import annotations

import json
import math
import os
import re
import shutil
from pathlib import Path
from typing import Any

from palettes import semantic_fill

# The official think-cell automation template, kept under templates/. Resolved
# relative to THIS file so it works regardless of the process working dir.
MODULE_DIR = Path(__file__).resolve().parent
AUTO_TEMPLATE_PATH = MODULE_DIR / "templates" / "thinkcell_auto.pptx"
OUTPUT_DIR = MODULE_DIR / "output"

# think-cell installs an identical automation template under its program
# folder. We copy it from there on first use rather than redistributing
# think-cell's own file. Override the install dir with THINKCELL_DIR.
THINKCELL_DIR = Path(
    os.environ.get("THINKCELL_DIR", r"C:\Program Files (x86)\think-cell")
)
_INSTALL_TEMPLATE = THINKCELL_DIR / "ppttc" / "template.pptx"


def ensure_auto_template() -> bool:
    """Make sure the automation template exists locally; copy it if missing.

    The template is think-cell's own property and is intentionally NOT
    shipped with this project. The first time it is needed, it is copied from
    the local think-cell installation (``<THINKCELL_DIR>/ppttc/template.pptx``).

    Returns:
        True if ``AUTO_TEMPLATE_PATH`` exists (or was successfully copied).
    """
    if AUTO_TEMPLATE_PATH.is_file():
        return True
    if _INSTALL_TEMPLATE.is_file():
        try:
            AUTO_TEMPLATE_PATH.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(_INSTALL_TEMPLATE, AUTO_TEMPLATE_PATH)
        except OSError:
            return False
        if not AUTO_TEMPLATE_PATH.is_file():
            return False
        # De-brand on first copy: drop think-cell's logo and apply the
        # default professional palette. Best-effort -- the template is still
        # usable (just with think-cell's own styling) if this fails.
        try:
            from branding import apply_branding

            apply_branding(AUTO_TEMPLATE_PATH)
        except Exception:  # noqa: BLE001 - branding must never block usage
            pass
        return True
    return False

# Named think-cell elements present in the bundled template.
SLIDE_TITLE_NAME = "SlideTitle"
LEFT_TITLE_NAME = "LeftChartTitle"
RIGHT_TITLE_NAME = "RightChartTitle"
LEFT_CHART_NAME = "LeftChart"
RIGHT_CHART_NAME = "RightChart"

# ISO-ish date pattern (YYYY-MM-DD). A chart whose categories ALL match this
# gets a date axis; otherwise the categories are treated as a string axis.
_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")

# Hex colour accepted on a series' optional "fill".
_FILL_RE = re.compile(r"^#[0-9A-Fa-f]{6}$")


def _safe_filename(name: str, fallback: str) -> str:
    """Turn an arbitrary title into a safe filesystem stem."""
    cleaned = "".join(
        char for char in name if char.isalnum() or char in (" ", "-", "_")
    ).strip()
    return cleaned.replace(" ", "_") or fallback


def _text_element(name: str, text: str) -> dict[str, Any]:
    """Build a single-cell text-field element for the ``.ppttc`` data array."""
    return {"name": name, "table": [[{"string": str(text)}]]}


def _validate_chart(
    chart: Any, where: str, value_key: str
) -> tuple[list[str], list[list[Any]] | None]:
    """Validate one chart definition and build its ``.ppttc`` table.

    Each category must be a non-empty string. If *every* category matches the
    ISO date pattern (``YYYY-MM-DD``) the chart gets a date axis (``date``
    header cells); otherwise the categories form a string axis (``string``
    header cells).

    Args:
        chart: The chart object: ``{"categories": [...], "series": [...]}``.
        where: Location label for error messages.
        value_key: ``"percentage"`` (left chart) or ``"number"`` (right chart).

    Returns:
        ``(errors, table)``. ``table`` is None when ``errors`` is non-empty.
    """
    errors: list[str] = []

    if not isinstance(chart, dict):
        return [f"{where} must be an object"], None

    categories = chart.get("categories")
    if not isinstance(categories, list) or len(categories) == 0:
        errors.append(
            f"{where}.categories must be a non-empty list of strings "
            "(ISO dates 'YYYY-MM-DD' for a date axis, or plain labels)"
        )
        categories = []
    else:
        for i, category in enumerate(categories):
            if not isinstance(category, str) or not category:
                errors.append(
                    f"{where}.categories[{i}] must be a non-empty string, "
                    f"got {category!r}"
                )

    series = chart.get("series")
    if not isinstance(series, list) or len(series) == 0:
        errors.append(f"{where}.series must be a non-empty list of series")
        series = []

    width = len(categories)
    parsed_series: list[list[Any]] = []
    for s_index, entry in enumerate(series):
        s_where = f"{where}.series[{s_index}]"
        if not isinstance(entry, dict):
            errors.append(f"{s_where} must be an object")
            continue

        s_name = entry.get("name")
        if not isinstance(s_name, str) or not s_name:
            errors.append(f"{s_where}.name must be a non-empty string")
            s_name = ""

        values = entry.get("values")
        if not isinstance(values, list):
            errors.append(f"{s_where}.values must be a list of numbers")
            continue
        if width and len(values) != width:
            errors.append(
                f"{s_where}.values has {len(values)} value(s) but there are "
                f"{width} categor(ies); they must match"
            )

        fill = entry.get("fill")
        if fill is not None and (
            not isinstance(fill, str) or not _FILL_RE.match(fill)
        ):
            errors.append(
                f"{s_where}.fill must be a hex colour like '#ff0000', "
                f"got {fill!r}"
            )

        # Optional semantic auto-colouring: colour each bar by the sign of its
        # value (gain teal / loss red / zero gray). This is the RAG-by-delta
        # idea applied to think-cell -- ideal for variance, bridge and delta
        # charts. An explicit `fill` always wins over it. `reverse` flips the
        # mapping for "down is good" metrics (cost, churn, defects).
        semantic = entry.get("semantic", False)
        reverse = entry.get("reverse", False)
        if not isinstance(semantic, bool):
            errors.append(f"{s_where}.semantic must be true or false")
            semantic = False
        if not isinstance(reverse, bool):
            errors.append(f"{s_where}.reverse must be true or false")
            reverse = False

        row: list[Any] = [{"string": s_name}]
        for v_index, value in enumerate(values):
            if not isinstance(value, (int, float)) or isinstance(value, bool):
                errors.append(
                    f"{s_where}.values[{v_index}] must be a number, "
                    f"got {value!r}"
                )
                continue
            if not math.isfinite(value):
                errors.append(
                    f"{s_where}.values[{v_index}] must be a finite number "
                    f"(NaN/Infinity break think-cell's .ppttc), got {value!r}"
                )
                continue
            cell: dict[str, Any] = {value_key: value}
            if isinstance(fill, str) and _FILL_RE.match(fill):
                cell["fill"] = fill
            elif semantic:
                cell["fill"] = semantic_fill(value, reverse=reverse)
            row.append(cell)
        parsed_series.append(row)

    if errors:
        return errors, None

    # Date axis only when EVERY category is an ISO date; else a string axis.
    cell_key = "date" if all(_DATE_RE.match(c) for c in categories) else "string"
    header: list[Any] = [None] + [{cell_key: c} for c in categories]
    table: list[list[Any]] = [header]
    table.extend(parsed_series)
    return [], table


def _chart_warnings(chart: Any, where: str, value_key: str) -> list[str]:
    """Non-fatal sanity warnings for two documented silent-failure gotchas.

    1. Single-category axis -- a chart with one category renders with little or
       nothing visible. Provide at least two categories.
    2. Decimal-scale percentages -- the left chart uses ``percentage`` cells,
       where think-cell expects whole points (``21.9`` for 21.9%). Passing
       fractions (``0.219``) is the most common silent mistake and yields a
       chart that looks empty/flat.
    """
    warnings: list[str] = []
    if not isinstance(chart, dict):
        return warnings

    categories = chart.get("categories")
    if isinstance(categories, list) and len(categories) == 1:
        warnings.append(
            f"{where} has a single category: the chart will render with no "
            f"meaningful bars. Provide at least two categories."
        )

    if value_key == "percentage":
        magnitudes = [
            abs(value)
            for series in (chart.get("series") or [])
            if isinstance(series, dict)
            for value in (series.get("values") or [])
            if isinstance(value, (int, float)) and not isinstance(value, bool)
        ]
        nonzero = [m for m in magnitudes if m > 0]
        if nonzero and all(m < 1 for m in nonzero):
            warnings.append(
                f"{where} percentage values all look like decimals "
                f"(every |value| < 1). think-cell expects whole points "
                f"(e.g. 21.9 for 21.9%), so multiply by 100 if these are "
                f"fractions."
            )

    return warnings


def build_auto_deck(
    slides: list[dict[str, Any]],
    output_name: str = "auto_deck",
) -> dict[str, Any]:
    """Build a ``.ppttc`` document from slide definitions, no template setup.

    Each slide is rendered onto the bundled ``thinkcell_auto.pptx`` template
    and becomes one ``.ppttc`` entry (one PowerPoint slide).

    Args:
        slides: Non-empty list of slide objects. See :func:`create_auto_deck`
            in ``server.py`` for the full slide schema.
        output_name: Filename stem for the generated ``.ppttc`` file.

    Returns:
        ``{"success": bool, "ppttc_path": str | None, "slide_count": int,
        "template": str, "errors": list[str], "warnings": list[str]}``.
        ``warnings`` is non-fatal advice (e.g. decimal-scale percentages or a
        single-category date axis) -- the build still succeeds.
    """
    if not isinstance(slides, list) or len(slides) == 0:
        return {
            "success": False,
            "ppttc_path": None,
            "slide_count": 0,
            "template": str(AUTO_TEMPLATE_PATH),
            "errors": ["`slides` must be a non-empty list of slide objects"],
            "warnings": [],
        }

    template = str(AUTO_TEMPLATE_PATH)
    errors: list[str] = []
    warnings: list[str] = []
    document: list[dict[str, Any]] = []

    for index, slide in enumerate(slides):
        where = f"slides[{index}]"
        if not isinstance(slide, dict):
            errors.append(f"{where} must be an object")
            continue

        data: list[dict[str, Any]] = []

        slide_title = slide.get("slide_title")
        if slide_title is not None:
            if not isinstance(slide_title, str):
                errors.append(f"{where}.slide_title must be a string")
            elif slide_title:
                data.append(_text_element(SLIDE_TITLE_NAME, slide_title))

        for side, title_key, title_name, chart_key, chart_name, value_key in (
            ("left", "left_title", LEFT_TITLE_NAME,
             "left_chart", LEFT_CHART_NAME, "percentage"),
            ("right", "right_title", RIGHT_TITLE_NAME,
             "right_chart", RIGHT_CHART_NAME, "number"),
        ):
            title = slide.get(title_key)
            if title is not None:
                if not isinstance(title, str):
                    errors.append(f"{where}.{title_key} must be a string")
                elif title:
                    data.append(_text_element(title_name, title))

            chart = slide.get(chart_key)
            if chart is None:
                continue
            warnings.extend(
                _chart_warnings(chart, f"{where}.{chart_key}", value_key)
            )
            chart_errors, table = _validate_chart(
                chart, f"{where}.{chart_key}", value_key
            )
            if chart_errors:
                errors.extend(chart_errors)
            elif table is not None:
                data.append({"name": chart_name, "table": table})

        if not data:
            errors.append(
                f"{where} has no content: provide at least one of "
                "slide_title, left_title/left_chart, right_title/right_chart"
            )
            continue

        document.append({"template": template, "data": data})

    if errors:
        return {
            "success": False,
            "ppttc_path": None,
            "slide_count": 0,
            "template": template,
            "errors": errors,
            "warnings": warnings,
        }

    # Input is valid -- now make sure the template file is actually available
    # (copied from the local think-cell install on first use).
    if not ensure_auto_template():
        return {
            "success": False,
            "ppttc_path": None,
            "slide_count": 0,
            "template": template,
            "errors": [
                "Could not obtain the think-cell automation template. It is "
                f"copied on first use from your think-cell install at "
                f"'{_INSTALL_TEMPLATE}'. Install think-cell, set the "
                "THINKCELL_DIR environment variable if it lives elsewhere, "
                f"or place the template manually at '{AUTO_TEMPLATE_PATH}'."
            ],
            "warnings": warnings,
        }

    out_path = OUTPUT_DIR / f"{_safe_filename(output_name, 'auto_deck')}.ppttc"
    try:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        out_path.write_text(
            json.dumps(document, indent=2), encoding="utf-8"
        )
    except OSError as exc:
        return {
            "success": False,
            "ppttc_path": None,
            "slide_count": 0,
            "template": template,
            "errors": [f"Could not write the .ppttc file: {exc}"],
            "warnings": warnings,
        }

    return {
        "success": True,
        "ppttc_path": str(out_path.resolve()),
        "slide_count": len(document),
        "template": template,
        "errors": [],
        "warnings": warnings,
    }
