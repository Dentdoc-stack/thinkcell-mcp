"""think-cell MCP server.

Exposes eight tools over stdio for building think-cell JSON automation
(``.ppttc``) files and converting them to PowerPoint (``.pptx``) with
think-cell's ``ppttc.exe``:

    * create_chart       -- build a single-chart .ppttc file
    * build_presentation -- combine many charts/slides into one .ppttc file
    * create_auto_deck   -- build a multi-slide .ppttc with NO template setup
    * set_deck_branding  -- recolour the auto-deck template / drop the logo
    * convert_to_pptx    -- run ppttc.exe to produce a .pptx
    * validate_ppttc     -- structurally validate a .ppttc document
    * list_chart_types   -- describe every supported chart type
    * diagnose_thinkcell -- check why think-cell automation is not working

Run directly (``python server.py``) to start the server on stdio.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from mcp.server.fastmcp import FastMCP

from charts import (
    CHART_BUILDERS,
    ChartBuilder,
    ChartError,
    chart_type_catalog,
    get_builder,
    write_ppttc_document,
    write_ppttc_slides,
)
from autodeck import AUTO_TEMPLATE_PATH, build_auto_deck, ensure_auto_template
from branding import apply_branding
from converter import convert_ppttc
from diagnostics import run_diagnostics
from validator import validate_ppttc_data

PROJECT_ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = PROJECT_ROOT / "output"
DEFAULT_TEMPLATE = "template.pptx"

mcp = FastMCP("thinkcell")


# --------------------------------------------------------------------------
# helpers
# --------------------------------------------------------------------------
def _safe_filename(name: str, fallback: str) -> str:
    """Turn an arbitrary title into a safe filesystem stem."""
    cleaned = "".join(
        char for char in name if char.isalnum() or char in (" ", "-", "_")
    ).strip()
    return cleaned.replace(" ", "_") or fallback


def _template_error(template_path: str) -> str | None:
    """Return an error string if a given template path is not a .pptx."""
    if template_path and not template_path.lower().endswith(".pptx"):
        return f"template_path must reference a .pptx file, got '{template_path}'"
    return None


def _supported_types() -> str:
    """Comma-separated list of supported chart-type names."""
    return ", ".join(CHART_BUILDERS)


# --------------------------------------------------------------------------
# tools
# --------------------------------------------------------------------------
@mcp.tool()
def create_chart(
    chart_type: str,
    data: list[dict[str, Any]],
    title: str,
    template_path: str = "",
    chart_name: str = "Chart1",
    options: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Build a think-cell ``.ppttc`` file containing a single chart.

    Args:
        chart_type: One of waterfall, bar, stacked_bar, line, scatter, mekko,
            area, combo. Call ``list_chart_types`` for the data shape of each.
        data: List of row objects describing the chart. The expected shape
            depends on ``chart_type`` (see ``list_chart_types``).
        title: Human-readable chart title; also used for the output filename.
        template_path: Optional path to the think-cell PowerPoint template
            (.pptx). When omitted, the placeholder ``template.pptx`` is written
            and must be corrected before conversion.
        chart_name: Name of the target think-cell element in the template.
            This MUST match the name given to the chart inside think-cell:
            select the chart in PowerPoint and type the name in the "Name"
            field of its think-cell mini toolbar. If the names do not match,
            ppttc.exe runs but updates nothing. Defaults to ``Chart1``.
        options: Optional chart-type-specific options. ``category_key`` (any
            categorical chart) forces which row key is the category column;
            ``line_series`` (combo only) lists which series are lines.

    Returns:
        ``{"success": bool, "ppttc_path": str | None, "errors": list[str], ...}``.
    """
    builder_cls = get_builder(chart_type)
    if builder_cls is None:
        return {
            "success": False,
            "ppttc_path": None,
            "errors": [
                f"Unknown chart_type '{chart_type}'. "
                f"Supported types: {_supported_types()}"
            ],
        }

    template_issue = _template_error(template_path)
    if template_issue:
        return {"success": False, "ppttc_path": None, "errors": [template_issue]}

    options = options or {}
    if not isinstance(options, dict):
        return {
            "success": False,
            "ppttc_path": None,
            "errors": ["`options` must be an object of chart-type options"],
        }
    unknown = set(options) - builder_cls.known_option_keys
    if unknown:
        return {
            "success": False,
            "ppttc_path": None,
            "errors": [
                f"Unknown option(s) {sorted(unknown)} for chart_type "
                f"'{chart_type}'. Supported option(s): "
                f"{sorted(builder_cls.known_option_keys)}"
            ],
        }

    builder = builder_cls(
        data=data, title=title, chart_name=chart_name, **options
    )
    errors = builder.validate()
    if errors:
        return {"success": False, "ppttc_path": None, "errors": errors}

    template = template_path or DEFAULT_TEMPLATE
    out_path = OUTPUT_DIR / f"{_safe_filename(title, chart_type)}.ppttc"
    try:
        written = write_ppttc_document(template, [builder], out_path)
    except ChartError as exc:
        return {
            "success": False,
            "ppttc_path": None,
            "errors": [f"think-cell rejected the data: {exc}"],
        }
    except OSError as exc:
        return {
            "success": False,
            "ppttc_path": None,
            "errors": [f"Could not write the .ppttc file: {exc}"],
        }

    return {
        "success": True,
        "ppttc_path": written,
        "chart_type": chart_type,
        "chart_name": builder.chart_name,
        "template": template,
        "errors": [],
    }


@mcp.tool()
def build_presentation(
    slides: list[dict[str, Any]],
    template_path: str,
    output_name: str,
) -> dict[str, Any]:
    """Combine multiple charts across multiple slides into one ``.ppttc`` file.

    Each slide becomes its own top-level ``.ppttc`` entry, which think-cell
    renders as one PowerPoint slide -- so an N-slide ``slides`` list yields an
    N-slide deck.

    Args:
        slides: Non-empty list of slide objects. Each slide has a ``title``
            string and a non-empty ``charts`` list. Each chart entry is an
            object with ``chart_type`` and ``data`` (as in ``create_chart``),
            plus optional ``title`` and ``chart_name``. ``chart_name`` must
            match the named think-cell element in the template; when omitted,
            the names Chart1, Chart2, ... are assigned in order. A slide title
            is published as a text field named ``title_field`` if given, else
            ``Title1``, ``Title2``, ... Chart and text-field names must be
            unique *within a slide*; different slides may reuse the same name
            (each slide is a separate ``.ppttc`` entry).
        template_path: Path to the think-cell PowerPoint template (.pptx)
            that contains every named chart (and any text fields).
        output_name: Filename stem for the generated ``.ppttc`` file.

    Returns:
        ``{"success": bool, "ppttc_path": str | None, "slide_count": int,
        "chart_count": int, "errors": list[str]}``.
    """
    if not isinstance(slides, list) or len(slides) == 0:
        return {
            "success": False,
            "ppttc_path": None,
            "errors": ["`slides` must be a non-empty list of slide objects"],
        }
    if not template_path:
        return {
            "success": False,
            "ppttc_path": None,
            "errors": ["`template_path` is required for build_presentation"],
        }
    template_issue = _template_error(template_path)
    if template_issue:
        return {"success": False, "ppttc_path": None, "errors": [template_issue]}

    errors: list[str] = []
    # One (builders, textfields) group per slide -> one .ppttc entry each.
    slide_groups: list[tuple[list[ChartBuilder], list[tuple[str, str]]]] = []
    chart_count = 0

    for slide_index, slide in enumerate(slides):
        if not isinstance(slide, dict):
            errors.append(f"slides[{slide_index}] must be an object")
            continue

        # Names are scoped per slide: each slide is its own .ppttc entry,
        # so a name only has to be unique within that slide.
        slide_builders: list[ChartBuilder] = []
        slide_textfields: list[tuple[str, str]] = []
        used_names: set[str] = set()
        auto_index = 1

        slide_title = slide.get("title", "")
        if slide_title:
            field_name = str(
                slide.get("title_field") or f"Title{slide_index + 1}"
            )
            used_names.add(field_name)
            slide_textfields.append((field_name, str(slide_title)))

        charts = slide.get("charts")
        if not isinstance(charts, list) or len(charts) == 0:
            errors.append(
                f"slides[{slide_index}] must have a non-empty 'charts' list"
            )
            continue

        for chart_index, config in enumerate(charts):
            where = f"slides[{slide_index}].charts[{chart_index}]"
            if not isinstance(config, dict):
                errors.append(f"{where} must be an object")
                continue

            chart_type = config.get("chart_type")
            builder_cls = (
                get_builder(chart_type) if isinstance(chart_type, str) else None
            )
            if builder_cls is None:
                errors.append(
                    f"{where} has unknown chart_type '{chart_type}'. "
                    f"Supported types: {_supported_types()}"
                )
                continue

            # Resolve a name unique within this slide (explicit or assigned).
            chart_name = config.get("chart_name")
            if not chart_name:
                while f"Chart{auto_index}" in used_names:
                    auto_index += 1
                chart_name = f"Chart{auto_index}"
            chart_name = str(chart_name)
            if chart_name in used_names:
                errors.append(
                    f"{where} reuses the name '{chart_name}'; chart and "
                    f"text-field names must be unique within a slide"
                )
                continue
            used_names.add(chart_name)

            reserved = {"chart_type", "data", "title", "chart_name"}
            options = {
                key: value
                for key, value in config.items()
                if key not in reserved
            }
            unknown = set(options) - builder_cls.known_option_keys
            if unknown:
                errors.append(
                    f"{where} has unknown key(s) {sorted(unknown)}; allowed "
                    f"option(s) for '{chart_type}': "
                    f"{sorted(builder_cls.known_option_keys)}"
                )
                continue
            builder = builder_cls(
                data=config.get("data"),
                title=str(config.get("title", slide_title or chart_type)),
                chart_name=chart_name,
                **options,
            )
            chart_errors = builder.validate()
            if chart_errors:
                errors.extend(f"{where}: {error}" for error in chart_errors)
                continue
            slide_builders.append(builder)

        slide_groups.append((slide_builders, slide_textfields))
        chart_count += len(slide_builders)

    if errors:
        return {"success": False, "ppttc_path": None, "errors": errors}

    out_path = OUTPUT_DIR / f"{_safe_filename(output_name, 'presentation')}.ppttc"
    try:
        written = write_ppttc_slides(template_path, slide_groups, out_path)
    except ChartError as exc:
        return {
            "success": False,
            "ppttc_path": None,
            "errors": [f"think-cell rejected the data: {exc}"],
        }
    except OSError as exc:
        return {
            "success": False,
            "ppttc_path": None,
            "errors": [f"Could not write the .ppttc file: {exc}"],
        }

    return {
        "success": True,
        "ppttc_path": written,
        "slide_count": len(slides),
        "chart_count": chart_count,
        "template": template_path,
        "errors": [],
    }


@mcp.tool()
def create_auto_deck(
    slides: list[dict[str, Any]],
    output_name: str = "auto_deck",
) -> dict[str, Any]:
    """Build a multi-slide ``.ppttc`` deck with ZERO template preparation.

    Unlike ``create_chart`` / ``build_presentation`` -- which need a ``.pptx``
    template whose think-cell charts you named by hand in PowerPoint -- this
    tool uses the official think-cell automation template bundled with the
    project (``templates/thinkcell_auto.pptx``). Its think-cell elements are
    already named, so the user never opens PowerPoint or names anything: they
    just supply data. Each slide becomes one ``.ppttc`` entry (one slide).

    Every slide has the same fixed layout: a slide title, plus a left and a
    right chart, each with its own title. All fields are optional, but a slide
    must carry at least one. The left chart renders ``percentage`` values; the
    right chart renders plain ``number`` values.

    Args:
        slides: Non-empty list of slide objects. Each slide object accepts:

            * ``slide_title`` (str, optional): heading shown on the slide.
            * ``left_title`` (str, optional): title above the left chart.
            * ``right_title`` (str, optional): title above the right chart.
            * ``left_chart`` (object, optional): the left (percentage) chart.
            * ``right_chart`` (object, optional): the right (number) chart.

            A chart object has:

            * ``categories`` (list[str]): non-empty list of x-axis labels.
              When every label is an ISO date (``"YYYY-MM-DD"``) the chart
              gets a date axis; otherwise the labels form a category axis.
            * ``series`` (list[object]): non-empty list of data series. Each
              series object has ``name`` (str), ``values`` (list of numbers,
              one per category) and an optional ``fill`` (hex colour like
              ``"#ff0000"`` applied to every point of that series).

        output_name: Filename stem for the generated ``.ppttc`` file. The file
            is written to the project's ``output/`` folder.

    Example invocation::

        {
          "output_name": "Germany_Review",
          "slides": [
            {
              "slide_title": "Competition: Germany",
              "left_title": "Market share",
              "right_title": "Our orders (10K)",
              "left_chart": {
                "categories": ["2024-01-01", "2025-01-01"],
                "series": [
                  {"name": "Our brand", "values": [55, 55]},
                  {"name": "Competitor 1", "values": [4, 4.5]}
                ]
              },
              "right_chart": {
                "categories": ["2024-01-01", "2025-01-01"],
                "series": [
                  {"name": "Pending", "values": [0, 14], "fill": "#ff0000"},
                  {"name": "Delivered", "values": [760, 747]}
                ]
              }
            }
          ]
        }

    Returns:
        ``{"success": bool, "ppttc_path": str | None, "slide_count": int,
        "template": str, "errors": list[str], "warnings": list[str]}``. On
        success the deck can be converted straight away with ``convert_to_pptx``
        -- no manual step. ``warnings`` is non-fatal advice worth surfacing to
        the user: percentage values that look like decimals (think-cell expects
        whole points, e.g. 21.9 for 21.9%) or a single-category date axis that
        would render with no visible bars.
    """
    return build_auto_deck(slides, output_name)


@mcp.tool()
def set_deck_branding(
    accent_colors: list[str] | None = None,
    remove_logo: bool = True,
) -> dict[str, Any]:
    """Re-brand the template that ``create_auto_deck`` uses.

    think-cell's automation template ships with think-cell's own logo (top
    right of every slide) and a green/blue accent palette. This tool rewrites
    that bundled template -- once -- so every future ``create_auto_deck`` deck
    inherits your colours and drops the logo. It only touches standard
    PowerPoint parts (theme + slide master); the named think-cell elements are
    left intact, so automation keeps working.

    By default the template is already de-branded on first use (logo removed,
    a professional Big-Four-inspired palette applied). Call this tool to set a
    custom palette or to toggle the logo.

    Args:
        accent_colors: Up to six ``#RRGGBB`` hex colours for the chart accent
            palette (theme accent1..accent6). When omitted, a professional
            blue-anchored default palette is applied. Per-series ``fill``
            values in ``create_auto_deck`` still override these per series.
        remove_logo: When True (default) the think-cell logo is removed from
            the slide master.

    Returns:
        ``{"success": bool, "template": str, "accents": list[str],
        "logo_removed": bool, "errors": list[str]}``.
    """
    if not ensure_auto_template():
        return {
            "success": False,
            "template": str(AUTO_TEMPLATE_PATH),
            "accents": [],
            "logo_removed": False,
            "errors": [
                "The think-cell automation template is not available. Run "
                "create_auto_deck once, or install think-cell, so the "
                "template can be obtained first."
            ],
        }
    return apply_branding(AUTO_TEMPLATE_PATH, accent_colors, remove_logo)


@mcp.tool()
def convert_to_pptx(ppttc_path: str, output_path: str = "") -> dict[str, Any]:
    """Convert a ``.ppttc`` file to PowerPoint (.pptx) using think-cell's ppttc.exe.

    Args:
        ppttc_path: Path to an existing ``.ppttc`` file.
        output_path: Optional destination ``.pptx`` path. Defaults to the
            input path with a ``.pptx`` extension.

    Returns:
        On success: ``{"success": True, "pptx_path": str, "stdout": str,
        "stderr": str, "error": None}``. On failure: ``{"success": False,
        "pptx_path": None, "error": str, ...}`` with diagnostic detail.
    """
    return convert_ppttc(ppttc_path, output_path or None)


@mcp.tool()
def validate_ppttc(ppttc_path: str = "", ppttc_json: str = "") -> dict[str, Any]:
    """Structurally validate a think-cell ``.ppttc`` document.

    Provide exactly one of the two inputs. ``validate_ppttc`` checks the
    ``.ppttc`` schema: the template/data/name/table structure, required
    fields, cell typing rules and row-width consistency. Per-chart-type data
    rules are enforced separately by ``create_chart``/``build_presentation``.

    Args:
        ppttc_path: Path to a ``.ppttc`` file to read and validate.
        ppttc_json: Raw ``.ppttc`` JSON content as a string.

    Returns:
        ``{"valid": bool, "errors": list[str]}``.
    """
    if ppttc_path and ppttc_json:
        return {
            "valid": False,
            "errors": [
                "Provide either 'ppttc_path' or 'ppttc_json', not both"
            ],
        }
    if ppttc_path:
        path = Path(ppttc_path).expanduser()
        if not path.is_file():
            return {"valid": False, "errors": [f"File not found: {ppttc_path}"]}
        try:
            raw = path.read_text(encoding="utf-8")
        except OSError as exc:
            return {"valid": False, "errors": [f"Could not read file: {exc}"]}
    elif ppttc_json:
        raw = ppttc_json
    else:
        return {
            "valid": False,
            "errors": ["Provide either 'ppttc_path' or 'ppttc_json'"],
        }

    try:
        document = json.loads(raw)
    except json.JSONDecodeError as exc:
        return {"valid": False, "errors": [f"Invalid JSON: {exc}"]}

    return validate_ppttc_data(document)


@mcp.tool()
def list_chart_types() -> dict[str, Any]:
    """List every supported chart type with its parameters and an example.

    Returns:
        ``{"count": int, "chart_types": {<name>: {"required_params": [...],
        "optional_params": [...], "data_shape": str, "example": {...}}}}``.
    """
    return {
        "count": len(CHART_BUILDERS),
        "chart_types": chart_type_catalog(),
    }


@mcp.tool()
def diagnose_thinkcell(template_path: str = "") -> dict[str, Any]:
    """Diagnose why think-cell automation may not be working.

    think-cell exposes no general API to "control" its PowerPoint add-in --
    the ``.ppttc`` / ``ppttc.exe`` pipeline is the only supported automation.
    When that pipeline does nothing, the cause is environmental. This tool
    runs read-only Windows registry and filesystem checks (no processes are
    launched) covering: the think-cell installation, Microsoft PowerPoint,
    the think-cell COM add-in registration and load state, PowerPoint's
    disabled-add-in list, think-cell registry/licence presence, and -- if a
    template is supplied -- whether a ``.pptx`` actually contains think-cell
    content.

    Args:
        template_path: Optional path to a ``.pptx`` think-cell template to
            inspect as part of the report.

    Returns:
        ``{"platform_supported": bool, "summary": {"status", "headline",
        "recommendations"}, "checks": {<name>: {"status", "summary",
        "details", "recommendation"}}}``. ``status`` is one of ok, warning,
        error or info.
    """
    return run_diagnostics(template_path or None)


def main() -> None:
    """Start the MCP server on the stdio transport."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
