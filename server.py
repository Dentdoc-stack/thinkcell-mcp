"""think-cell MCP server.

Exposes its core tools over stdio for building think-cell JSON automation
(``.ppttc``) files and converting them to PowerPoint (``.pptx``) with
think-cell's ``ppttc.exe``:

    * create_chart        -- build a single-chart .ppttc file
    * build_presentation  -- combine many charts/slides into one .ppttc file
    * create_auto_deck    -- build a multi-slide .ppttc with NO template setup
    * set_deck_branding   -- recolour the auto-deck template / drop the logo
    * list_color_palettes -- list built-in palettes and semantic colours
    * import_chart_data   -- read chart/table data from a local .csv/.xlsx
    * convert_to_pptx     -- run ppttc.exe to produce a .pptx
    * validate_ppttc      -- structurally validate a .ppttc document
    * list_chart_types    -- describe every supported chart type
    * diagnose_thinkcell  -- check why think-cell automation is not working

plus ``insert_into_powerpoint``, which puts a converted deck's slides into
the presentation open in PowerPoint on this PC, and three template tools
(``library.py``): ``list_chart_templates`` (think-cell's own chart library
and saved slides), ``template_from_slide`` (any slide with think-cell charts
becomes a template) and ``fill_chart_template``, and ``update_open_charts``, which
refills named charts in the open presentation in place.

Run directly (``python server.py``) to start the server on stdio. As a
remote connector (Claude in PowerPoint, claude.ai) it runs with ``--http``
behind a tunnel and an OAuth sign-in; ``connector.py`` sets that up and
starts it -- see CONNECTOR.md.
"""
from __future__ import annotations

import argparse
import functools
import html
import inspect
import json
import sys
import threading
from pathlib import Path
from typing import Any
from urllib.parse import urlsplit

import anyio
from mcp.server.fastmcp import FastMCP
from mcp.server.transport_security import TransportSecuritySettings

from charts import (
    CHART_BUILDERS,
    ChartBuilder,
    ChartError,
    chart_type_catalog,
    get_builder,
    write_ppttc_document,
    write_ppttc_slides,
)
from autodeck import AUTO_TEMPLATE_PATH, OUTPUT_DIR, build_auto_deck, ensure_auto_template
from branding import apply_branding, load_style
from palettes import get_palette, list_palettes
from converter import convert_ppttc
from data_import import import_spreadsheet
from diagnostics import run_diagnostics
from layout import widen_single_charts_in_pptx
from validator import validate_ppttc_data

PROJECT_ROOT = Path(__file__).resolve().parent
DEFAULT_TEMPLATE = "template.pptx"

# Set by ``--http``: the server is then reachable from the internet, so path
# arguments are kept inside REMOTE_ROOTS (plus the connector's chosen folders).
REMOTE = False
REMOTE_ROOTS = [OUTPUT_DIR, PROJECT_ROOT / "templates"]

__version__ = "0.3.0"  # keep in step with pyproject.toml and manifest.json
mcp = FastMCP("thinkcell")


# --------------------------------------------------------------------------
# helpers
# --------------------------------------------------------------------------
def _inside_roots(path: str) -> bool:
    """True when ``path`` resolves inside one of REMOTE_ROOTS."""
    target = Path(path).expanduser().resolve()
    return any(target.is_relative_to(root.resolve()) for root in REMOTE_ROOTS)


def _local_paths(*names: str):
    """In remote mode, refuse the named path arguments outside REMOTE_ROOTS."""
    def wrap(fn):
        signature = inspect.signature(fn)

        @functools.wraps(fn)
        def inner(*args, **kwargs):
            if REMOTE:
                bound = signature.bind_partial(*args, **kwargs).arguments
                for name in names:
                    value = bound.get(name)
                    if value and not _inside_roots(value):
                        roots = ", ".join(str(r) for r in REMOTE_ROOTS)
                        return {
                            "success": False,
                            "errors": [
                                f"{name} '{value}' is outside the folders this "
                                f"connector may use ({roots}). Add folders "
                                f"with: python connector.py setup"
                            ],
                        }
            return fn(*args, **kwargs)
        return inner
    return wrap


_JOB_LOCK = threading.Lock()  # think-cell and PowerPoint take one job at a time


def _slow_tool(fn):
    """Register ``fn`` as a tool that runs in a worker thread.

    The SDK calls plain ``def`` tools on the event loop, so one ppttc.exe run
    (up to 3 minutes) would stall sign-in and every other request of the
    remote connector. The plain function is returned for direct calls.
    """
    @functools.wraps(fn)
    async def runner(*args, **kwargs):
        def job():
            with _JOB_LOCK:
                return fn(*args, **kwargs)
        return await anyio.to_thread.run_sync(job)
    mcp.tool()(runner)
    return fn


def _chunked_only(app):
    """ASGI wrapper that drops Content-Length, so uvicorn chunks every reply.

    Antivirus web shields that scan local HTTP (Norton 360 here) rewrite a
    reply that has a Content-Length into a chunked one with no end, which the
    tunnel reports as 520; replies uvicorn chunks itself pass intact.
    """
    async def wrapped(scope, receive, send):
        async def strip(message):
            if message["type"] == "http.response.start":
                message = {**message, "headers": [
                    (key, value) for key, value in message.get("headers", [])
                    if key.lower() != b"content-length"]}
            await send(message)
        await app(scope, receive, strip if scope["type"] == "http" else send)
    return wrapped


# think-cell Assist calls MCP servers from a browser page on this origin; its
# manual lists the CORS headers below (Configure thinkcell Assist > CORS).
ASSIST_ORIGIN = "https://app.prod.ai.think-cell.com"
_CORS_ALLOW_HEADERS = ("content-type, authorization, mcp-protocol-version, mcp-session-id, "
                       "last-event-id, mcp-method, mcp-name, ngrok-skip-browser-warning")
_CORS_EXPOSE_HEADERS = "mcp-session-id, mcp-protocol-version, www-authenticate"


def _browser_access(app, allowed_origins: list[str]):
    """ASGI wrapper for browser clients such as think-cell Assist.

    MCP's Streamable HTTP transport says servers MUST check ``Origin`` and
    answer 403 to an invalid one. Requests without ``Origin`` (Claude's cloud
    connectors, curl) pass. Allowed origins get CORS headers, and preflights
    are answered here, before sign-in, because browsers send them without
    credentials. GET and DELETE on /mcp get 405: the server is stateless, so
    there is no standalone stream or session to open or close.
    """
    from starlette.responses import JSONResponse, Response

    allowed = set(allowed_origins)

    async def wrapped(scope, receive, send):
        if scope["type"] != "http":
            return await app(scope, receive, send)
        headers = {k.decode().lower(): v.decode() for k, v in scope.get("headers", [])}
        origin = headers.get("origin")
        cors = {"Access-Control-Allow-Origin": origin, "Vary": "Origin",
                "Access-Control-Expose-Headers": _CORS_EXPOSE_HEADERS} if origin in allowed else {}
        if origin and not cors:
            response = JSONResponse({"jsonrpc": "2.0", "error": {
                "code": -32600, "message": f"Origin {origin} is not allowed"}}, status_code=403)
        elif scope["method"] == "OPTIONS" and cors:
            response = Response(status_code=204, headers={
                **cors, "Access-Control-Allow-Methods": "POST, GET, DELETE, OPTIONS",
                "Access-Control-Allow-Headers": _CORS_ALLOW_HEADERS, "Access-Control-Max-Age": "600"})
        elif scope["method"] in ("GET", "DELETE") and scope["path"].rstrip("/") == "/mcp":
            response = Response(status_code=405, headers={**cors, "Allow": "POST, OPTIONS"})
        else:
            async def add_cors(message):
                if message["type"] == "http.response.start" and cors:
                    message = {**message, "headers": [*message.get("headers", []),
                               *((k.lower().encode(), v.encode()) for k, v in cors.items())]}
                await send(message)
            return await app(scope, receive, add_cors)
        await response(scope, receive, send)
    return wrapped


_CONNECT_PAGE = """<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>Connect think-cell</title>
<style>body{{font:16px/1.5 system-ui,sans-serif;background:#f4f5f7;color:#1d1f23;margin:0;
display:grid;place-items:center;min-height:100vh}}main{{background:#fff;border-radius:12px;
box-shadow:0 2px 12px #0002;padding:32px;max-width:420px;margin:16px}}h1{{font-size:22px;margin:0 0 8px}}
input{{font:600 22px ui-monospace,monospace;letter-spacing:3px;width:100%;box-sizing:border-box;
padding:10px;border:1px solid #c6c9d0;border-radius:8px;text-transform:uppercase}}
button{{margin-top:16px;width:100%;padding:12px;font-size:16px;border:0;border-radius:8px;
background:#1d4ed8;color:#fff;cursor:pointer}}.err{{color:#b42318;margin:8px 0}}
code{{background:#eef0f3;padding:2px 6px;border-radius:4px}}</style></head><body><main>
<h1>Connect {client} to think-cell</h1>{body}</main></body></html>"""

_CONNECT_FORM = """<p>This lets {client} make think-cell charts and put them in PowerPoint on the PC
running this connector. Type the connect code shown by <code>python connector.py code</code> on that PC.</p>
{error}<form method="post"><input type="hidden" name="request" value="{request}">
<input name="code" autocomplete="one-time-code" autofocus required placeholder="XXXX-XXXX" aria-label="Connect code">
<button>Connect</button></form>"""


def _connect_route(provider):
    """``/connect``: where the user approves Claude with the connect code."""
    from starlette.responses import HTMLResponse, RedirectResponse

    headers = {"X-Frame-Options": "DENY", "Content-Security-Policy": "frame-ancestors 'none'",
               "Cache-Control": "no-store"}

    async def connect(request):
        form = await request.form() if request.method == "POST" else request.query_params
        request_id = str(form.get("request", ""))
        error = ""
        if request.method == "POST":
            url, message = provider.approve(request_id, str(form.get("code", "")))
            if url:
                return RedirectResponse(url, status_code=302, headers=headers)
            error = f"<p class=err>{html.escape(message)}</p>"
        client = provider.client_name(request_id)
        if client is None:  # expired, unknown, or dropped after too many wrong codes
            body = error or "<p class=err>This sign-in has expired. Start again from Claude (Connect).</p>"
            return HTMLResponse(_CONNECT_PAGE.format(client="Claude", body=body), 400, headers)
        body = _CONNECT_FORM.format(client=html.escape(client), error=error, request=html.escape(request_id))
        return HTMLResponse(_CONNECT_PAGE.format(client=html.escape(client), body=body),
                            400 if error else 200, headers)
    return connect


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
@_local_paths("template_path")
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
            area, combo, table, process. Call ``list_chart_types`` for the
            data shape of each.
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
@_local_paths("template_path")
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
              For RAG-style colouring, set ``semantic`` (bool) on a series to
              colour each bar by the sign of its value -- gains teal, losses
              red, zero gray (great for variance/bridge/delta charts). Add
              ``reverse`` (bool) when down is good (cost, churn) to flip the
              colours. An explicit ``fill`` always overrides ``semantic``.

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


@_slow_tool
@_local_paths("style_file")
def set_deck_branding(
    accent_colors: list[str] | None = None,
    palette: str = "",
    remove_logo: bool = True,
    style_file: str = "",
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
        palette: Optional name of a built-in palette (``vizro``, ``vizro_full``
            or ``big_four``) to use instead of passing raw ``accent_colors``.
            See ``list_color_palettes`` for the full set. Ignored when
            ``accent_colors`` is given.
        remove_logo: When True (default) the think-cell logo is removed from
            the slide master.
        style_file: Optional path to a think-cell style file (``.xml``, e.g.
            your organisation's, or one from think-cell's ``styles`` folder).
            It is loaded with think-cell's own ``LoadStyle`` and sets what
            theme colours cannot: fill schemes, line styles and label fonts.
            Needs PowerPoint with think-cell on this PC.

    Returns:
        ``{"success": bool, "template": str, "accents": list[str],
        "logo_removed": bool, "errors": list[str]}``, plus ``"style"`` (the
        name think-cell reports) when ``style_file`` is given.
    """
    if accent_colors is None and palette:
        try:
            accent_colors = get_palette(palette)
        except KeyError as exc:
            return {
                "success": False,
                "template": str(AUTO_TEMPLATE_PATH),
                "accents": [],
                "logo_removed": False,
                "errors": [str(exc)],
            }
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
    result = apply_branding(AUTO_TEMPLATE_PATH, accent_colors, remove_logo)
    if style_file and result.get("success"):
        styled = load_style(AUTO_TEMPLATE_PATH, style_file)
        result.update(success=styled["success"], style=styled["style"], errors=styled["errors"])
    return result


@mcp.tool()
def list_color_palettes() -> dict[str, Any]:
    """List the built-in colour palettes and semantic colours.

    Use this to pick a ready-made, colourblind-safe palette for
    ``set_deck_branding`` instead of hand-picking hex codes, and to look up the
    semantic colours (positive/negative/neutral/info/warning) for colouring a
    waterfall's gains and losses or a delta series.

    Returns:
        ``{"palettes": {name: [hex, ...]}, "semantic": {role: hex, ...},
        "default": str}``. Pass any ``palettes`` key as ``set_deck_branding``'s
        ``palette`` argument; use ``semantic`` values as per-series ``fill``
        colours in ``create_auto_deck`` (e.g. losses in the negative red).
    """
    return list_palettes()


@mcp.tool()
@_local_paths("file_path")
def import_chart_data(file_path: str, sheet_name: str = "") -> dict[str, Any]:
    """Read chart/table data from a local ``.csv`` or ``.xlsx`` spreadsheet.

    The first row is treated as column headers; every following row becomes
    one row object with those headers as keys -- the exact ``data`` shape
    ``create_chart``/``build_presentation`` (for any categorical chart type)
    and the ``table`` chart type already expect, so the result can be passed
    straight through without reshaping.

    Args:
        file_path: Path to a local ``.csv`` or ``.xlsx``/``.xlsm`` file.
        sheet_name: For ``.xlsx``, the worksheet to read. Defaults to the
            active sheet. Ignored for ``.csv``.

    Returns:
        ``{"success": bool, "data": list[dict] | None, "columns": list[str],
        "row_count": int, "percentage_columns": list[str], "errors": list[str]}``.
        ``percentage_columns`` (``.xlsx`` only) lists columns Excel formatted
        as a percentage -- the values are already rescaled to whole points
        (21.9 for 21.9%, matching what think-cell's `percentage` cells
        expect); pass this list straight through as the ``table`` chart
        type's ``options.percentage_columns``.
    """
    return import_spreadsheet(file_path, sheet_name)


@_slow_tool
@_local_paths("ppttc_path", "output_path")
def convert_to_pptx(
    ppttc_path: str,
    output_path: str = "",
    widen_single_chart: bool = True,
) -> dict[str, Any]:
    """Convert a ``.ppttc`` file to PowerPoint (.pptx) using think-cell's ppttc.exe.

    Args:
        ppttc_path: Path to an existing ``.ppttc`` file.
        output_path: Optional destination ``.pptx`` path. Defaults to the
            input path with a ``.pptx`` extension.
        widen_single_chart: After a successful conversion, widen any
            ``create_auto_deck`` slide that only populated
            ``left_chart``/``right_chart`` (not both) to the full slide
            content width instead of leaving it next to an empty
            placeholder box. This only ever touches the *converted* output
            (editing chart positions before conversion breaks think-cell's
            automation link -- the data silently fails to update), and it
            no-ops for anything that isn't an auto-deck-shaped slide, so
            it's safe to leave on for ``build_presentation``/``create_chart``
            output too. Default True.

    Returns:
        On success: ``{"success": True, "pptx_path": str, "stdout": str,
        "stderr": str, "error": None, "widened_slides": int}``. On failure:
        ``{"success": False, "pptx_path": None, "error": str, ...}`` with
        diagnostic detail. ``widened_slides`` counts how many slides were
        adjusted by ``widen_single_chart`` (0 when disabled or inapplicable).
    """
    result = convert_ppttc(ppttc_path, output_path or None)
    if not result.get("success"):
        return result

    widened = 0
    if widen_single_chart:
        try:
            document = json.loads(Path(ppttc_path).read_text(encoding="utf-8"))
            if isinstance(document, list):
                widened = widen_single_charts_in_pptx(
                    result["pptx_path"], document
                )
        except (OSError, json.JSONDecodeError):
            pass  # widening is a cosmetic best-effort step, never fatal here
    result["widened_slides"] = widened
    return result


@mcp.tool()
@_local_paths("ppttc_path")
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
@_local_paths("template_path")
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


@_slow_tool
@_local_paths("pptx_path")
def insert_into_powerpoint(pptx_path: str, after_slide: int = -1) -> dict[str, Any]:
    """Insert every slide of a ``.pptx`` into the presentation open in PowerPoint.

    Use it after ``convert_to_pptx`` to put the think-cell slides into the deck
    the user is working on (for example the deck Claude in PowerPoint has open
    on this PC). PowerPoint copies the slides with their think-cell data, so
    the charts stay editable in think-cell.

    Args:
        pptx_path: Path to the ``.pptx`` to insert (a ``convert_to_pptx``
            ``pptx_path``).
        after_slide: Slide number to insert after: -1 (default) appends at
            the end, 0 inserts before the first slide.

    Returns:
        ``{"success": bool, "presentation": str, "inserted": int,
        "first_slide": int, "errors": list[str]}``. Fails when PowerPoint is
        not running on this PC or has no presentation open.
    """
    path = Path(pptx_path).expanduser().resolve()
    if not path.is_file():
        return {"success": False, "errors": [f"File not found: {pptx_path}"]}
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return {"success": False, "errors": ["pywin32 is not installed: python -m pip install pywin32"]}
    pythoncom.CoInitialize()
    try:
        pres = win32com.client.GetActiveObject("PowerPoint.Application").ActivePresentation
        count = pres.Slides.Count
        index = count if after_slide < 0 else min(after_slide, count)
        inserted = pres.Slides.InsertFromFile(str(path), index)
        return {"success": True, "presentation": pres.Name, "inserted": inserted,
                "first_slide": index + 1, "errors": []}
    except Exception as exc:  # pywintypes.com_error: PowerPoint closed, no deck open, ...
        return {"success": False, "errors": [
            f"PowerPoint did not take the slides ({exc}). Open the deck in "
            f"PowerPoint on this PC and try again."]}
    finally:
        pythoncom.CoUninitialize()


def _chart_summary(entry: dict[str, Any]) -> dict[str, Any]:
    return {**{k: entry[k] for k in ("id", "group", "title") if k in entry},
            "charts": [{"name": c["name"], "kind": c["kind"], "categories": c["categories"],
                        "series_count": c["series"]} for c in entry["charts"]],
            "texts": entry.get("texts", [])}


@_slow_tool
def list_chart_templates(search: str = "") -> dict[str, Any]:
    """List ready-to-fill think-cell slide templates: think-cell's own chart library plus the user's saved slides.

    think-cell's library (installed with think-cell) has stacked, clustered and 100% columns and bars,
    butterfly/tornado, line, area, combination, football field, box plot, candlestick, Mekko, pie, doughnut,
    bar-of-pie, scatter, bubble, build-up and build-down waterfalls, funnel, and charts with CAGR, difference and
    value-line annotations. Every chart keeps think-cell's styling and stays editable in think-cell. The first
    call copies the library from the think-cell install (a few seconds). Templates the user saved with
    ``template_from_slide`` are listed with ids ``user/<name>``.

    Fill one with ``fill_chart_template``, then ``insert_into_powerpoint``.

    Args:
        search: Optional words to filter by (matched against id, group and title), e.g. "waterfall", "mekko".

    Returns:
        ``{"templates": [{"id", "group", "title", "charts": [{"name", "kind", "categories", "series_count"}],
        "texts": [...]}], "errors": [...]}``. ``categories``/``series_count`` show what the example holds;
        new data can have any number of categories and series. ``texts`` are the slide's own texts (title,
        headings, comment box) that ``fill_chart_template`` can replace.
    """
    import library

    try:
        entries = library.ensure_library() + library.user_templates()
    except (OSError, ValueError) as exc:
        return {"templates": [], "errors": [f"Could not read think-cell's chart library: {exc}"]}
    words = search.lower().split()
    hits = [e for e in entries if all(w in f"{e['id']} {e['group']} {e['title']}".lower() for w in words)]
    return {"templates": [_chart_summary(e) for e in hits], "errors": []}


@_slow_tool
@_local_paths("pptx_path")
def template_from_slide(slide_number: int, name: str, pptx_path: str = "") -> dict[str, Any]:
    """Save a slide with think-cell charts as a reusable template, keeping its exact look.

    Use it to update a slide of the user's own deck with new numbers in the same style: save it here, fill it
    with ``fill_chart_template`` (id ``user/<name>``), then ``insert_into_powerpoint``. Its think-cell charts
    are named ``c1``, ``c2``, ... in reading order (top to bottom, left to right); nothing is typed in PowerPoint.

    Args:
        slide_number: 1-based slide number.
        name: Name for the template (letters, digits, spaces).
        pptx_path: The deck to take the slide from. Leave empty to use the presentation open in PowerPoint on
            this PC (it is saved to a copy first; the open deck is not changed).

    Returns:
        ``{"success": bool, "template": {"id", "title", "charts", "texts"}, "errors": [...]}``.
    """
    import tempfile

    import library

    tmp = None
    try:
        if pptx_path:
            source = Path(pptx_path).expanduser().resolve()
            if not source.is_file():
                return {"success": False, "errors": [f"File not found: {pptx_path}"]}
        else:
            import pythoncom
            import win32com.client

            pythoncom.CoInitialize()
            try:
                pres = win32com.client.GetActiveObject("PowerPoint.Application").ActivePresentation
                tmp = Path(tempfile.mkdtemp()) / "open_deck.pptx"
                pres.SaveCopyAs(str(tmp))
            except Exception as exc:  # pywintypes.com_error: PowerPoint closed, no deck open
                return {"success": False, "errors": [
                    f"Could not read the open presentation ({exc}). Open the deck in PowerPoint on this PC, "
                    f"or give pptx_path."]}
            finally:
                pythoncom.CoUninitialize()
            source = tmp
        entry = library.save_user_template(source, slide_number, name)
    except (OSError, ValueError, KeyError) as exc:
        return {"success": False, "errors": [str(exc)]}
    finally:
        if tmp:
            import shutil
            shutil.rmtree(tmp.parent, ignore_errors=True)
    if not entry["charts"]:
        return {"success": False, "template": _chart_summary(entry),
                "errors": [f"Slide {slide_number} has no think-cell charts that can be filled"]}
    return {"success": True, "template": _chart_summary(entry), "errors": []}


@_slow_tool
def fill_chart_template(
    template_id: str,
    charts: dict[str, Any],
    title: str = "",
    texts: dict[str, str] | None = None,
    output_name: str = "",
) -> dict[str, Any]:
    """Fill a template from ``list_chart_templates`` / ``template_from_slide`` with data and make the .pptx.

    Runs think-cell's ppttc.exe, so the result is real think-cell charts in the template's style. Then
    ``insert_into_powerpoint`` puts the slide into the open deck.

    Args:
        template_id: A template id, e.g. ``"waterfall-1"`` or ``"user/q3-revenue"``.
        charts: Data per chart name (``c1``, ``c2``, ...)::

              {"c1": {"categories": ["2024", "2025", "2026"],
                      "series": [{"name": "EMEA", "values": [40, 45, 52]},
                                 {"name": "APAC", "values": [20, 24, 31], "fill": "#1f77b4"}],
                      "percent": false}}

            ``percent: true`` writes percentage cells (whole points: 45 = 45%). Pie/doughnut: one category,
            one series per segment. Scatter/bubble: categories ``["X", "Y"]`` (or ``["X", "Y", "Size"]``),
            one series per point. Waterfall: a cell ``"e"`` is a calculated total (think-cell's "equals").
            For full control pass ``{"table": [...]}``, a raw think-cell datasheet (rows of cells, see
            ``validate_ppttc``).
        title: New slide title (optional).
        texts: Replace the slide's own texts: ``{"<current text>": "<new text>"}`` with the current text
            exactly as ``texts`` lists it; ``""`` removes that shape; ``\\n`` starts a new line (each line
            keeps the formatting of the matching old line, so bullet lists stay bullets).
        output_name: Filename stem for the result in ``output/``.

    Returns:
        ``{"success", "pptx_path", "filled": [...], "not_filled": [...], "not_checked": [...], "warnings",
        "errors"}``. ``filled`` charts were checked to show the new categories; pie and scatter charts are
        ``not_checked``.
    """
    import library

    path = library.template_path(template_id)
    if path is None:
        return {"success": False, "pptx_path": None,
                "errors": [f"Unknown template '{template_id}': see list_chart_templates"]}
    try:
        return library.fill_template(path, charts, title=title, output_name=output_name, texts=texts)
    except (OSError, ValueError) as exc:
        return {"success": False, "pptx_path": None, "errors": [str(exc)]}


@_slow_tool
def update_open_charts(slide_number: int, charts: dict[str, Any]) -> dict[str, Any]:
    """Refill think-cell charts on a slide of the presentation open in PowerPoint, in place.

    Use it to give a chart that is already in the user's deck new numbers without inserting a new slide: the
    chart keeps its position, style and annotations. It works on charts that have a think-cell name. Slides
    made by ``fill_chart_template`` (``c1``, ``c2``, ...) and ``create_auto_deck`` (``LeftChart``,
    ``RightChart``) keep their names after ``insert_into_powerpoint``; a wrong name returns the names on the
    slide. The open deck is changed but not saved (the user can undo or close without saving).

    It uses think-cell's ``UpdateBatch`` API, which reads its data from Excel: a hidden Excel is started for
    the call and closed again. Needs PowerPoint and Excel with think-cell on this PC.

    Args:
        slide_number: 1-based number of the slide in the open presentation.
        charts: Data per chart name, in the same form as ``fill_chart_template``'s ``charts``.

    Returns:
        ``{"success", "presentation", "slide", "updated": [...], "not_updated": [...], "not_checked": [...],
        "errors"}``. ``updated`` charts were checked to show the new categories.
    """
    import library

    try:
        return library.update_open_charts(charts, slide_number)
    except ImportError:
        return {"success": False, "errors": ["pywin32 is not installed: python -m pip install pywin32"]}
    except (OSError, ValueError, KeyError) as exc:
        return {"success": False, "errors": [str(exc)]}


def main() -> None:
    """Start the MCP server on stdio, or with ``--http`` as a remote connector."""
    parser = argparse.ArgumentParser(description="think-cell MCP server")
    parser.add_argument("--http", action="store_true",
                        help="serve as a remote connector on 127.0.0.1 (connector.py run starts it with a tunnel)")
    parser.add_argument("--port", type=int, help="default: the connector config's port, else 8765")
    parser.add_argument("--public-url", help="the tunnel's https address (default: from the connector config)")
    args = parser.parse_args()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    if not args.http:
        mcp.run(transport="stdio")
        return
    import uvicorn

    uvicorn.run(connector_app(args.public_url), host="127.0.0.1",
                port=args.port or _connector_config().get("port", 8765), log_level="info")


def _connector_config() -> dict[str, Any]:
    from connector import load_config

    return load_config()


def connector_app(public_url: str | None = None):
    """The ASGI app of the remote connector: MCP at ``/mcp`` behind OAuth sign-in."""
    from mcp.server.auth.provider import ProviderTokenVerifier
    from mcp.server.auth.settings import AuthSettings, ClientRegistrationOptions, RevocationOptions

    from connector import STATE_FILE
    from oauth import ConnectorAuth

    global REMOTE
    config = _connector_config()
    public_url = (public_url or config.get("public_url") or "").rstrip("/")
    if not public_url or not config.get("connect_code"):
        sys.exit("The connector is not set up yet: run  python connector.py setup")
    REMOTE = True
    REMOTE_ROOTS.extend(Path(folder) for folder in config.get("folders", []))

    provider = ConnectorAuth(STATE_FILE, public_url, config["connect_code"], config.get("api_key"))
    mcp.settings.auth = AuthSettings(
        issuer_url=public_url,
        resource_server_url=f"{public_url}/mcp",
        client_registration_options=ClientRegistrationOptions(enabled=True),
        revocation_options=RevocationOptions(enabled=True),
    )
    # shortcut: FastMCP takes these only in its constructor, but the module-level
    # `mcp` must stay plain for stdio; set them before the app is built.
    mcp._auth_server_provider = provider
    mcp._token_verifier = ProviderTokenVerifier(provider)
    mcp.settings.streamable_http_path = "/mcp"
    mcp.settings.stateless_http = True
    # The tunnel forwards its own public Host header, which the SDK's
    # localhost-only check rejects; _browser_access checks Origin instead.
    mcp.settings.transport_security = TransportSecuritySettings(enable_dns_rebinding_protection=False)
    mcp.custom_route("/connect", methods=["GET", "POST"])(_connect_route(provider))
    own = urlsplit(public_url)  # the /connect form posts from the server's own origin
    origins = [ASSIST_ORIGIN, f"{own.scheme}://{own.netloc}", *config.get("origins", [])]
    print(f"think-cell connector: {public_url}/mcp\n"
          f"Folders it may use: {', '.join(str(r) for r in REMOTE_ROOTS)}\n"
          f"Browser origins allowed: {', '.join(origins)}", file=sys.stderr)
    from modern_http import modern_era

    # MCP 2026-07-28 requests (stateless, no initialize) are answered by modern_era;
    # legacy ones go on to the SDK.
    app = modern_era(mcp.streamable_http_app(), mcp, mcp._token_verifier,
                     f"{public_url}/.well-known/oauth-protected-resource/mcp",
                     {"name": mcp.name, "version": __version__}, mcp.instructions)
    return _chunked_only(_browser_access(app, origins))


if __name__ == "__main__":
    main()
