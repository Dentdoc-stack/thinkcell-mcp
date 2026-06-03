"""Self-contained test suite for the think-cell MCP server.

Run with::

    python tests.py

Exercises the chart builders (``charts/``), the ``.ppttc`` structural
validator (``validator.py``) and the zero-setup deck builder (``autodeck.py``)
-- the deterministic, OS-independent core. It does NOT run ``ppttc.exe`` (that
needs Windows + think-cell + PowerPoint). Exits non-zero if any check fails.
"""
from __future__ import annotations

import json
import re
import sys

from autodeck import (
    AUTO_TEMPLATE_PATH,
    _validate_chart,
    build_auto_deck,
    ensure_auto_template,
)
from branding import (
    DEFAULT_ACCENTS,
    apply_branding,
    recolor_theme_xml,
    strip_logo_from_master,
)
from charts import CHART_BUILDERS, chart_type_catalog, get_builder
from charts.area import AreaChart
from charts.bar import BarChart
from charts.combo import ComboChart
from charts.line import LineChart
from charts.mekko import MekkoChart
from charts.scatter import ScatterChart
from charts.stacked_bar import StackedBarChart
from charts.waterfall import WaterfallChart
from validator import validate_ppttc_data

_PASSED = 0
_FAILED = 0


def check(label: str, condition: bool) -> None:
    """Record one assertion."""
    global _PASSED, _FAILED
    if condition:
        _PASSED += 1
    else:
        _FAILED += 1
        print(f"  FAIL: {label}")


def has(errors: list[str], needle: str) -> bool:
    """True if any error message contains ``needle``."""
    return any(needle in error for error in errors)


# --------------------------------------------------------------------------
# chart registry / catalog
# --------------------------------------------------------------------------
def test_registry() -> None:
    check("8 chart types registered", len(CHART_BUILDERS) == 8)
    check("get_builder('bar') resolves", get_builder("bar") is BarChart)
    check("get_builder unknown -> None", get_builder("pie") is None)

    catalog = chart_type_catalog()
    check("catalog covers every type", set(catalog) == set(CHART_BUILDERS))
    check(
        "catalog entries describe data_shape + example",
        all(entry["data_shape"] and entry["example"]
            for entry in catalog.values()),
    )
    check(
        "catalog lists 'options' as optional param",
        "options" in catalog["bar"]["optional_params"],
    )


# --------------------------------------------------------------------------
# categorical builders
# --------------------------------------------------------------------------
def test_bar() -> None:
    good = BarChart(
        data=[{"category": "N", "Rev": 10}, {"category": "S", "Rev": 20}],
        title="T",
    )
    check("bar: valid data has no errors", good.validate() == [])

    categories, series = good.build_table()
    check("bar: categories extracted", categories == ["N", "S"])
    check("bar: series row is [name, *values]", series == [["Rev", 10, 20]])

    check(
        "bar: empty data rejected",
        has(BarChart(data=[], title="T").validate(), "non-empty list"),
    )
    check(
        "bar: non-object row rejected",
        has(BarChart(data=[1], title="T").validate(), "must be an object"),
    )
    check(
        "bar: row missing the category key rejected",
        has(
            BarChart(
                data=[{"category": "N", "Rev": 1}, {"Rev": 2}], title="T"
            ).validate(),
            "missing the category key",
        ),
    )
    check(
        "bar: non-numeric series value rejected",
        has(
            BarChart(
                data=[{"category": "N", "Rev": "lots"}], title="T"
            ).validate(),
            "must be a number",
        ),
    )


def test_missing_series_is_rejected() -> None:
    # A row that omits a series must be an error (would chart as 0, not a gap).
    builder = LineChart(
        data=[
            {"category": "Jan", "2024": 100, "2025": 130},
            {"category": "Feb", "2024": 110},  # missing "2025"
        ],
        title="T",
    )
    errors = builder.validate()
    check("line: row missing a series is rejected", has(errors, "2025"))
    check("line: error explains the gap rationale", has(errors, "gap"))


def test_min_rows() -> None:
    one_row = [{"category": "A", "V": 1}]
    check(
        "waterfall: needs >= 2 rows",
        has(WaterfallChart(data=one_row, title="T").validate(), "at least 2"),
    )
    check(
        "line: needs >= 2 rows",
        has(LineChart(data=one_row, title="T").validate(), "at least 2"),
    )
    check(
        "area: needs >= 2 rows",
        has(AreaChart(data=one_row, title="T").validate(), "at least 2"),
    )
    check(
        "bar: 1 row is allowed",
        BarChart(data=one_row, title="T").validate() == [],
    )
    check(
        "stacked_bar: 1 row is allowed",
        StackedBarChart(data=one_row, title="T").validate() == [],
    )
    check(
        "mekko: 1 row is allowed",
        MekkoChart(data=one_row, title="T").validate() == [],
    )


def test_category_key_option() -> None:
    builder = BarChart(
        data=[{"region": "N", "Rev": 10}, {"region": "S", "Rev": 20}],
        title="T",
        category_key="region",
    )
    check("category_key option: no errors", builder.validate() == [])
    categories, _ = builder.build_table()
    check("category_key option: honored", categories == ["N", "S"])


# --------------------------------------------------------------------------
# scatter
# --------------------------------------------------------------------------
def test_scatter() -> None:
    good = ScatterChart(
        data=[{"label": "A", "x": 1, "y": 2}, {"label": "B", "x": 3, "y": 4}],
        title="T",
    )
    check("scatter: valid xy data has no errors", good.validate() == [])
    categories, series = good.build_table()
    check("scatter: categories are X/Y", categories == ["X", "Y"])
    check("scatter: point is [label, x, y]", series[0] == ["A", 1, 2])

    check(
        "scatter: missing y rejected",
        has(ScatterChart(data=[{"x": 1}], title="T").validate(), "'x' and 'y'"),
    )
    check(
        "scatter: non-numeric x rejected",
        has(
            ScatterChart(data=[{"x": "a", "y": 1}], title="T").validate(),
            "['x'] must be a number",
        ),
    )

    bubble = ScatterChart(
        data=[{"x": 1, "y": 2, "size": 5}, {"x": 3, "y": 4, "size": 9}],
        title="T",
    )
    check("scatter: bubble (size) valid", bubble.validate() == [])
    categories, series = bubble.build_table()
    check("scatter: bubble adds Size column", categories == ["X", "Y", "Size"])
    check("scatter: bubble point carries size", series[0][-1] == 5)


# --------------------------------------------------------------------------
# combo
# --------------------------------------------------------------------------
def test_combo() -> None:
    data = [
        {"category": "Q1", "Rev": 100, "Margin": 18},
        {"category": "Q2", "Rev": 150, "Margin": 22},
    ]
    check("combo: two series is valid", ComboChart(data=data, title="T").validate() == [])
    check(
        "combo: one series rejected",
        has(
            ComboChart(
                data=[{"category": "Q1", "Rev": 100},
                      {"category": "Q2", "Rev": 150}],
                title="T",
            ).validate(),
            "at least 2",
        ),
    )
    check(
        "combo: valid line_series option accepted",
        ComboChart(data=data, title="T", line_series=["Margin"]).validate() == [],
    )
    check(
        "combo: line_series naming a missing series rejected",
        has(
            ComboChart(data=data, title="T", line_series=["Nope"]).validate(),
            "not a series",
        ),
    )
    check(
        "combo: line_series wrong type rejected",
        has(
            ComboChart(data=data, title="T", line_series="Margin").validate(),
            "must be a list",
        ),
    )


# --------------------------------------------------------------------------
# known_option_keys (typo guard)
# --------------------------------------------------------------------------
def test_known_option_keys() -> None:
    check(
        "bar accepts category_key option",
        "category_key" in BarChart.known_option_keys,
    )
    check(
        "combo also accepts line_series option",
        ComboChart.known_option_keys == {"category_key", "line_series"},
    )
    check(
        "scatter accepts no options",
        ScatterChart.known_option_keys == frozenset(),
    )


# --------------------------------------------------------------------------
# .ppttc structural validator
# --------------------------------------------------------------------------
def test_validator() -> None:
    valid_doc = [
        {
            "template": "deck.pptx",
            "data": [
                {
                    "name": "Chart1",
                    "table": [
                        [None, {"string": "Q1"}, {"string": "Q2"}],
                        [],
                        [{"string": "Rev"}, {"number": 10}, {"number": 20}],
                    ],
                }
            ],
        }
    ]
    check("validator: well-formed doc is valid", validate_ppttc_data(valid_doc)["valid"])

    check(
        "validator: top level must be a list",
        not validate_ppttc_data({"template": "x"})["valid"],
    )
    check(
        "validator: empty doc is invalid",
        not validate_ppttc_data([])["valid"],
    )
    check(
        "validator: missing 'template' rejected",
        has(validate_ppttc_data([{"data": []}])["errors"], "template"),
    )
    check(
        "validator: non-.pptx template rejected",
        has(
            validate_ppttc_data([{"template": "deck.key", "data": [
                {"name": "C", "table": [[None]]}]}])["errors"],
            ".pptx",
        ),
    )
    check(
        "validator: missing 'data' rejected",
        has(validate_ppttc_data([{"template": "d.pptx"}])["errors"], "data"),
    )
    check(
        "validator: element missing 'name' rejected",
        has(
            validate_ppttc_data([{"template": "d.pptx", "data": [
                {"table": [[None]]}]}])["errors"],
            "name",
        ),
    )
    check(
        "validator: element missing 'table' rejected",
        has(
            validate_ppttc_data([{"template": "d.pptx", "data": [
                {"name": "C"}]}])["errors"],
            "table",
        ),
    )
    check(
        "validator: empty table rejected",
        has(
            validate_ppttc_data([{"template": "d.pptx", "data": [
                {"name": "C", "table": []}]}])["errors"],
            "at least one row",
        ),
    )
    check(
        "validator: cell with no type key rejected",
        has(
            validate_ppttc_data([{"template": "d.pptx", "data": [
                {"name": "C", "table": [[{"fill": "red"}]]}]}])["errors"],
            "exactly one of",
        ),
    )
    check(
        "validator: cell with two type keys rejected",
        has(
            validate_ppttc_data([{"template": "d.pptx", "data": [
                {"name": "C", "table": [
                    [{"string": "a", "number": 1}]]}]}])["errors"],
            "exactly one type key",
        ),
    )
    check(
        "validator: non-numeric 'number' cell rejected",
        has(
            validate_ppttc_data([{"template": "d.pptx", "data": [
                {"name": "C", "table": [[{"number": "x"}]]}]}])["errors"],
            "must be numeric",
        ),
    )
    check(
        "validator: inconsistent row width rejected",
        has(
            validate_ppttc_data([{"template": "d.pptx", "data": [
                {"name": "C", "table": [
                    [{"string": "a"}, {"string": "b"}],
                    [{"string": "c"}]]}]}])["errors"],
            "same width",
        ),
    )
    check(
        "validator: null cells and empty separator row are allowed",
        validate_ppttc_data([{"template": "d.pptx", "data": [
            {"name": "C", "table": [[None, None], [], [None, None]]}]}])["valid"],
    )
    check(
        "validator: 'fill' alongside a type key is allowed",
        validate_ppttc_data([{"template": "d.pptx", "data": [
            {"name": "C", "table": [
                [{"number": 1, "fill": "#ff0000"}]]}]}])["valid"],
    )


# --------------------------------------------------------------------------
# validator: percentage cell support (bug fix)
# --------------------------------------------------------------------------
def test_validator_percentage() -> None:
    # A 'percentage' cell is part of the official think-cell schema.
    valid = [{"template": "d.pptx", "data": [
        {"name": "C", "table": [[{"percentage": 46.5}]]}]}]
    check(
        "validator: 'percentage' cell is accepted",
        validate_ppttc_data(valid)["valid"],
    )
    check(
        "validator: 'percentage' alongside 'fill' is accepted",
        validate_ppttc_data([{"template": "d.pptx", "data": [
            {"name": "C", "table": [
                [{"percentage": 12, "fill": "#00ff00"}]]}]}])["valid"],
    )
    check(
        "validator: non-numeric 'percentage' cell rejected",
        has(
            validate_ppttc_data([{"template": "d.pptx", "data": [
                {"name": "C", "table": [[{"percentage": "x"}]]}]}])["errors"],
            "must be numeric",
        ),
    )
    check(
        "validator: 'percentage' + 'number' is two type keys -> rejected",
        has(
            validate_ppttc_data([{"template": "d.pptx", "data": [
                {"name": "C", "table": [
                    [{"percentage": 1, "number": 2}]]}]}])["errors"],
            "exactly one type key",
        ),
    )


# --------------------------------------------------------------------------
# autodeck: zero-setup deck builder
# --------------------------------------------------------------------------
def _sample_slide() -> dict:
    return {
        "slide_title": "Competition: Germany",
        "left_title": "Market share",
        "right_title": "Our orders",
        "left_chart": {
            "categories": ["2024-01-01", "2025-01-01"],
            "series": [
                {"name": "Our brand", "values": [55, 55]},
                {"name": "Competitor 1", "values": [4, 4.5]},
            ],
        },
        "right_chart": {
            "categories": ["2024-01-01", "2025-01-01"],
            "series": [
                {"name": "Pending", "values": [0, 14], "fill": "#ff0000"},
                {"name": "Delivered", "values": [760, 747]},
            ],
        },
    }


def test_autodeck() -> None:
    # The deck-building checks need the think-cell template; on a machine
    # without think-cell (e.g. CI) it cannot be obtained, so skip just those.
    # The input-validation checks below do not need a template and run always.
    if ensure_auto_template():
        check(
            "autodeck: bundled template is available",
            AUTO_TEMPLATE_PATH.is_file(),
        )
        result = build_auto_deck(
            [_sample_slide(), _sample_slide()], "test_deck"
        )
        check("autodeck: valid two-slide deck succeeds", result["success"])
        check(
            "autodeck: slide_count reflects 2 slides",
            result["slide_count"] == 2,
        )
        check(
            "autodeck: ppttc written to output/ folder",
            bool(result["ppttc_path"])
            and result["ppttc_path"].endswith("test_deck.ppttc"),
        )

        # The generated .ppttc must itself pass structural validation.
        document = json.loads(
            open(result["ppttc_path"], encoding="utf-8").read()
        )
        check(
            "autodeck: each slide is its own .ppttc entry",
            isinstance(document, list) and len(document) == 2,
        )
        check(
            "autodeck: generated .ppttc passes the validator",
            validate_ppttc_data(document)["valid"],
        )
        names = {el["name"] for el in document[0]["data"]}
        check(
            "autodeck: uses the bundled named elements",
            names == {"SlideTitle", "LeftChartTitle", "RightChartTitle",
                      "LeftChart", "RightChart"},
        )
        left = next(
            el for el in document[0]["data"] if el["name"] == "LeftChart"
        )
        check(
            "autodeck: left chart header is a date axis",
            left["table"][0][0] is None
            and left["table"][0][1] == {"date": "2024-01-01"},
        )
        check(
            "autodeck: left chart series use 'percentage' cells",
            left["table"][1][1] == {"percentage": 55},
        )
        right = next(
            el for el in document[0]["data"] if el["name"] == "RightChart"
        )
        check(
            "autodeck: right chart series use 'number' cells",
            right["table"][1][1] == {"number": 0, "fill": "#ff0000"},
        )
        check(
            "autodeck: title-only slide is allowed",
            build_auto_deck(
                [{"slide_title": "Just a title"}], "x"
            )["success"],
        )
    else:
        print("  SKIP: autodeck deck-building checks "
              "(think-cell template unavailable)")

    # -- input validation (template-independent) --------------------------
    check(
        "autodeck: empty slides list rejected",
        has(build_auto_deck([], "x")["errors"], "non-empty list"),
    )
    check(
        "autodeck: a slide with no content is rejected",
        has(build_auto_deck([{}], "x")["errors"], "no content"),
    )
    length_mismatch = {
        "slide_title": "T",
        "left_chart": {
            "categories": ["2024-01-01", "2025-01-01"],
            "series": [{"name": "S", "values": [1]}],
        },
    }
    check(
        "autodeck: series length must match category count",
        has(build_auto_deck([length_mismatch], "x")["errors"], "must match"),
    )
    bad_fill = {
        "slide_title": "T",
        "right_chart": {
            "categories": ["2024-01-01"],
            "series": [{"name": "S", "values": [1], "fill": "red"}],
        },
    }
    check(
        "autodeck: invalid fill colour rejected",
        has(build_auto_deck([bad_fill], "x")["errors"], "hex colour"),
    )
    bad_value = {
        "slide_title": "T",
        "right_chart": {
            "categories": ["2024-01-01"],
            "series": [{"name": "S", "values": ["lots"]}],
        },
    }
    check(
        "autodeck: non-numeric series value rejected",
        has(build_auto_deck([bad_value], "x")["errors"], "must be a number"),
    )

    # -- _validate_chart: date vs. string category axes -------------------
    # _validate_chart is template-independent: it just returns (errors, table).
    str_errors, str_table = _validate_chart(
        {"categories": ["Q1", "Q2"],
         "series": [{"name": "Rev", "values": [10, 20]}]},
        "chart", "number",
    )
    check("autodeck: string categories are valid", str_errors == [])
    check(
        "autodeck: string categories emit 'string' header cells",
        str_table is not None
        and str_table[0] == [None, {"string": "Q1"}, {"string": "Q2"}],
    )

    date_errors, date_table = _validate_chart(
        {"categories": ["2024-01-01", "2025-01-01"],
         "series": [{"name": "Rev", "values": [10, 20]}]},
        "chart", "number",
    )
    check("autodeck: ISO-date categories are valid", date_errors == [])
    check(
        "autodeck: ISO-date categories emit 'date' header cells (regression)",
        date_table is not None
        and date_table[0] == [None, {"date": "2024-01-01"},
                              {"date": "2025-01-01"}],
    )

    # A single non-ISO entry tips the whole axis to a string axis.
    mixed_errors, mixed_table = _validate_chart(
        {"categories": ["2024-01-01", "Forecast"],
         "series": [{"name": "Rev", "values": [10, 20]}]},
        "chart", "number",
    )
    check("autodeck: mixed categories are valid", mixed_errors == [])
    check(
        "autodeck: any non-ISO category makes it a 'string' axis",
        mixed_table is not None
        and mixed_table[0] == [None, {"string": "2024-01-01"},
                               {"string": "Forecast"}],
    )

    empty_str_errors, empty_str_table = _validate_chart(
        {"categories": ["Q1", ""],
         "series": [{"name": "Rev", "values": [10, 20]}]},
        "chart", "number",
    )
    check(
        "autodeck: empty-string category is rejected",
        empty_str_table is None
        and has(empty_str_errors, "must be a non-empty string"),
    )

    nonstr_errors, nonstr_table = _validate_chart(
        {"categories": [2024],
         "series": [{"name": "Rev", "values": [10]}]},
        "chart", "number",
    )
    check(
        "autodeck: non-string category is rejected",
        nonstr_table is None
        and has(nonstr_errors, "must be a non-empty string"),
    )

    empty_cats_errors, empty_cats_table = _validate_chart(
        {"categories": [], "series": [{"name": "Rev", "values": []}]},
        "chart", "number",
    )
    check(
        "autodeck: empty categories list is rejected",
        empty_cats_table is None
        and has(empty_cats_errors, "non-empty list"),
    )


# --------------------------------------------------------------------------
# branding: re-theme the auto-deck template
# --------------------------------------------------------------------------
def test_branding() -> None:
    check("branding: 6 default accent colours", len(DEFAULT_ACCENTS) == 6)
    check(
        "branding: default accents are #RRGGBB hex",
        all(re.fullmatch(r"#[0-9A-Fa-f]{6}", c) for c in DEFAULT_ACCENTS),
    )

    # recolor_theme_xml swaps accent srgbClr values, leaves others intact.
    theme = (
        '<a:accent1><a:srgbClr val="111111"/></a:accent1>'
        '<a:accent2><a:srgbClr val="222222"/></a:accent2>'
        '<a:accent3><a:srgbClr val="333333"/></a:accent3>'
    )
    recoloured = recolor_theme_xml(theme, ["AABBCC", "DDEEFF"])
    check("branding: accent1 recoloured", 'val="AABBCC"' in recoloured)
    check("branding: accent2 recoloured", 'val="DDEEFF"' in recoloured)
    check(
        "branding: accent not in palette is left intact",
        'val="333333"' in recoloured,
    )

    # strip_logo_from_master drops visible pics, keeps hidden data shapes.
    master = (
        '<p:pic><p:nvPicPr><p:cNvPr id="1" name="Picture 4"/>'
        "</p:nvPicPr></p:pic>"
        '<p:pic><p:nvPicPr><p:cNvPr id="2" name="data" hidden="1"/>'
        "</p:nvPicPr></p:pic>"
    )
    stripped = strip_logo_from_master(master)
    check(
        "branding: visible logo picture removed",
        'name="Picture 4"' not in stripped,
    )
    check(
        "branding: hidden think-cell data shape kept",
        'name="data"' in stripped,
    )

    # apply_branding validates colours before touching any file.
    bad_hex = apply_branding("does_not_exist.pptx", accents=["nothex"])
    check("branding: invalid hex rejected", not bad_hex["success"])
    check(
        "branding: invalid hex error is descriptive",
        has(bad_hex["errors"], "hex colour"),
    )
    check(
        "branding: more than 6 accents rejected",
        not apply_branding(
            "does_not_exist.pptx", accents=["#000000"] * 7
        )["success"],
    )
    check(
        "branding: missing template file reported",
        has(
            apply_branding("does_not_exist.pptx")["errors"], "not found"
        ),
    )


# --------------------------------------------------------------------------
# palettes: named palettes + semantic colours
# --------------------------------------------------------------------------
def test_palettes() -> None:
    from palettes import (
        PALETTES,
        SEMANTIC,
        get_palette,
        list_palettes,
        semantic_fill,
    )

    hex_re = re.compile(r"#[0-9A-Fa-f]{6}")
    check("palettes: registry is non-empty", len(PALETTES) >= 2)
    check(
        "palettes: every colour is #RRGGBB hex",
        all(
            hex_re.fullmatch(c)
            for colors in PALETTES.values()
            for c in colors
        ),
    )

    # vizro maps to PowerPoint's six accent slots; full keeps all ten.
    check("palettes: 'vizro' fits 6 accent slots", len(get_palette("vizro")) == 6)
    check("palettes: 'vizro_full' has 10 colours", len(get_palette("vizro_full")) == 10)
    check("palettes: get_palette is case-insensitive", get_palette("VIZRO") == get_palette("vizro"))
    check(
        "palettes: get_palette returns a copy",
        get_palette("vizro") is not PALETTES["vizro"],
    )
    try:
        get_palette("does_not_exist")
        check("palettes: unknown name raises KeyError", False)
    except KeyError:
        check("palettes: unknown name raises KeyError", True)

    # Semantic colours carry meaning: gains teal, losses red, baseline gray.
    check(
        "palettes: all five semantic roles present",
        set(SEMANTIC) == {"positive", "negative", "neutral", "info", "warning"},
    )
    check("palettes: gain is the positive colour", semantic_fill(120) == SEMANTIC["positive"])
    check("palettes: loss is the negative colour", semantic_fill(-40) == SEMANTIC["negative"])
    check("palettes: zero is the neutral colour", semantic_fill(0) == SEMANTIC["neutral"])
    # reverse=True is for 'down is good' metrics (cost, churn).
    check("palettes: reverse flips a drop to positive", semantic_fill(-40, reverse=True) == SEMANTIC["positive"])
    check("palettes: reverse flips a rise to negative", semantic_fill(120, reverse=True) == SEMANTIC["negative"])
    # Garbage in -> neutral out, never a crash.
    check("palettes: non-numeric falls back to neutral", semantic_fill("nope") == SEMANTIC["neutral"])
    check("palettes: NaN falls back to neutral", semantic_fill(float("nan")) == SEMANTIC["neutral"])
    check("palettes: bool is not treated as a number", semantic_fill(True) == SEMANTIC["neutral"])

    described = list_palettes()
    check("palettes: list_palettes exposes palettes", "vizro" in described["palettes"])
    check("palettes: list_palettes exposes semantic map", described["semantic"] == SEMANTIC)


# --------------------------------------------------------------------------
# autodeck: semantic auto-colouring of a series by sign
# --------------------------------------------------------------------------
def test_autodeck_semantic() -> None:
    from palettes import SEMANTIC

    def number_cells(result):
        doc = json.loads(open(result["ppttc_path"], encoding="utf-8").read())
        chart = next(e for e in doc[0]["data"] if e["name"] == "RightChart")
        return [c for c in chart["table"][1][1:] if isinstance(c, dict)]

    def deck(values, **opts):
        cats = [f"2024-{m + 1:02d}-01" for m in range(len(values))]
        return build_auto_deck([{
            "slide_title": "Variance",
            "right_chart": {
                "categories": cats,
                "series": [{"name": "Delta", "values": values, **opts}],
            },
        }], "_test_semantic")

    # The colour checks read the generated .ppttc back from disk, which needs
    # the bundled think-cell template. Skip them where it is unavailable (e.g.
    # CI without a think-cell install); the validation check below runs always.
    if ensure_auto_template():
        res = deck([100.0, -40.0, 0.0], semantic=True)
        check("autodeck semantic: build succeeds", res["success"])
        cells = number_cells(res)
        check("autodeck semantic: gain bar is teal", cells[0]["fill"] == SEMANTIC["positive"])
        check("autodeck semantic: loss bar is red", cells[1]["fill"] == SEMANTIC["negative"])
        check("autodeck semantic: zero bar is gray", cells[2]["fill"] == SEMANTIC["neutral"])

        res = deck([-40.0, 50.0], semantic=True, reverse=True)
        cells = number_cells(res)
        check("autodeck semantic: reverse makes a drop teal", cells[0]["fill"] == SEMANTIC["positive"])

        res = deck([-40.0, 50.0], semantic=True, fill="#123456")
        cells = number_cells(res)
        check(
            "autodeck semantic: explicit fill overrides",
            all(c["fill"] == "#123456" for c in cells),
        )

        res = deck([10.0, 20.0])
        cells = number_cells(res)
        check(
            "autodeck semantic: un-flagged series stays uncoloured",
            all("fill" not in c for c in cells),
        )
    else:
        print("  SKIP: autodeck semantic colouring checks "
              "(think-cell template unavailable)")

    # Input validation runs before the template check inside build_auto_deck,
    # so this rejection holds regardless of whether the template is present.
    res = deck([1, 2], semantic="yes")
    check("autodeck semantic: non-boolean flag rejected", not res["success"])


# --------------------------------------------------------------------------
# server: every MCP tool is importable (the eval's invariant check)
# --------------------------------------------------------------------------
def test_server_smoke() -> None:
    import server

    tools = [
        "create_chart", "build_presentation", "create_auto_deck",
        "set_deck_branding", "list_color_palettes", "convert_to_pptx",
        "validate_ppttc", "list_chart_types", "diagnose_thinkcell",
    ]
    for name in tools:
        check(f"server: tool '{name}' is defined", hasattr(server, name))


# --------------------------------------------------------------------------
# build_presentation: one .ppttc entry per slide
# --------------------------------------------------------------------------
def test_build_presentation() -> None:
    import server

    rows = [{"category": "A", "V": 1}, {"category": "B", "V": 2}]
    result = server.build_presentation(
        slides=[
            {"title": "Slide One", "charts": [
                {"chart_type": "bar", "chart_name": "C1", "data": rows}]},
            {"title": "Slide Two", "charts": [
                {"chart_type": "line", "chart_name": "C2", "data": rows}]},
        ],
        template_path="deck.pptx",
        output_name="_test_buildpres",
    )
    check("build_presentation: two-slide deck succeeds", result["success"])
    check("build_presentation: slide_count is 2", result.get("slide_count") == 2)
    check("build_presentation: chart_count is 2", result.get("chart_count") == 2)

    document = json.loads(open(result["ppttc_path"], encoding="utf-8").read())
    check(
        "build_presentation: one .ppttc entry per slide (not all on one)",
        isinstance(document, list) and len(document) == 2,
    )
    check(
        "build_presentation: slide 1 entry holds only its own elements",
        {e["name"] for e in document[0]["data"]} == {"Title1", "C1"},
    )
    check(
        "build_presentation: slide 2 entry holds only its own elements",
        {e["name"] for e in document[1]["data"]} == {"Title2", "C2"},
    )

    # A name may repeat across slides -- each slide is a separate entry.
    across = server.build_presentation(
        slides=[
            {"charts": [{"chart_type": "bar", "chart_name": "Chart1",
                         "data": [{"category": "A", "V": 1}]}]},
            {"charts": [{"chart_type": "bar", "chart_name": "Chart1",
                         "data": [{"category": "A", "V": 2}]}]},
        ],
        template_path="deck.pptx",
        output_name="_test_buildpres_across",
    )
    check(
        "build_presentation: same name on different slides is allowed",
        across["success"],
    )

    # Within one slide a duplicate name is still rejected.
    dup = server.build_presentation(
        slides=[{"charts": [
            {"chart_type": "bar", "chart_name": "X",
             "data": [{"category": "A", "V": 1}]},
            {"chart_type": "bar", "chart_name": "X",
             "data": [{"category": "A", "V": 2}]}]}],
        template_path="deck.pptx",
        output_name="_test_buildpres_dup",
    )
    check(
        "build_presentation: duplicate name within a slide rejected",
        not dup["success"] and has(dup["errors"], "unique within a slide"),
    )


def main() -> int:
    suites = [
        test_registry,
        test_bar,
        test_missing_series_is_rejected,
        test_min_rows,
        test_category_key_option,
        test_scatter,
        test_combo,
        test_known_option_keys,
        test_validator,
        test_validator_percentage,
        test_autodeck,
        test_branding,
        test_palettes,
        test_autodeck_semantic,
        test_server_smoke,
        test_build_presentation,
    ]
    for suite in suites:
        suite()

    total = _PASSED + _FAILED
    print(f"\n{_PASSED}/{total} checks passed.")
    if _FAILED:
        print(f"{_FAILED} FAILED.")
        return 1
    print("All checks passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
