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
from charts.process import ProcessChart
from charts.scatter import ScatterChart
from charts.stacked_bar import StackedBarChart
from charts.table import TableChart
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
    check("10 chart types registered", len(CHART_BUILDERS) == 10)
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
# table (raw .ppttc bypass: per-cell conditional fill)
# --------------------------------------------------------------------------
def test_table() -> None:
    rows = [
        {"Department": "Agriculture", "Proposed": 2217, "Difference": -453},
        {"Department": "Health", "Proposed": 2768, "Difference": 3553},
    ]
    good = TableChart(data=rows, title="T")
    check("table: valid rows have no errors", good.validate() == [])
    name, table = good.build_raw_table()
    check("table: header row is column names", table[0] == [
        {"string": "Department"}, {"string": "Proposed"}, {"string": "Difference"},
    ])
    check("table: negative Difference cell not auto-filled without a rule",
          "fill" not in table[1][2])

    highlighted = TableChart(
        data=rows, title="T",
        columns=["Department", "Proposed", "Difference"],
        highlight=[{"column": "Difference", "rule": "negative"}],
    )
    check("table: highlight option accepted", highlighted.validate() == [])
    _, table2 = highlighted.build_raw_table()
    check("table: negative cell gets a fill", "fill" in table2[1][2])
    check("table: positive cell has no fill", "fill" not in table2[2][2])

    check(
        "table: unknown column in row rejected",
        has(
            TableChart(data=[{"X": 1}], title="T", columns=["Y"]).validate(),
            "not in columns",
        ),
    )
    check(
        "table: bad highlight rule rejected",
        has(
            TableChart(
                data=rows, title="T",
                highlight=[{"column": "Difference", "rule": "nope"}],
            ).validate(),
            "must be one of",
        ),
    )
    check(
        "table: gt/lt/eq rule without fill rejected",
        has(
            TableChart(
                data=rows, title="T",
                highlight=[{"column": "Proposed", "rule": "gt", "value": 1000}],
            ).validate(),
            "fill is required",
        ),
    )
    check(
        "table: no columns at all rejected",
        has(TableChart(data=[{}], title="T").validate(), "at least one column"),
    )

    pct = TableChart(
        data=[{"Region": "N", "Share": 48}], title="T",
        percentage_columns=["Share"],
    )
    check("table: percentage_columns accepted", pct.validate() == [])
    _, pct_table = pct.build_raw_table()
    check("table: percentage column uses 'percentage' key",
          "percentage" in pct_table[1][1])


# --------------------------------------------------------------------------
# process (raw .ppttc bypass: dated milestone timeline)
# --------------------------------------------------------------------------
def test_process() -> None:
    milestones = [
        {"date": "2025-05-27", "label": "Roadmap Launched", "detail": "by CM"},
        {"date": "2026-01-19", "label": "Launch of v2.0"},
    ]
    good = ProcessChart(data=milestones, title="T")
    check("process: valid milestones have no errors", good.validate() == [])
    name, table = good.build_raw_table()
    check("process: header has one cell per milestone", len(table[0]) == 2)
    check("process: header cells are date-typed",
          all("date" in cell for cell in table[0]))
    check("process: milestone row starts with 'Milestone' label",
          table[1][0] == {"string": "Milestone"})
    check("process: detail row present when any milestone has detail",
          len(table) == 3 and table[2][0] == {"string": "Detail"})

    check(
        "process: single milestone rejected",
        has(ProcessChart(data=milestones[:1], title="T").validate(), "at least 2"),
    )
    check(
        "process: bad date format rejected",
        has(
            ProcessChart(
                data=[{"date": "27-05-2025", "label": "X"},
                      {"date": "2026-01-01", "label": "Y"}],
                title="T",
            ).validate(),
            "ISO date",
        ),
    )
    check(
        "process: missing label rejected",
        has(
            ProcessChart(
                data=[{"date": "2025-01-01"}, {"date": "2026-01-01", "label": "Y"}],
                title="T",
            ).validate(),
            "non-empty string",
        ),
    )

    current = ProcessChart(data=milestones, title="T", current_date="2025-09-01")
    check("process: current_date between milestones accepted", current.validate() == [])
    _, current_table = current.build_raw_table()
    check("process: current_date inserted chronologically (3 columns)",
          len(current_table[0]) == 3)
    marked_index = next(
        i for i, cell in enumerate(current_table[0])
        if cell.get("date") == "2025-09-01"
    )
    check("process: current_date cell gets a fill", "fill" in current_table[0][marked_index])
    check("process: current_date has no label (None cell)",
          current_table[1][marked_index + 1] is None)

    check(
        "process: bad current_date rejected",
        has(
            ProcessChart(
                data=milestones, title="T", current_date="not-a-date"
            ).validate(),
            "current_date",
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
    fill_doc = lambda fill: [{"template": "d.pptx", "data": [  # noqa: E731
        {"name": "C", "table": [[{"number": 1, "fill": fill}]]}]}]
    check("validator: rgb(r,g,b) fill accepted (think-cell's schema)",
          validate_ppttc_data(fill_doc("rgb(255,0,0)"))["valid"])
    for bad in ("red", "#f00", "rgb(255, 0, 0)", "rgb(256,0,0)", 5):
        check(f"validator: fill {bad!r} rejected",
              has(validate_ppttc_data(fill_doc(bad))["errors"], "fill must be"))
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
    # think-cell manual, Cell data types: dates must be "YYYY-MM-DD".
    def date_doc(value):
        return [{"template": "d.pptx", "data": [{"name": "C", "table": [[None, {"date": value}]]}]}]
    check("validator: ISO date accepted", validate_ppttc_data(date_doc("2024-01-31"))["valid"])
    for bad in ("01/02/2024", "2024-1-5", "2024-02-30", "2024-01-31T00:00"):
        check(f"validator: non-ISO date {bad!r} rejected", not validate_ppttc_data(date_doc(bad))["valid"])
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
# layout: single-chart auto-widen variants of the auto-deck template
# --------------------------------------------------------------------------
def test_layout() -> None:
    import zipfile

    from layout import _find_chart_frames, _slide_keep, widen_single_chart, \
        widen_single_charts_in_pptx

    if not ensure_auto_template():
        print("  SKIP: layout checks (think-cell template unavailable)")
        return

    with zipfile.ZipFile(AUTO_TEMPLATE_PATH) as zin:
        original_xml = zin.read("ppt/slides/slide1.xml").decode("utf-8")

    frames = _find_chart_frames(original_xml)
    check("layout: base template has exactly 2 real chart frames", len(frames) == 2)
    union_width = (frames[1]["x"] + frames[1]["cx"]) - frames[0]["x"]

    widened_left = widen_single_chart(original_xml, "left")
    check("layout: widening left changes the XML", widened_left != original_xml)
    kept, dropped = sorted(
        _find_chart_frames(widened_left), key=lambda f: f["cx"], reverse=True
    )
    check("layout: kept (left) chart spans the union width", kept["cx"] == union_width)
    check("layout: dropped (right) chart moved off-slide", dropped["x"] >= 12192000)

    widened_right = widen_single_chart(original_xml, "right")
    kept_r, dropped_r = sorted(
        _find_chart_frames(widened_right), key=lambda f: f["cx"], reverse=True
    )
    check("layout: kept (right) chart spans the union width", kept_r["cx"] == union_width)
    check("layout: dropped (left) chart moved off-slide", dropped_r["x"] >= 12192000)

    check(
        "layout: widen_single_chart is a no-op on unexpected shapes",
        widen_single_chart("<p:sld></p:sld>", "left") == "<p:sld></p:sld>",
    )

    # _slide_keep: derives which chart to widen straight from a .ppttc entry.
    check(
        "layout: _slide_keep picks 'left' when only LeftChart is present",
        _slide_keep({"data": [{"name": "LeftChart", "table": []}]}) == "left",
    )
    check(
        "layout: _slide_keep picks 'right' when only RightChart is present",
        _slide_keep({"data": [{"name": "RightChart", "table": []}]}) == "right",
    )
    check(
        "layout: _slide_keep is None when both charts are present",
        _slide_keep({"data": [
            {"name": "LeftChart", "table": []}, {"name": "RightChart", "table": []},
        ]}) is None,
    )
    check(
        "layout: _slide_keep is None for a non-auto-deck slide",
        _slide_keep({"data": [{"name": "Chart1", "table": []}]}) is None,
    )

    # End-to-end through the real pipeline: convert a mixed 2-slide deck
    # (dual-chart, then single-chart) and confirm both (a) the pushed data
    # is intact and (b) the single-chart slide is actually widened. This is
    # the regression test for the bug this module works around: widening
    # the *template* before conversion silently breaks the data push.
    deck = build_auto_deck(
        [
            _sample_slide(),
            {
                "slide_title": "Solo",
                "left_title": "Solo chart",
                "left_chart": {
                    "categories": ["2024-01-01", "2025-01-01"],
                    "series": [{"name": "UniqueMarker", "values": [777, 888]}],
                },
            },
        ],
        "_test_layout_pipeline",
    )
    check("layout: mixed-deck build succeeds", deck["success"])

    from converter import convert_ppttc

    conv = convert_ppttc(deck["ppttc_path"])
    check("layout: mixed-deck conversion succeeds", conv["success"])
    if conv["success"]:
        document = json.loads(open(deck["ppttc_path"], encoding="utf-8").read())
        widened = widen_single_charts_in_pptx(conv["pptx_path"], document)
        check("layout: exactly the one single-chart slide was widened", widened == 1)

        from pptx import Presentation

        prs = Presentation(conv["pptx_path"])
        slide2_charts = {
            shape.name: shape
            for shape in prs.slides[1].shapes
            if getattr(shape, "has_chart", False)
        }
        widened_shape = max(slide2_charts.values(), key=lambda s: s.width)
        narrowed_shape = min(slide2_charts.values(), key=lambda s: s.width)
        check(
            "layout: widened slide2 chart carries the real pushed data "
            "(not stale template placeholder values)",
            [round(v) for v in widened_shape.chart.series[0].values] == [777, 888],
        )
        check(
            "layout: widened shape is wider than the unwidened one",
            widened_shape.width > narrowed_shape.width,
        )
        check(
            "layout: dropped shape moved off the visible slide",
            narrowed_shape.left >= 12192000,
        )

        slide1_charts = [
            shape for shape in prs.slides[0].shapes
            if getattr(shape, "has_chart", False)
        ]
        check(
            "layout: dual-chart slide1 left untouched (still 2 on-slide charts)",
            len(slide1_charts) == 2
            and all(shape.left < 12192000 for shape in slide1_charts),
        )

    # widen_single_chart=False on convert_to_pptx: no widening applied.
    import server

    off = server.create_auto_deck(
        [{
            "slide_title": "Solo2",
            "left_chart": {
                "categories": ["2024-01-01", "2025-01-01"],
                "series": [{"name": "S", "values": [1, 2]}],
            },
        }],
        "_test_layout_opt_out",
    )
    conv_off = server.convert_to_pptx(off["ppttc_path"], widen_single_chart=False)
    check("layout: convert_to_pptx succeeds with widening disabled", conv_off["success"])
    check(
        "layout: widened_slides is 0 when widen_single_chart=False",
        conv_off.get("widened_slides") == 0,
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
    check("palettes: 'kp_ta_team' fits 6 accent slots", len(get_palette("kp_ta_team")) == 6)
    check(
        "palettes: 'kp_ta_team' starts with the deep teal accent",
        get_palette("kp_ta_team")[0] == "#004B53",
    )
    check(
        "palettes: 'kp_ta_team' listed in list_palettes()",
        "kp_ta_team" in list_palettes()["palettes"],
    )
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
        "set_deck_branding", "list_color_palettes", "import_chart_data",
        "convert_to_pptx", "validate_ppttc", "list_chart_types",
        "diagnose_thinkcell", "insert_into_powerpoint",
        "list_chart_templates", "template_from_slide", "fill_chart_template",
        "update_open_charts",
    ]
    for name in tools:
        check(f"server: tool '{name}' is defined", hasattr(server, name))

    # Remote (--http) mode keeps path arguments inside the allowed folders.
    server.REMOTE = True
    try:
        outside = server.import_chart_data(str(server.PROJECT_ROOT.parent / "secret.xlsx"))
        check("remote: path outside the allowed folders is refused",
              outside["success"] is False and "outside the folders" in outside["errors"][0])
        inside = server.validate_ppttc(str(server.OUTPUT_DIR / "_missing.ppttc"))
        check("remote: path inside output/ is let through", "File not found" in inside["errors"][0])
    finally:
        server.REMOTE = False


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


# --------------------------------------------------------------------------
# table/process through the real server pipeline (write_ppttc_slides bypass)
# --------------------------------------------------------------------------
def test_raw_table_pipeline() -> None:
    import server

    result = server.create_chart(
        chart_type="table",
        data=[{"Dept": "A", "Diff": -5}, {"Dept": "B", "Diff": 5}],
        title="_test_table_pipeline",
        options={"highlight": [{"column": "Diff", "rule": "negative"}]},
        chart_name="TestTable",
    )
    check("create_chart(table): succeeds", result["success"])
    document = json.loads(open(result["ppttc_path"], encoding="utf-8").read())
    elements = {e["name"]: e for e in document[0]["data"]}
    check("create_chart(table): element name matches chart_name",
          "TestTable" in elements)
    check("create_chart(table): negative cell carries a fill",
          "fill" in elements["TestTable"]["table"][1][1])

    # Mixed slide: a normal (add_chart-backed) chart plus a raw-table chart.
    mixed = server.build_presentation(
        slides=[{
            "title": "Mixed",
            "charts": [
                {"chart_type": "bar", "chart_name": "Bars",
                 "data": [{"category": "A", "V": 1}, {"category": "B", "V": 2}]},
                {"chart_type": "table", "chart_name": "Grid",
                 "data": [{"X": "a", "Y": 1}]},
            ],
        }],
        template_path="deck.pptx",
        output_name="_test_mixed_table",
    )
    check("build_presentation: mixed bar+table slide succeeds", mixed["success"])
    mixed_doc = json.loads(open(mixed["ppttc_path"], encoding="utf-8").read())
    names = {e["name"] for e in mixed_doc[0]["data"]}
    check("build_presentation: both bar and table elements present",
          {"Bars", "Grid"} <= names)

    process_result = server.create_chart(
        chart_type="process",
        data=[{"date": "2025-01-01", "label": "Start"},
              {"date": "2026-01-01", "label": "End"}],
        title="_test_process_pipeline",
        chart_name="TestTimeline",
    )
    check("create_chart(process): succeeds", process_result["success"])


# --------------------------------------------------------------------------
# data_import: read chart/table data from .csv / .xlsx
# --------------------------------------------------------------------------
def test_import_chart_data() -> None:
    import tempfile
    from datetime import date as dt_date
    from pathlib import Path as _Path

    from data_import import import_spreadsheet

    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
        tmp_path = _Path(tmp)

        # -- CSV -------------------------------------------------------
        csv_path = tmp_path / "sample.csv"
        csv_path.write_text(
            "Department,Proposed,Difference\n"
            "Agriculture,2217,-453\n"
            "Health,2768,3553\n"
            "\n"  # blank line -- should be skipped
            "Trailing, ,\n",
            encoding="utf-8",
        )
        result = import_spreadsheet(str(csv_path))
        check("import: csv succeeds", result["success"])
        check("import: csv columns detected", result["columns"] == [
            "Department", "Proposed", "Difference"
        ])
        check("import: csv row_count is 3", result["row_count"] == 3)
        check(
            "import: csv numeric cells coerced to numbers",
            result["data"][0] == {
                "Department": "Agriculture", "Proposed": 2217, "Difference": -453
            },
        )
        check(
            "import: csv blank cells dropped rather than kept as ''",
            "Proposed" not in result["data"][2]
            and "Difference" not in result["data"][2],
        )

        # Data feeds straight into a categorical chart builder.
        check(
            "import: csv data validates as a bar chart",
            BarChart(data=result["data"][:2], title="T").validate() == [],
        )
        # ... and straight into the table chart builder.
        check(
            "import: csv data validates as a table chart",
            TableChart(data=result["data"][:2], title="T").validate() == [],
        )

        check(
            "import: missing file reported",
            not import_spreadsheet(str(tmp_path / "nope.csv"))["success"],
        )

        txt_path = tmp_path / "sample.txt"
        txt_path.write_text("x", encoding="utf-8")
        check(
            "import: unsupported extension reported",
            has(import_spreadsheet(str(txt_path))["errors"], "Unsupported file type"),
        )

        dup_path = tmp_path / "dup.csv"
        dup_path.write_text("A,A\n1,2\n", encoding="utf-8")
        check(
            "import: duplicate column headers rejected",
            has(import_spreadsheet(str(dup_path))["errors"], "Duplicate"),
        )

        empty_path = tmp_path / "empty.csv"
        empty_path.write_text("", encoding="utf-8")
        check(
            "import: empty file reported",
            not import_spreadsheet(str(empty_path))["success"],
        )

        header_only = tmp_path / "header_only.csv"
        header_only.write_text("A,B\n", encoding="utf-8")
        header_only_result = import_spreadsheet(str(header_only))
        check(
            "import: header-only csv reports no data rows",
            not header_only_result["success"]
            and has(header_only_result["errors"], "No data rows"),
        )

        # -- XLSX --------------------------------------------------------
        try:
            import openpyxl
        except ImportError:
            print("  SKIP: xlsx import checks (openpyxl unavailable)")
            return

        xlsx_path = tmp_path / "sample.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Budget"
        sheet.append(["Region", "Share", "Date"])
        sheet.append(["North", 0.219, dt_date(2025, 1, 1)])
        sheet.cell(row=2, column=2).number_format = "0.0%"
        sheet.append(["South", 0.5, None])
        workbook.save(xlsx_path)

        xlsx_result = import_spreadsheet(str(xlsx_path), sheet_name="Budget")
        check("import: xlsx succeeds", xlsx_result["success"])
        check(
            "import: xlsx percentage column detected and rescaled to whole points",
            "Share" in xlsx_result["percentage_columns"]
            and abs(xlsx_result["data"][0]["Share"] - 21.9) < 0.01,
        )
        check(
            "import: xlsx date cell converted to ISO string",
            xlsx_result["data"][0]["Date"] == "2025-01-01",
        )
        check(
            "import: xlsx blank cell dropped from its row",
            "Date" not in xlsx_result["data"][1],
        )
        check(
            "import: unknown sheet name reported",
            has(
                import_spreadsheet(str(xlsx_path), sheet_name="Nope")["errors"],
                "not found",
            ),
        )


def test_connector_auth() -> None:
    """Connect-code sign-in: wrong codes refused and limited, tokens hashed and rotated."""
    import asyncio
    import tempfile
    from pathlib import Path

    import oauth
    from mcp.server.auth.provider import AuthorizationParams
    from mcp.shared.auth import OAuthClientInformationFull

    code = oauth.new_connect_code()
    check("connect code looks like XXXX-XXXX", re.fullmatch(r"[A-Z2-9]{4}-[A-Z2-9]{4}", code) is not None)
    check("connect code skips 0/O/1/I/L", not set(code) & set("01OIL"))

    async def flow(state: Path) -> None:
        auth = oauth.ConnectorAuth(state, "https://pc.example/", "k7mp q2xd")
        client = OAuthClientInformationFull(client_id="c1", client_name="Claude",
                                            redirect_uris=["https://claude.ai/cb"])
        await auth.register_client(client)
        params = AuthorizationParams(state="s", scopes=[], code_challenge="x" * 43,
                                     redirect_uri="https://claude.ai/cb",
                                     redirect_uri_provided_explicitly=True)
        url = await auth.authorize(client, params)
        check("authorize sends the user to /connect", url.startswith("https://pc.example/connect?request="))
        req = url.split("request=")[1]
        check("connect page knows the client name", auth.client_name(req) == "Claude")
        redirect, err = auth.approve(req, "WRONG-CODE")
        check("wrong code refused", redirect is None and "not right" in err)
        redirect, err = auth.approve(req, "K7MP-Q2XD")
        check("right code (any spacing/case) approves", err is None and "code=" in redirect and "state=s" in redirect)
        check("a sign-in cannot be reused", auth.approve(req, "K7MP-Q2XD")[0] is None)
        auth_code = await auth.load_authorization_code(client, redirect.split("code=")[1].split("&")[0])
        check("authorization code loads", auth_code is not None)
        token = await auth.exchange_authorization_code(client, auth_code)
        check("access token works", await auth.load_access_token(token.access_token) is not None)
        check("state file holds no raw token", token.access_token not in state.read_text())
        again = oauth.ConnectorAuth(state, "https://pc.example", "K7MP-Q2XD")
        check("token survives a restart", await again.load_access_token(token.access_token) is not None)
        rt = await again.load_refresh_token(client, token.refresh_token)
        new = await again.exchange_refresh_token(client, rt, [])
        check("refresh issues a new token", await again.load_access_token(new.access_token) is not None)
        check("old refresh token refused after use",
              await again.load_refresh_token(client, token.refresh_token) is None)
        await again.revoke_token(await again.load_access_token(new.access_token))
        check("revoked token refused", await again.load_access_token(new.access_token) is None)
        # Overall limit on wrong codes, across sign-ins.
        for _ in range(oauth.FAILURES_PER_WINDOW):
            r = (await auth.authorize(client, params)).split("request=")[1]
            auth.approve(r, "nope")
        r = (await auth.authorize(client, params)).split("request=")[1]
        check("too many wrong codes locks sign-in", "Too many" in auth.approve(r, "K7MP-Q2XD")[1])
        keyed = oauth.ConnectorAuth(state, "https://pc.example", "K7MP-Q2XD", api_key="k-123")
        check("API key is accepted as a bearer token", await keyed.load_access_token("k-123") is not None)
        check("wrong API key refused", await keyed.load_access_token("k-124") is None)
        check("no API key configured -> none accepted", await auth.load_access_token("k-123") is None)
        # Registration and /authorize need no sign-in, so they are capped.
        for i in range(oauth.MAX_CLIENTS + 5):
            await keyed.register_client(OAuthClientInformationFull(client_id=f"x{i}",
                                                                   redirect_uris=["https://claude.ai/cb"]))
        check("client registrations are capped", len(keyed.clients) <= oauth.MAX_CLIENTS)
        check("a signed-in client survives the cap", "c1" in keyed.clients and "x0" not in keyed.clients)
        for _ in range(oauth.MAX_PENDING + 5):
            await keyed.authorize(client, params)
        check("pending sign-ins are capped", len(keyed.pending) <= oauth.MAX_PENDING)
        check("state is written in one step", not state.with_suffix(".tmp").exists())
        state.write_text('{"clients": {"c1"', encoding="utf-8")
        check("an unreadable state file starts empty instead of crashing",
              oauth.ConnectorAuth(state, "https://pc.example", "K7MP-Q2XD").clients == {})

    with tempfile.TemporaryDirectory() as tmp:
        asyncio.run(flow(Path(tmp) / "oauth_state.json"))


def test_connector_app() -> None:
    """The remote app: OAuth metadata, 401 without a token, /connect page, no Content-Length."""
    import asyncio
    import os
    import tempfile

    import httpx

    import server

    with tempfile.TemporaryDirectory() as tmp:
        os.environ["THINKCELL_CONNECTOR_HOME"] = tmp
        import connector

        connector.CONFIG_DIR = connector.Path(tmp)
        connector.CONFIG_FILE = connector.CONFIG_DIR / "config.json"
        connector.STATE_FILE = connector.CONFIG_DIR / "oauth_state.json"
        connector.save_config({"public_url": "https://pc.example", "connect_code": "K7MP-Q2XD",
                               "api_key": "assist-key-123", "folders": [tmp]})
        saved = (server.REMOTE, list(server.REMOTE_ROOTS))
        try:
            app = server.connector_app()
            check("remote mode turns on path limits", server.REMOTE and connector.Path(tmp) in
                  [connector.Path(p) for p in server.REMOTE_ROOTS])

            async def calls() -> None:
                async with httpx.AsyncClient(transport=httpx.ASGITransport(app=app),
                                             base_url="https://pc.example") as c:
                    meta = await c.get("/.well-known/oauth-authorization-server")
                    check("authorization server metadata", meta.status_code == 200
                          and meta.json()["issuer"].rstrip("/") == "https://pc.example")
                    check("replies carry no Content-Length (antivirus fix)",
                          "content-length" not in meta.headers)
                    res = await c.post("/mcp", json={})
                    check("/mcp without a token is 401", res.status_code == 401
                          and "resource_metadata" in res.headers.get("www-authenticate", ""))
                    page = await c.get("/connect?request=missing")
                    check("unknown sign-in request refused", page.status_code == 400)
                    check("connect page cannot be framed", page.headers.get("x-frame-options") == "DENY")

                    # think-cell Assist: browser CORS + API-key header (its manual's CORS table).
                    assist = server.ASSIST_ORIGIN
                    pre = await c.options("/mcp", headers={
                        "Origin": assist, "Access-Control-Request-Method": "POST",
                        "Access-Control-Request-Headers": "content-type,authorization"})
                    allow = pre.headers.get("access-control-allow-headers", "")
                    check("Assist preflight answered before sign-in", pre.status_code == 204)
                    check("preflight allows the Assist origin",
                          pre.headers.get("access-control-allow-origin") == assist)
                    check("preflight allows Assist's required headers", all(
                        h in allow for h in ("content-type", "authorization",
                                             "mcp-protocol-version", "mcp-session-id")))
                    mcp_headers = {"Origin": assist, "Accept": "application/json, text/event-stream",
                                   "MCP-Protocol-Version": "2025-11-25"}
                    init = {"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {
                        "protocolVersion": "2025-11-25", "capabilities": {},
                        "clientInfo": {"name": "assist-test", "version": "1"}}}
                    # The MCP endpoint needs the SDK's session manager running.
                    async with server.mcp.session_manager.run():
                        ok = await c.post("/mcp", json=init, headers={
                            **mcp_headers, "Authorization": "Bearer assist-key-123"})
                    check("Assist API key signs in", ok.status_code == 200 and "protocolVersion" in ok.text)
                    check("reply carries Access-Control-Allow-Origin",
                          ok.headers.get("access-control-allow-origin") == assist)
                    check("session id header exposed to the browser",
                          "mcp-session-id" in ok.headers.get("access-control-expose-headers", ""))
                    bad = await c.post("/mcp", json=init, headers={**mcp_headers, "Authorization": "Bearer nope"})
                    check("wrong API key -> 401, readable by the browser", bad.status_code == 401
                          and bad.headers.get("access-control-allow-origin") == assist)
                    evil = await c.post("/mcp", json=init, headers={**mcp_headers, "Origin": "https://evil.example",
                                                                    "Authorization": "Bearer assist-key-123"})
                    check("other Origin -> 403 even with the key", evil.status_code == 403
                          and "access-control-allow-origin" not in evil.headers)
                    get = await c.get("/mcp", headers={"Origin": assist})
                    check("GET /mcp -> 405 (stateless: no stream to open)", get.status_code == 405)
                    import modern_http
                    huge = await c.post("/mcp", content=b"x" * (modern_http.MAX_BODY + 1), headers={"Origin": assist})
                    check("oversized body -> 413 before sign-in", huge.status_code == 413)
                    own = await c.post("/connect", data={"request": "missing", "code": "x"},
                                       headers={"Origin": "https://pc.example"})
                    check("connect form posts from the server's own origin", own.status_code == 400)

                    # MCP 2026-07-28: stateless, version in _meta + headers, no initialize.
                    async def modern(method, params=None, id_=7, **extra):
                        body = {"jsonrpc": "2.0", "method": method, "params": {
                            **(params or {}), "_meta": {"io.modelcontextprotocol/protocolVersion":
                                                        extra.pop("meta_version", "2026-07-28")}}}
                        if id_ is not None:
                            body["id"] = id_
                        headers = {"Origin": assist, "Authorization": "Bearer assist-key-123",
                                   "MCP-Protocol-Version": "2026-07-28", "Mcp-Method": method, **extra}
                        return await c.post("/mcp", json=body, headers={k: v for k, v in headers.items() if v})
                    disc = await modern("server/discover")
                    result = disc.json().get("result", {})
                    check("2026-07-28 server/discover answered", disc.status_code == 200
                          and "2026-07-28" in result.get("supportedVersions", [])
                          and "2025-11-25" in result["supportedVersions"]
                          and result.get("resultType") == "complete" and "tools" in result.get("capabilities", {}))
                    check("discover names the server",
                          result.get("_meta", {}).get("io.modelcontextprotocol/serverInfo", {}).get("name") == "thinkcell")
                    check("modern reply keeps CORS", disc.headers.get("access-control-allow-origin") == assist)
                    listed = (await modern("tools/list")).json().get("result", {})
                    names = [t["name"] for t in listed.get("tools", [])]
                    check("2026-07-28 tools/list lists every tool, sorted",
                          len(names) == len(await server.mcp.list_tools()) and names == sorted(names)
                          and listed.get("cacheScope") and listed.get("ttlMs"))
                    called = await modern("tools/call", {"name": "list_chart_types", "arguments": {}},
                                          **{"Mcp-Name": "list_chart_types"})
                    out = called.json().get("result", {})
                    check("2026-07-28 tools/call runs the tool", called.status_code == 200
                          and out.get("resultType") == "complete" and not out.get("isError") and out.get("content"))
                    slow = await modern("tools/call", {"name": "fill_chart_template", "arguments": {
                        "template_id": "no-such-template", "charts": {"c1": {}}}}, **{"Mcp-Name": "fill_chart_template"})
                    check("a slow tool answers from its worker thread", slow.status_code == 200
                          and "list_chart_templates" in slow.text)
                    b64 = await modern("tools/call", {"name": "list_chart_types", "arguments": {}},
                                       **{"Mcp-Name": "=?base64?bGlzdF9jaGFydF90eXBlcw==?="})
                    check("base64 Mcp-Name accepted", b64.status_code == 200)
                    wrong_name = await modern("tools/call", {"name": "list_chart_types"}, **{"Mcp-Name": "other"})
                    check("Mcp-Name mismatch -> 400 / -32020", wrong_name.status_code == 400
                          and wrong_name.json()["error"]["code"] == -32020)
                    no_header = await modern("tools/list", **{"MCP-Protocol-Version": ""})
                    check("missing MCP-Protocol-Version -> 400 / -32020", no_header.status_code == 400
                          and no_header.json()["error"]["code"] == -32020)
                    future = await modern("tools/list", meta_version="2099-01-01", **{"MCP-Protocol-Version": "2099-01-01"})
                    err = future.json().get("error", {})
                    check("unknown version -> 400 / -32022 with supported list", future.status_code == 400
                          and err.get("code") == -32022 and "2026-07-28" in err.get("data", {}).get("supported", []))
                    unknown = await modern("prompts/list")
                    check("unsupported method -> 404 / -32601", unknown.status_code == 404
                          and unknown.json()["error"]["code"] == -32601)
                    note = await modern("notifications/cancelled", id_=None)
                    check("modern notification -> 202", note.status_code == 202)
                    anon = await modern("server/discover", Authorization="")
                    check("modern request without a token -> 401", anon.status_code == 401
                          and "resource_metadata" in anon.headers.get("www-authenticate", ""))

            asyncio.run(calls())
        finally:
            server.REMOTE, server.REMOTE_ROOTS[:] = saved[0], saved[1]
            os.environ.pop("THINKCELL_CONNECTOR_HOME", None)


def test_chart_library() -> None:
    """Template library: datasheet tables always; cutting, naming and filling when think-cell is installed."""
    import tempfile
    from pathlib import Path

    import library
    import server

    table = library.chart_table({"categories": ["2025", "2026-01-01"], "series": [
        {"name": "A", "values": [1, 2]}, {"values": [3, None], "fill": "#ff0000"}]})
    check("library: header row has a blank corner and the categories", table[0] == [
        None, {"string": "2025"}, {"date": "2026-01-01"}])
    check("library: series rows carry label and numbers", table[1] == [{"string": "A"}, {"number": 1}, {"number": 2}])
    check("library: unnamed series and fill colour", table[2] == [None, {"number": 3, "fill": "#ff0000"}, None])
    pct = library.chart_table({"categories": ["x"], "series": [{"name": "s", "values": [45]}], "percent": True})
    check("library: percent writes percentage cells", pct[1][1] == {"percentage": 45})
    total = library.chart_table({"categories": ["a", "b"], "total": 100, "series": [{"name": "s", "values": [1, 2]}]})
    check("library: total adds the 100%= row second", total[1] == [None, {"number": 100}, {"number": 100}])
    check("library: waterfall 'e' stays a text cell",
          library.chart_table({"categories": ["a"], "series": [{"values": ["e"]}]})[1][1] == {"string": "e"})
    for label, bad in (("value count", {"categories": ["a", "b"], "series": [{"values": [1]}]}),
                       ("no categories", {"categories": [], "series": [{"values": []}]}),
                       ("total count", {"categories": ["a", "b"], "total": [1], "series": [{"values": [1, 2]}]})):
        try:
            library.chart_table(bad)
            check(f"library: rejects wrong {label}", False)
        except ValueError:
            check(f"library: rejects wrong {label}", True)
    check("library: raw table passes through", library.chart_table({"table": [[None]]}) == [[None]])
    gap = library.chart_table({"categories": ["a"], "series": [{"values": [1]}, None, {"values": [2]}]})
    check("library: a null series is an empty row (keeps the next series' colours)", gap[2] == [] and len(gap) == 4)
    from converter import find_ppttc_exe
    with tempfile.TemporaryDirectory() as tmp:
        (Path(tmp) / "ppttc").mkdir()
        (Path(tmp) / "ppttc" / "ppttc.exe").write_bytes(b"")
        check("converter: ppttc.exe is also found in the ppttc subfolder",
              find_ppttc_exe(tmp) == str(Path(tmp) / "ppttc" / "ppttc.exe"))
        (Path(tmp) / "ppttc.exe").write_bytes(b"")
        check("converter: the install root wins when both exist", find_ppttc_exe(tmp) == str(Path(tmp) / "ppttc.exe"))
    check("library: update_open_charts needs data", not server.update_open_charts(1, {})["success"])
    unknown = server.fill_chart_template("no-such-template", {"c1": {}})
    check("library: unknown template id refused", not unknown["success"] and "list_chart_templates" in unknown["errors"][0])

    # Saved templates live beside the outputs; ones that came with the package are still found.
    import tempfile
    from pathlib import Path
    saved_dirs = (library.USER_DIR, library.BUNDLED_USER_DIR)
    with tempfile.TemporaryDirectory() as tmp:
        library.USER_DIR, library.BUNDLED_USER_DIR = Path(tmp) / "mine", Path(tmp) / "bundled"
        try:
            for folder, title in ((library.BUNDLED_USER_DIR, "old"), (library.USER_DIR, "new")):
                folder.mkdir()
                (folder / "t.json").write_text(json.dumps({"id": "user/t", "title": title}), encoding="utf-8")
            (library.BUNDLED_USER_DIR / "only.pptx").write_bytes(b"")
            check("library: a saved template replaces a bundled one of the same name",
                  [e["title"] for e in library.user_templates()] == ["new"])
            check("library: a bundled user template is still found",
                  library.template_path("user/only") == library.BUNDLED_USER_DIR / "only.pptx")
        finally:
            library.USER_DIR, library.BUNDLED_USER_DIR = saved_dirs

    # Cutting and naming need think-cell's installed library (absent in CI) and pywin32.
    try:
        import pythoncom  # noqa: F401
    except ImportError:
        return
    if not library.INSTALL_LIBRARY.is_dir():
        return
    index = library.ensure_library()
    ids = {e["id"] for e in index}
    check("library: think-cell's chart library becomes 30+ templates", len(index) >= 30)
    check("library: Gantt charts are left out (ppttc garbles them)", not any("gantt" in i for i in ids))
    bar = next(e for e in index if e["id"] == "bar-column-1")
    check("library: bar template reads its categories and series",
          bar["charts"][0]["categories"][:2] == ["2016", "2017"] and bar["charts"][0]["series"] == 3)
    check("library: demo notes beside the slide are dropped", not any(t.startswith("Keywords:") for t in bar["texts"]))
    from pptx import Presentation
    named = [c["name"] for c in library.charts(Presentation(str(library.template_path("bar-column-5"))).slides[0])]
    check("library: several charts named c1..c4 in reading order", named == ["c1", "c2", "c3", "c4"])
    mekko = next(e for e in index if e["id"] == "mekko-2")
    check("library: a datasheet with the 100%= row is detected", mekko["charts"][0].get("total") == 11.0)
    check("library: one without it is not", "total" not in bar["charts"][0])
    bad_name = library.fill_template(library.template_path("bar-column-1"), {"zz": {"categories": ["a"],
                                                                                    "series": [{"values": [1]}]}})
    check("library: unknown chart name refused", not bad_name["success"] and "c1" in bad_name["errors"][0])
    bad_fill = library.fill_template(library.template_path("bar-column-1"), {"c1": {
        "categories": ["a"], "series": [{"values": [1], "fill": "red"}]}})
    check("library: a bad fill colour is refused before ppttc.exe runs",
          not bad_fill["success"] and "fill must be" in bad_fill["errors"][0])
    check("library: think-cell's library charts do not use datasheet fills",
          library.charts(Presentation(str(library.template_path("bar-column-1"))).slides[0])[0]["datasheet_fill"] is False)
    from pathlib import Path
    from converter import PPTTC_EXE
    if not Path(PPTTC_EXE).is_file():
        return
    filled = server.fill_chart_template("waterfall-1", {"c1": {"categories": ["Rev", "Cost", "Profit"],
                                                               "series": [{"name": "m", "values": [10, -4, "e"]}]}},
                                        title="Bridge", texts={"Insert chart title here": "USD m"},
                                        output_name="test_library_fill")
    check("library: fill runs ppttc.exe and the chart shows the new categories",
          filled["success"] and filled["filled"] == ["c1"])
    coloured = server.fill_chart_template("bar-column-1", {"c1": {"categories": ["A", "B"], "series": [
        {"name": "s", "values": [1, 2], "fill": "#ff0000"}]}}, output_name="test_library_fill_colour")
    check("library: a fill the template cannot show gives a warning",
          coloured["success"] and any("Use Datasheet Fill" in w for w in coloured["warnings"]))
    if filled["success"]:
        slide = Presentation(filled["pptx_path"]).slides[0]
        texts = [sh.text_frame.text for sh in library.text_shapes(slide)]
        check("library: title and texts replaced", "Bridge" in texts and "USD m" in texts)


def main() -> int:
    suites = [
        test_registry,
        test_bar,
        test_missing_series_is_rejected,
        test_min_rows,
        test_category_key_option,
        test_scatter,
        test_combo,
        test_table,
        test_process,
        test_known_option_keys,
        test_validator,
        test_validator_percentage,
        test_autodeck,
        test_layout,
        test_branding,
        test_palettes,
        test_autodeck_semantic,
        test_server_smoke,
        test_build_presentation,
        test_raw_table_pipeline,
        test_import_chart_data,
        test_connector_auth,
        test_connector_app,
        test_chart_library,
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
