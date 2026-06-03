"""Read-only quality harness for the think-cell MCP server (autoresearch).

This file is the *ground-truth judge* for the autonomous improvement loop
described in ``program.md``. It is the analogue of ``prepare.py`` /
``evaluate_bpb`` in karpathy/autoresearch: a fixed, self-contained scorer that
the optimizing agent is **not allowed to edit**. The agent edits the library
code (``charts/``, ``autodeck.py``, ``validator.py``) to raise the score.

Design rules that keep the metric honest:

* The validity gate is **independent** of the code under test. We serialise
  every produced ``.ppttc`` with ``json.dumps(..., allow_nan=False)`` and treat
  a failure as an invalid document. The agent cannot game this by weakening
  ``validator.py`` -- non-finite numbers (``NaN``/``Infinity``) are not legal
  JSON and break think-cell's ``ppttc.exe`` downstream, so they must be fixed
  at the source, in the builders.
* Structural checks (row widths, series counts) are computed here, not borrowed
  from ``validator.py``.
* Headroom dimensions map to *real* MCP defects: non-finite values that produce
  broken ``.ppttc`` files, and two documented silent gotchas -- percentages
  passed as decimals, and a date-axis chart with a single category (renders
  empty). Every point of headroom corresponds to a genuine improvement.

Run with::

    py eval.py

Prints a human-readable breakdown and, on the last lines, two machine-readable
summary fields the loop parses::

    score: <achieved> / <max>
    loc:   <editable-lines-of-code>
"""
from __future__ import annotations

import json
import math
import re
import tempfile
from pathlib import Path
from typing import Any, Callable

from autodeck import build_auto_deck
from charts import CHART_BUILDERS, get_builder, write_ppttc_slides
from validator import validate_ppttc_data

MODULE_DIR = Path(__file__).resolve().parent

# Files the optimizing agent is allowed to edit. Their size is the "cost"
# column in results.tsv (mirrors autoresearch tracking peak VRAM): all else
# equal, fewer lines is better.
EDITABLE_FILES: tuple[str, ...] = (
    "charts/base.py",
    "charts/waterfall.py",
    "charts/bar.py",
    "charts/stacked_bar.py",
    "charts/line.py",
    "charts/area.py",
    "charts/scatter.py",
    "charts/mekko.py",
    "charts/combo.py",
    "autodeck.py",
    "validator.py",
    "palettes.py",
)

INF = float("inf")
NAN = float("nan")


# --------------------------------------------------------------------------
# scoring scaffold
# --------------------------------------------------------------------------
class Score:
    """Accumulates points and prints a per-check breakdown."""

    def __init__(self) -> None:
        self.earned = 0.0
        self.possible = 0.0
        self._fails: list[str] = []

    def add(self, label: str, ok: bool, points: float) -> None:
        self.possible += points
        if ok:
            self.earned += points
        else:
            self._fails.append(f"  MISS ({points:g} pts): {label}")

    def report(self) -> None:
        for line in self._fails:
            print(line)
        if not self._fails:
            print("  (all checks earned full points)")


def _safe(fn: Callable[[], Any]) -> tuple[Any, Exception | None]:
    """Run ``fn`` and capture any exception instead of crashing the harness."""
    try:
        return fn(), None
    except Exception as exc:  # noqa: BLE001 - the harness must never crash
        return None, exc


def _builder_to_ppttc(builder: Any) -> dict[str, Any]:
    """Convert a builder's (categories, series) into a one-chart .ppttc doc.

    Mirrors what ``thinkcell.add_chart`` does: a header row of category
    strings preceded by a null corner cell, then one row per series whose
    first cell is the series name and the rest are numeric cells.
    """
    categories, series = builder.build_table()
    header: list[Any] = [None] + [{"string": str(c)} for c in categories]
    table: list[list[Any]] = [header]
    for row in series:
        cells: list[Any] = [{"string": str(row[0])}]
        for value in row[1:]:
            cells.append({"number": value})
        table.append(cells)
    return {
        "template": "deck.pptx",
        "data": [{"name": builder.chart_name, "table": table}],
    }


def _json_finite(document: Any) -> bool:
    """True iff the document serialises as strict JSON (no NaN/Infinity)."""
    try:
        json.dumps(document, allow_nan=False)
        return True
    except (ValueError, TypeError):
        return False


# --------------------------------------------------------------------------
# Section A -- core build & validity (anchors the score)
# --------------------------------------------------------------------------
def section_core(score: Score) -> None:
    """Every chart type must accept its own example and emit valid JSON."""
    for name, builder_cls in CHART_BUILDERS.items():
        example = builder_cls.example
        data = example.get("data", [])
        title = example.get("title", "Example")
        options = {
            k: v
            for k, v in example.items()
            if k not in ("chart_type", "title", "data")
        }

        builder, exc = _safe(lambda: builder_cls(data, title, **options))
        score.add(f"[{name}] construct builder", builder is not None, 1)
        if builder is None:
            continue

        errors, exc = _safe(builder.validate)
        score.add(
            f"[{name}] accepts its own example (validate == [])",
            errors == [],
            2,
        )

        doc, exc = _safe(lambda: _builder_to_ppttc(builder))
        score.add(f"[{name}] build_table succeeds", doc is not None, 2)
        if doc is None:
            continue

        score.add(
            f"[{name}] emits strict-JSON .ppttc (finite numbers)",
            _json_finite(doc),
            3,
        )

        # Structural correctness, computed independently of validator.py.
        table = doc["data"][0]["table"]
        widths = {len(r) for r in table if r}
        score.add(
            f"[{name}] all rows share one width",
            len(widths) == 1,
            2,
        )


# --------------------------------------------------------------------------
# Section B -- non-finite robustness in the chart builders (HEADROOM)
# --------------------------------------------------------------------------
def _bar_with(value: Any) -> Any:
    builder_cls = get_builder("bar")
    data = [
        {"category": "A", "Revenue": 100},
        {"category": "B", "Revenue": value},
    ]
    return builder_cls(data, "Robustness")


def section_builder_robustness(score: Score) -> None:
    """NaN/Infinity must never reach the output: reject, or emit finite JSON."""
    for label, bad in (("NaN", NAN), ("Infinity", INF), ("-Infinity", -INF)):
        builder, _ = _safe(lambda: _bar_with(bad))
        if builder is None:
            score.add(f"builder handles {label}", False, 5)
            continue

        errors, _ = _safe(builder.validate)
        if errors:
            # Rejected up front -- the clean, preferred outcome.
            score.add(f"builder rejects {label} value", True, 5)
            continue

        # Not rejected: then the produced document MUST be finite JSON.
        doc, _ = _safe(lambda: _builder_to_ppttc(builder))
        ok = doc is not None and _json_finite(doc)
        score.add(f"builder rejects {label} value", ok, 5)


# --------------------------------------------------------------------------
# Section C -- non-finite robustness in autodeck (HEADROOM)
# --------------------------------------------------------------------------
def _auto_left(values: list[Any], categories: list[str] | None = None) -> dict:
    cats = categories or ["2024-01-01", "2024-02-01", "2024-03-01"]
    return {
        "slide_title": "Robustness",
        "left_chart": {
            "categories": cats,
            "series": [{"name": "Share", "values": values}],
        },
    }


def section_autodeck_robustness(score: Score) -> None:
    """build_auto_deck must reject non-finite series values, not write them."""
    for label, bad in (("NaN", NAN), ("Infinity", INF)):
        slide = _auto_left([10.0, bad, 30.0])
        result, exc = _safe(lambda: build_auto_deck([slide], "eval_robust"))
        if result is None:
            score.add(f"autodeck handles {label}", False, 5)
            continue
        # Either rejected, or (if accepted) the written file must be finite.
        rejected = result.get("success") is False
        finite_ok = True
        path = result.get("ppttc_path")
        if not rejected and path and Path(path).is_file():
            finite_ok = _json_finite(json.loads(Path(path).read_text("utf-8")))
        score.add(
            f"autodeck rejects {label} series value",
            rejected or finite_ok,
            5,
        )


# --------------------------------------------------------------------------
# Section D -- documented-gotcha guardrails in autodeck (HEADROOM)
# --------------------------------------------------------------------------
def section_autodeck_guardrails(score: Score) -> None:
    """Warn on the two silent failures documented in the project's learnings."""
    # D1: percentages passed as decimals (every |value| < 1) -- the #1 silent
    # bug. A healthy build should still succeed but surface a warning.
    slide = _auto_left([0.21, 0.18, 0.25])
    result, _ = _safe(lambda: build_auto_deck([slide], "eval_decimal"))
    warnings = (result or {}).get("warnings") or []
    flagged = any(
        "percent" in w.lower() or "decimal" in w.lower() or "scale" in w.lower()
        for w in warnings
    )
    score.add("autodeck warns on decimal-scale percentages", flagged, 6)

    # D2: a date-axis chart with a single category renders empty (no bars).
    slide = _auto_left([42.0], categories=["2024-01-01"])
    result, _ = _safe(lambda: build_auto_deck([slide], "eval_single"))
    warnings = (result or {}).get("warnings") or []
    flagged = any(
        "categor" in w.lower() or "single" in w.lower() or "empty" in w.lower()
        for w in warnings
    )
    score.add("autodeck warns on single-category date axis", flagged, 4)


# --------------------------------------------------------------------------
# Section E -- validator hardening (HEADROOM)
# --------------------------------------------------------------------------
def section_validator(score: Score) -> None:
    """validate_ppttc_data must treat non-finite numeric cells as invalid."""
    # Good document stays valid.
    good = {
        "template": "deck.pptx",
        "data": [{"name": "C1", "table": [[None, {"string": "A"}],
                                          [{"string": "S"}, {"number": 1}]]}],
    }
    verdict, _ = _safe(lambda: validate_ppttc_data([good]))
    score.add("validator accepts a clean document", bool(verdict and verdict["valid"]), 3)

    # NaN / Infinity numeric cells must be flagged invalid.
    for label, bad in (("NaN", NAN), ("Infinity", INF)):
        doc = {
            "template": "deck.pptx",
            "data": [{"name": "C1", "table": [[None, {"string": "A"}],
                                              [{"string": "S"}, {"number": bad}]]}],
        }
        verdict, _ = _safe(lambda: validate_ppttc_data([doc]))
        invalid = bool(verdict) and verdict["valid"] is False
        score.add(f"validator rejects {label} numeric cell", invalid, 4)


# --------------------------------------------------------------------------
# Section F -- build_presentation round-trip coverage (regression guard)
# --------------------------------------------------------------------------
def _build_one(builder: Any) -> dict[str, Any]:
    """Round-trip one builder through the real thinkcell library and reload."""
    tmp = Path(tempfile.gettempdir()) / "eval_bp.ppttc"
    out = write_ppttc_slides("deck.pptx", [([builder], [])], tmp)
    return json.loads(Path(out).read_text(encoding="utf-8"))


def section_build_presentation(score: Score) -> None:
    """Every chart type must survive the multi-chart build path used by
    ``build_presentation`` and yield a finite, structurally valid .ppttc."""
    for name, builder_cls in CHART_BUILDERS.items():
        example = builder_cls.example
        data = example.get("data", [])
        options = {
            k: v
            for k, v in example.items()
            if k not in ("chart_type", "title", "data")
        }
        builder, _ = _safe(
            lambda: builder_cls(data, example.get("title", "X"),
                                chart_name="Chart1", **options)
        )
        if builder is None:
            score.add(f"[{name}] build_presentation round-trip", False, 4)
            continue

        doc, exc = _safe(lambda: _build_one(builder))
        score.add(f"[{name}] write_ppttc_slides succeeds", doc is not None, 1)
        if doc is None:
            continue
        score.add(f"[{name}] round-trip is finite JSON", _json_finite(doc), 2)
        verdict, _ = _safe(lambda: validate_ppttc_data(doc))
        score.add(
            f"[{name}] round-trip validates",
            bool(verdict) and verdict["valid"],
            1,
        )

    # A three-slide presentation must yield exactly three .ppttc entries.
    bar = get_builder("bar")
    slides = [
        ([bar([{"category": "A", "V": 1}, {"category": "B", "V": 2}],
              "S", chart_name="Chart1")], [])
        for _ in range(3)
    ]
    tmp = Path(tempfile.gettempdir()) / "eval_multi.ppttc"
    doc, _ = _safe(
        lambda: json.loads(
            Path(write_ppttc_slides("deck.pptx", slides, tmp))
            .read_text(encoding="utf-8")
        )
    )
    score.add(
        "3-slide presentation yields 3 .ppttc entries",
        isinstance(doc, list) and len(doc) == 3,
        3,
    )


# --------------------------------------------------------------------------
# Section G -- scatter bubble-size sanity (HEADROOM)
# --------------------------------------------------------------------------
def section_scatter(score: Score) -> None:
    """A bubble's size drives its area, so a negative size is meaningless."""
    cls = get_builder("scatter")
    bad, _ = _safe(
        lambda: cls([{"label": "A", "x": 1, "y": 2, "size": -5}], "neg")
    )
    errors, _ = _safe(bad.validate) if bad else ([], None)
    score.add("scatter rejects a negative bubble size", bool(errors), 5)

    # A valid positive-size bubble must still be accepted.
    good, _ = _safe(
        lambda: cls([{"label": "A", "x": 1, "y": 2, "size": 5}], "pos")
    )
    ok_errs, _ = _safe(good.validate) if good else (["x"], None)
    score.add("scatter still accepts a positive bubble size", ok_errs == [], 2)


# --------------------------------------------------------------------------
# Section H -- mekko column-width sanity (HEADROOM)
# --------------------------------------------------------------------------
def section_mekko(score: Score) -> None:
    """A Mekko 'Width' series sets column widths; zero/negative is degenerate."""
    cls = get_builder("mekko")
    bad, _ = _safe(
        lambda: cls(
            [{"category": "A", "V": 10, "Width": 0},
             {"category": "B", "V": 20, "Width": -3}],
            "bad",
        )
    )
    errors, _ = _safe(bad.validate) if bad else ([], None)
    score.add("mekko rejects non-positive column widths", bool(errors), 5)

    # A normal Mekko with positive widths must still build cleanly.
    good, _ = _safe(
        lambda: cls(
            [{"category": "A", "V": 10, "Width": 70},
             {"category": "B", "V": 20, "Width": 30}],
            "ok",
        )
    )
    ok_errs, _ = _safe(good.validate) if good else (["x"], None)
    score.add("mekko still accepts positive widths", ok_errs == [], 2)


# --------------------------------------------------------------------------
# Section I -- colour system: named palettes + semantic fills (HEADROOM)
# --------------------------------------------------------------------------
def section_palettes(score: Score) -> None:
    """The Vizro-derived colour system must expose valid palettes and map a
    signed value to the right semantic fill (gain teal / loss red / zero gray).

    These checks are computed here independently: hex validity is verified with
    this file's own regex, and the gain/loss mapping is asserted by sign, not
    borrowed from the module under test.
    """
    mod, exc = _safe(lambda: __import__("palettes"))
    if mod is None:
        for label in (
            "palettes module imports",
            "palettes are valid hex",
            "semantic colours present",
            "semantic_fill colours a gain",
            "semantic_fill colours a loss",
            "semantic_fill neutral on zero/garbage",
            "reverse flips the mapping",
        ):
            score.add(label, False, 2)
        return

    score.add("palettes module imports", True, 1)

    hex_re = re.compile(r"^#[0-9A-Fa-f]{6}$")
    palettes = getattr(mod, "PALETTES", {})
    all_hex = bool(palettes) and all(
        isinstance(c, str) and hex_re.match(c)
        for colors in palettes.values()
        for c in colors
    )
    score.add("every palette colour is #RRGGBB hex", all_hex, 3)

    semantic = getattr(mod, "SEMANTIC", {})
    needed = {"positive", "negative", "neutral", "info", "warning"}
    score.add(
        "five semantic roles are defined",
        needed.issubset(set(semantic)),
        2,
    )
    score.add(
        "semantic colours are all valid hex",
        bool(semantic)
        and all(isinstance(v, str) and hex_re.match(v) for v in semantic.values()),
        2,
    )

    fill = getattr(mod, "semantic_fill", None)
    if not callable(fill):
        score.add("semantic_fill colours a gain by sign", False, 4)
        score.add("semantic_fill colours a loss by sign", False, 4)
        score.add("semantic_fill is neutral on zero and garbage", False, 4)
        score.add("reverse flips the gain/loss mapping", False, 3)
    else:
        gain, _ = _safe(lambda: fill(125.0))
        loss, _ = _safe(lambda: fill(-30.0))
        score.add(
            "semantic_fill colours a gain by sign",
            gain == semantic.get("positive"),
            4,
        )
        score.add(
            "semantic_fill colours a loss by sign",
            loss == semantic.get("negative"),
            4,
        )
        zero, _ = _safe(lambda: fill(0))
        junk, _ = _safe(lambda: fill("x"))
        nan, _ = _safe(lambda: fill(NAN))
        neutral = semantic.get("neutral")
        score.add(
            "semantic_fill is neutral on zero and garbage",
            zero == neutral and junk == neutral and nan == neutral,
            4,
        )
        rev_down, _ = _safe(lambda: fill(-30.0, reverse=True))
        rev_up, _ = _safe(lambda: fill(125.0, reverse=True))
        score.add(
            "reverse flips the gain/loss mapping",
            rev_down == semantic.get("positive")
            and rev_up == semantic.get("negative"),
            3,
        )


# --------------------------------------------------------------------------
# cost / size of the editable surface
# --------------------------------------------------------------------------
def editable_loc() -> int:
    """Count non-blank, non-comment lines across the files under optimization."""
    total = 0
    for rel in EDITABLE_FILES:
        path = MODULE_DIR / rel
        if not path.is_file():
            continue
        for line in path.read_text(encoding="utf-8").splitlines():
            stripped = line.strip()
            if stripped and not stripped.startswith("#"):
                total += 1
    return total


# --------------------------------------------------------------------------
# entry point
# --------------------------------------------------------------------------
def main() -> int:
    score = Score()
    print("== Section A: core build & validity ==")
    section_core(score)
    print("== Section B: builder non-finite robustness ==")
    section_builder_robustness(score)
    print("== Section C: autodeck non-finite robustness ==")
    section_autodeck_robustness(score)
    print("== Section D: autodeck documented-gotcha guardrails ==")
    section_autodeck_guardrails(score)
    print("== Section E: validator hardening ==")
    section_validator(score)
    print("== Section F: build_presentation round-trip coverage ==")
    section_build_presentation(score)
    print("== Section G: scatter bubble-size sanity ==")
    section_scatter(score)
    print("== Section H: mekko column-width sanity ==")
    section_mekko(score)
    print("== Section I: colour system (palettes + semantic fills) ==")
    section_palettes(score)
    score.report()

    loc = editable_loc()
    print()
    print(f"score: {score.earned:g} / {score.possible:g}")
    print(f"loc:   {loc}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
