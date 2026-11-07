"""Table chart builder for think-cell JSON automation.

Unlike the categorical chart types (one category axis + numeric series), a
table is a literal rows x columns grid, so it does not fit the shared
``categorical_table()`` helper. It also needs *per-cell* conditional fill
(e.g. only the negative cells in a "Difference" column turn red) which the
``thinkcell`` library's ``add_chart`` cannot produce (only one uniform fill
per whole series) -- so this builder bypasses it via ``build_raw_table``
(see ``charts/base.py``).
"""
from __future__ import annotations

from typing import Any

from .base import ChartBuilder, Table, is_number

try:
    from palettes import semantic_fill
except ImportError:  # pragma: no cover - palettes.py always ships alongside
    def semantic_fill(value: Any, *, reverse: bool = False) -> str:
        return "#EA5748" if is_number(value) and value < 0 else "#00B5A9"

_HIGHLIGHT_RULES: frozenset[str] = frozenset(
    {"negative", "positive", "gt", "lt", "eq"}
)


class TableChart(ChartBuilder):
    """A plain rows x columns table, e.g. a budget-allocation summary.

    ``options.columns`` (list of strings) sets the header/column order; when
    omitted, columns are inferred from the union of row keys in first-seen
    order. A row missing a key renders as a blank cell rather than an error
    (tables commonly have ragged data). ``options.highlight`` drives
    conditional per-cell fill.
    """

    chart_type = "table"
    data_shape = (
        "List of row objects, one per table row. `options.columns` (list of "
        "strings) sets the header/column order; when omitted, columns are "
        "inferred from the union of keys across rows in first-seen order. A "
        "row missing a key renders as a blank cell rather than an error. "
        "`options.percentage_columns` (list of column names) renders those "
        "columns' numbers as think-cell `percentage` cells instead of plain "
        "`number` cells."
    )
    known_option_keys = frozenset(
        {"columns", "percentage_columns", "highlight"}
    )
    example = {
        "chart_type": "table",
        "title": "Budget allocation",
        "data": [
            {
                "Department": "Agriculture", "Proposed": 2217,
                "Allocation": 1764, "Difference": -453,
            },
            {
                "Department": "Health", "Proposed": 2768,
                "Allocation": 6321, "Difference": 3553,
            },
        ],
        "options": {
            "columns": ["Department", "Proposed", "Allocation", "Difference"],
            "highlight": [{"column": "Difference", "rule": "negative"}],
        },
    }

    # -- columns ----------------------------------------------------------
    def _columns(self) -> list[str]:
        explicit = self.options.get("columns")
        if isinstance(explicit, list) and explicit:
            return [str(c) for c in explicit]
        columns: list[str] = []
        for row in self.data:
            if not isinstance(row, dict):
                continue
            for key in row:
                if key not in columns:
                    columns.append(key)
        return columns

    # -- validation ---------------------------------------------------------
    def _validate(self) -> list[str]:
        errors: list[str] = []
        columns = self._columns()
        if not columns:
            errors.append(
                "table chart needs at least one column: pass "
                "`options.columns` or ensure at least one data row has keys"
            )
            return errors

        pct_columns = self.options.get("percentage_columns")
        if pct_columns is not None:
            if not isinstance(pct_columns, list):
                errors.append(
                    "`options.percentage_columns` must be a list of column "
                    "names"
                )
                pct_columns = []
            else:
                unknown_pct = [c for c in pct_columns if c not in columns]
                if unknown_pct:
                    errors.append(
                        "`options.percentage_columns` references unknown "
                        f"column(s) {unknown_pct}; known columns: {columns}"
                    )

        for index, row in enumerate(self.data):
            unknown_keys = [key for key in row if key not in columns]
            if unknown_keys:
                errors.append(
                    f"data[{index}] has key(s) {unknown_keys} not in "
                    f"columns {columns}"
                )
            for key, value in row.items():
                if value is None or isinstance(value, str):
                    continue
                if isinstance(value, bool) or not isinstance(
                    value, (int, float)
                ):
                    errors.append(
                        f"data[{index}][{key!r}] must be a string, number "
                        f"or null, got {type(value).__name__}"
                    )
                elif not is_number(value):
                    errors.append(
                        f"data[{index}][{key!r}] must be a finite number "
                        f"(NaN/Infinity are not valid JSON), got {value!r}"
                    )

        errors.extend(self._validate_highlight(columns))
        return errors

    def _validate_highlight(self, columns: list[str]) -> list[str]:
        errors: list[str] = []
        rules = self.options.get("highlight")
        if rules is None:
            return errors
        if not isinstance(rules, list):
            return ["`options.highlight` must be a list of highlight rules"]

        for index, rule in enumerate(rules):
            where = f"options.highlight[{index}]"
            if not isinstance(rule, dict):
                errors.append(f"{where} must be an object")
                continue
            column = rule.get("column")
            if column not in columns:
                errors.append(
                    f"{where}.column must be one of {columns}, got {column!r}"
                )
            kind = rule.get("rule")
            if kind not in _HIGHLIGHT_RULES:
                errors.append(
                    f"{where}.rule must be one of {sorted(_HIGHLIGHT_RULES)}, "
                    f"got {kind!r}"
                )
                continue
            if kind in ("gt", "lt", "eq"):
                if not is_number(rule.get("value")):
                    errors.append(
                        f"{where}.value must be a finite number for rule "
                        f"'{kind}'"
                    )
                fill = rule.get("fill")
                if not isinstance(fill, str):
                    errors.append(
                        f"{where}.fill is required (as '#RRGGBB') for rule "
                        f"'{kind}'"
                    )
            else:
                fill = rule.get("fill")
                if fill is not None and not isinstance(fill, str):
                    errors.append(
                        f"{where}.fill must be a '#RRGGBB' string if given"
                    )
        return errors

    # -- highlighting ---------------------------------------------------
    def _cell_fill(self, column: str, value: Any) -> str | None:
        if not is_number(value):
            return None
        for rule in self.options.get("highlight") or []:
            if not isinstance(rule, dict) or rule.get("column") != column:
                continue
            kind = rule.get("rule")
            matched = (
                (kind == "negative" and value < 0)
                or (kind == "positive" and value > 0)
                or (kind == "gt" and value > rule.get("value", value))
                or (kind == "lt" and value < rule.get("value", value))
                or (kind == "eq" and value == rule.get("value"))
            )
            if matched:
                fill = rule.get("fill")
                return fill if isinstance(fill, str) else semantic_fill(value)
        return None

    # -- table construction -----------------------------------------------
    def build_table(self) -> Table:
        raise NotImplementedError(
            "TableChart produces a raw .ppttc table; see build_raw_table()"
        )

    def build_raw_table(self) -> tuple[str, list[list[Any]]]:
        columns = self._columns()
        pct_columns = set(self.options.get("percentage_columns") or [])
        header: list[Any] = [{"string": c} for c in columns]
        rows: list[list[Any]] = [header]
        for row in self.data:
            cells: list[Any] = []
            for column in columns:
                value = row.get(column)
                if value is None:
                    cells.append(None)
                elif isinstance(value, str):
                    cells.append({"string": value})
                else:
                    key = "percentage" if column in pct_columns else "number"
                    cell: dict[str, Any] = {key: value}
                    fill = self._cell_fill(column, value)
                    if fill:
                        cell["fill"] = fill
                    cells.append(cell)
            rows.append(cells)
        return self.chart_name, rows
