"""Process / timeline (roadmap) chart builder for think-cell JSON automation.

Renders a dated milestone roadmap -- e.g. "Roadmap launched" -> "Review by
CM" -> "Launch of v2.0" laid out along a horizontal date axis, optionally
with a "we are here" marker. Like ``table``, this needs per-cell fill (to
highlight the current-date marker) and mixes ``date``/``string`` cells in
ways the ``thinkcell`` library's ``add_chart`` cannot produce, so it bypasses
that library via ``build_raw_table`` (see ``charts/base.py``).
"""
from __future__ import annotations

import re
from typing import Any

from .base import ChartBuilder, Table

_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
_FILL_RE = re.compile(r"^#[0-9A-Fa-f]{6}$")


class ProcessChart(ChartBuilder):
    """A dated milestone timeline/roadmap.

    Each row is one milestone: an ISO ``date`` plus a ``label``, and an
    optional ``detail`` line. Produces a table with dates as the header row
    and one text row per populated field (``Milestone``, and ``Detail`` if
    any milestone supplies one). ``options.current_date`` marks a "we are
    here" point -- it is inserted chronologically among the milestones (even
    if it falls between two of them) and highlighted with a fill colour.
    """

    chart_type = "process"
    data_shape = (
        "Ordered list of milestone objects: {'date': 'YYYY-MM-DD', 'label': "
        "str, 'detail': str (optional)}. At least two milestones are "
        "required. `options.current_date` (an ISO date, need not match a "
        "milestone) marks a 'we are here' point; `options.current_fill` "
        "overrides its default marker colour."
    )
    known_option_keys = frozenset({"current_date", "current_fill"})
    example = {
        "chart_type": "process",
        "title": "Roadmap",
        "data": [
            {
                "date": "2025-05-27",
                "label": "Roadmap Finalized & Launched",
                "detail": "by Chief Minister",
            },
            {"date": "2025-06-12", "label": "Review by CM"},
            {"date": "2026-01-19", "label": "Launch of DSD 2.0"},
        ],
        "options": {"current_date": "2026-07-06"},
    }

    # -- validation -----------------------------------------------------
    def _validate(self) -> list[str]:
        errors: list[str] = []
        if len(self.data) < 2:
            errors.append(
                f"process chart needs at least 2 milestones; got "
                f"{len(self.data)}"
            )

        for index, row in enumerate(self.data):
            date = row.get("date")
            if not isinstance(date, str) or not _DATE_RE.match(date):
                errors.append(
                    f"data[{index}].date must be an ISO date "
                    f"'YYYY-MM-DD', got {date!r}"
                )
            label = row.get("label")
            if not isinstance(label, str) or not label:
                errors.append(
                    f"data[{index}].label must be a non-empty string"
                )
            detail = row.get("detail")
            if detail is not None and not isinstance(detail, str):
                errors.append(f"data[{index}].detail must be a string if given")

        current_date = self.options.get("current_date")
        if current_date is not None and (
            not isinstance(current_date, str) or not _DATE_RE.match(current_date)
        ):
            errors.append(
                "`options.current_date` must be an ISO date 'YYYY-MM-DD'"
            )

        current_fill = self.options.get("current_fill")
        if current_fill is not None and (
            not isinstance(current_fill, str) or not _FILL_RE.match(current_fill)
        ):
            errors.append(
                "`options.current_fill` must be a hex colour like '#8C0A23'"
            )
        return errors

    # -- table construction -----------------------------------------------
    def build_table(self) -> Table:
        raise NotImplementedError(
            "ProcessChart produces a raw .ppttc table; see build_raw_table()"
        )

    def build_raw_table(self) -> tuple[str, list[list[Any]]]:
        current_date = self.options.get("current_date")
        current_fill = self.options.get("current_fill") or "#8C0A23"

        # Merge milestones with the current-date marker (if it doesn't
        # coincide with one) and lay everything out in chronological order,
        # so "we are here" can fall between two milestones.
        points: dict[str, dict[str, Any]] = {}
        for row in self.data:
            points.setdefault(
                row["date"], {"label": row["label"], "detail": row.get("detail")}
            )
        if current_date and current_date not in points:
            points[current_date] = {"label": None, "detail": None}

        header: list[Any] = []
        labels: list[Any] = [{"string": "Milestone"}]
        details: list[Any] = [{"string": "Detail"}]
        has_detail = False
        for date in sorted(points):
            cell: dict[str, Any] = {"date": date}
            if current_date and date == current_date:
                cell["fill"] = current_fill
            header.append(cell)

            entry = points[date]
            labels.append({"string": entry["label"]} if entry["label"] else None)
            if entry["detail"]:
                has_detail = True
                details.append({"string": entry["detail"]})
            else:
                details.append(None)

        table = [header, labels]
        if has_detail:
            table.append(details)
        return self.chart_name, table
