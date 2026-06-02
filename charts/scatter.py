"""Scatter / XY (and bubble) chart builder for think-cell JSON automation."""
from __future__ import annotations

from .base import ChartBuilder, Table, as_number, is_number


class ScatterChart(ChartBuilder):
    """An XY scatter chart. Each row is a single point with numeric x and y.

    Add an optional numeric ``size`` key to the points to drive a bubble
    chart instead. An optional ``label`` (or ``name``) string names the point.
    """

    chart_type = "scatter"
    data_shape = (
        "List of point objects. Each point requires numeric 'x' and 'y' "
        "keys, and may include an optional numeric 'size' (turns the chart "
        "into a bubble chart) and an optional 'label' string naming the point."
    )
    example = {
        "chart_type": "scatter",
        "title": "Price vs Demand",
        "data": [
            {"label": "Product A", "x": 9.99, "y": 1200},
            {"label": "Product B", "x": 14.99, "y": 850},
            {"label": "Product C", "x": 19.99, "y": 540},
        ],
    }

    # Scatter manages its own x/y layout, so it accepts no extra options.
    known_option_keys = frozenset()

    def _validate(self) -> list[str]:
        errors: list[str] = []
        for index, row in enumerate(self.data):
            if "x" not in row or "y" not in row:
                errors.append(
                    f"data[{index}] must include numeric 'x' and 'y' keys"
                )
                continue
            if not is_number(row["x"]):
                errors.append(
                    f"data[{index}]['x'] must be a number, "
                    f"got {type(row['x']).__name__}"
                )
            if not is_number(row["y"]):
                errors.append(
                    f"data[{index}]['y'] must be a number, "
                    f"got {type(row['y']).__name__}"
                )
            if "size" in row:
                if not is_number(row["size"]):
                    errors.append(
                        f"data[{index}]['size'] must be a number, "
                        f"got {type(row['size']).__name__}"
                    )
                elif row["size"] < 0:
                    errors.append(
                        f"data[{index}]['size'] must be non-negative; "
                        f"a bubble cannot have a negative area "
                        f"(got {row['size']})"
                    )
        return errors

    def build_table(self) -> Table:
        has_size = any("size" in row for row in self.data)
        categories: list = ["X", "Y", "Size"] if has_size else ["X", "Y"]
        series: list[list] = []
        for index, row in enumerate(self.data):
            label = row.get("label") or row.get("name") or f"Point {index + 1}"
            point: list = [str(label), as_number(row["x"]), as_number(row["y"])]
            if has_size:
                point.append(as_number(row.get("size", 0)))
            series.append(point)
        return categories, series
