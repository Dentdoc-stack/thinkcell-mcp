"""Marimekko (Mekko) chart builder for think-cell JSON automation."""
from __future__ import annotations

from .base import ChartBuilder, Table, is_number

# The optional series name that drives per-category column widths. A Mekko's
# horizontal axis is a proportion of total width, so a width must be strictly
# positive -- a zero/negative width collapses the column and silently distorts
# every other column's share.
WIDTH_KEY = "Width"


class MekkoChart(ChartBuilder):
    """A Marimekko chart: stacked bars whose column widths vary by category.

    Segment heights come from the numeric series, exactly like a stacked bar.
    Column widths are configured on the think-cell Mekko inside the PowerPoint
    template (commonly driven by a dedicated 'Width' series in the data).
    """

    chart_type = "mekko"
    data_shape = (
        "List of row objects. Each row has one category key ('category', "
        "'label' or 'name') and one or more numeric series keys for the "
        "stacked segments. Optionally include a 'Width' series to drive the "
        "per-category column widths."
    )
    example = {
        "chart_type": "mekko",
        "title": "Market Share by Segment",
        "data": [
            {"category": "Segment A", "Company X": 60, "Company Y": 40, "Width": 70},
            {"category": "Segment B", "Company X": 35, "Company Y": 65, "Width": 30},
        ],
    }

    def _validate(self) -> list[str]:
        errors = self.validate_categorical(min_categories=1, min_series=1)
        for index, row in enumerate(self.data):
            if WIDTH_KEY in row and is_number(row[WIDTH_KEY]):
                if row[WIDTH_KEY] <= 0:
                    errors.append(
                        f"data[{index}]['{WIDTH_KEY}'] must be greater than 0; "
                        f"a column width sets the category's share of the axis "
                        f"and cannot be zero or negative (got {row[WIDTH_KEY]})"
                    )
        return errors

    def build_table(self) -> Table:
        return self.categorical_table()
