"""Chart builders for the think-cell MCP server.

Each supported chart type has a dedicated module and a :class:`ChartBuilder`
subclass. ``CHART_BUILDERS`` maps the public chart-type name to its builder
class; the helpers below are what ``server.py`` consumes.
"""
from __future__ import annotations

from typing import Any

from .area import AreaChart
from .bar import BarChart
from .base import (
    ChartBuilder,
    ChartError,
    Table,
    write_ppttc_document,
    write_ppttc_slides,
)
from .combo import ComboChart
from .line import LineChart
from .mekko import MekkoChart
from .process import ProcessChart
from .scatter import ScatterChart
from .stacked_bar import StackedBarChart
from .table import TableChart
from .waterfall import WaterfallChart

# Public chart-type name -> builder class. Order is preserved for catalogs.
CHART_BUILDERS: dict[str, type[ChartBuilder]] = {
    "waterfall": WaterfallChart,
    "bar": BarChart,
    "stacked_bar": StackedBarChart,
    "line": LineChart,
    "scatter": ScatterChart,
    "mekko": MekkoChart,
    "area": AreaChart,
    "combo": ComboChart,
    "table": TableChart,
    "process": ProcessChart,
}

__all__ = [
    "CHART_BUILDERS",
    "ChartBuilder",
    "ChartError",
    "Table",
    "write_ppttc_document",
    "write_ppttc_slides",
    "get_builder",
    "chart_type_catalog",
]


def get_builder(chart_type: str) -> type[ChartBuilder] | None:
    """Return the builder class for ``chart_type``, or None if unsupported."""
    return CHART_BUILDERS.get(chart_type)


def chart_type_catalog() -> dict[str, dict[str, Any]]:
    """Describe every supported chart type (used by the list_chart_types tool)."""
    catalog: dict[str, dict[str, Any]] = {}
    for name, builder in CHART_BUILDERS.items():
        catalog[name] = {
            "required_params": list(builder.required_params),
            "optional_params": list(builder.optional_params),
            "data_shape": builder.data_shape,
            "example": builder.example,
        }
    return catalog
