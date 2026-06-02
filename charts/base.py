"""Shared foundation for think-cell chart builders.

Every supported chart type lives in its own module and subclasses
:class:`ChartBuilder`. A builder accepts user-friendly input (a list of row
objects), validates it for that specific chart type, and produces the
``(categories, series)`` pair that think-cell's JSON data automation expects.

The think-cell ``.ppttc`` format produced here (via the ``thinkcell`` library)
looks like::

    [
      {
        "template": "deck.pptx",
        "data": [
          {"name": "Chart1", "table": [[null, ...categories], [], ...series]}
        ]
      }
    ]
"""
from __future__ import annotations

import math
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any, ClassVar

from thinkcell import Thinkcell

# A think-cell table is the (categories, series) pair fed to Thinkcell.add_chart:
#   categories -> list of column headers
#   series     -> list of rows, each row is [series_name, value, value, ...]
Table = tuple[list[Any], list[list[Any]]]

# Row keys that are treated as the category/label column rather than a numeric
# data series, checked in priority order. The first match wins. ("x" is
# intentionally excluded -- scatter handles its own x/y layout, and treating a
# series literally named "x" as a category would silently misread the data.)
CATEGORY_KEYS: tuple[str, ...] = ("category", "label", "name")


class ChartError(Exception):
    """Raised when otherwise-valid input cannot be assembled into a .ppttc."""


def is_number(value: Any) -> bool:
    """Return True for real, finite numeric values.

    ``bool`` is intentionally excluded. ``NaN`` and ``+/-Infinity`` are also
    excluded: they are not valid JSON and would produce a ``.ppttc`` that
    think-cell's ``ppttc.exe`` rejects, so they must never reach the output.
    """
    return (
        isinstance(value, (int, float))
        and not isinstance(value, bool)
        and math.isfinite(value)
    )


def as_number(value: Any) -> float | int:
    """Coerce a value to a finite int/float, falling back to 0 otherwise."""
    if is_number(value):
        return value
    try:
        number = float(value)
    except (TypeError, ValueError):
        return 0
    if not math.isfinite(number):
        return 0
    return int(number) if number.is_integer() else number


class ChartBuilder(ABC):
    """Base class for all chart builders.

    Subclasses set the ``chart_type``, ``data_shape`` and ``example`` class
    attributes, implement :meth:`_validate` and :meth:`build_table`, and may
    reuse the categorical helpers below.
    """

    # -- catalog metadata (overridden per chart type) -------------------
    chart_type: ClassVar[str] = ""
    required_params: ClassVar[list[str]] = ["chart_type", "data", "title"]
    optional_params: ClassVar[list[str]] = [
        "template_path", "chart_name", "options",
    ]
    data_shape: ClassVar[str] = ""
    example: ClassVar[dict[str, Any]] = {}

    # Option keys accepted via **options. Used to reject typos in chart
    # configs. Every categorical builder understands "category_key"; chart
    # types with extra options widen this set.
    known_option_keys: ClassVar[frozenset[str]] = frozenset({"category_key"})

    def __init__(
        self,
        data: list[dict[str, Any]],
        title: str,
        chart_name: str = "Chart1",
        **options: Any,
    ) -> None:
        self.data = data
        self.title = title
        self.chart_name = chart_name or "Chart1"
        self.options = options

    # -- validation -----------------------------------------------------
    def validate(self) -> list[str]:
        """Run generic then chart-specific validation; return error strings."""
        if not isinstance(self.data, list) or len(self.data) == 0:
            return ["`data` must be a non-empty list of row objects"]

        generic: list[str] = []
        for index, row in enumerate(self.data):
            if not isinstance(row, dict):
                generic.append(
                    f"data[{index}] must be an object, "
                    f"got {type(row).__name__}"
                )
        if generic:
            return generic

        return self._validate()

    @abstractmethod
    def _validate(self) -> list[str]:
        """Chart-type-specific validation (runs after generic checks pass)."""

    # -- table construction --------------------------------------------
    @abstractmethod
    def build_table(self) -> Table:
        """Return the ``(categories, series)`` pair for thinkcell.add_chart."""

    # -- shared helpers -------------------------------------------------
    def category_key(self) -> str:
        """Identify which row key holds the category/label for each row."""
        explicit = self.options.get("category_key")
        if isinstance(explicit, str) and explicit:
            return explicit
        first_row = self.data[0]
        for candidate in CATEGORY_KEYS:
            if candidate in first_row:
                return candidate
        return next(iter(first_row), "category")

    def series_names(self, category_key: str) -> list[str]:
        """Collect every non-category key across all rows, preserving order."""
        names: list[str] = []
        for row in self.data:
            for key in row:
                if key != category_key and key not in names:
                    names.append(key)
        return names

    def categorical_table(self) -> Table:
        """Build a standard categories-across / series-down think-cell table."""
        category_key = self.category_key()
        categories = [str(row.get(category_key, "")) for row in self.data]
        series: list[list[Any]] = []
        for name in self.series_names(category_key):
            values: list[Any] = [name]
            for row in self.data:
                values.append(as_number(row.get(name, 0)))
            series.append(values)
        return categories, series

    def validate_categorical(
        self, *, min_categories: int = 1, min_series: int = 1
    ) -> list[str]:
        """Validate the common categorical row shape used by most charts."""
        errors: list[str] = []
        category_key = self.category_key()

        if len(self.data) < min_categories:
            errors.append(
                f"{self.chart_type} chart needs at least {min_categories} "
                f"data row(s) (categories); got {len(self.data)}"
            )

        names = self.series_names(category_key)
        if len(names) < min_series:
            errors.append(
                f"{self.chart_type} chart needs at least {min_series} data "
                f"series (row keys other than '{category_key}'); "
                f"got {len(names)}"
            )

        for index, row in enumerate(self.data):
            if category_key not in row:
                errors.append(
                    f"data[{index}] is missing the category key "
                    f"'{category_key}'"
                )
            for name in names:
                if name not in row:
                    errors.append(
                        f"data[{index}] is missing series '{name}'; every "
                        f"row must define the same series (a missing value "
                        f"would otherwise be charted as 0, not as a gap)"
                    )
                elif not is_number(row[name]):
                    errors.append(
                        f"data[{index}]['{name}'] must be a number, "
                        f"got {type(row[name]).__name__}"
                    )
        return errors


# One slide of a presentation: the chart builders on it plus any
# (field_name, text) text-field pairs.
Slide = tuple[list[ChartBuilder], list[tuple[str, str]]]


def write_ppttc_slides(
    template: str,
    slides: list[Slide],
    output_path: str | Path,
) -> str:
    """Assemble a multi-slide ``.ppttc`` file via the thinkcell library.

    Each item in ``slides`` becomes one top-level ``.ppttc`` template entry,
    and think-cell turns each entry into exactly one PowerPoint slide. (The
    thinkcell library appends every chart to the most recently added template
    entry, so one ``add_template`` call per slide is what separates them.)

    Args:
        template: Path/name of the think-cell PowerPoint template (.pptx).
        slides: One ``(builders, textfields)`` pair per slide, in order.
        output_path: Destination path for the ``.ppttc`` file.

    Returns:
        The absolute path to the written ``.ppttc`` file.

    Raises:
        ChartError: If the think-cell library rejects the data or template.
        OSError: If the file cannot be written.
    """
    document = Thinkcell()
    try:
        for builders, textfields in slides:
            document.add_template(template)
            for builder in builders:
                categories, series = builder.build_table()
                document.add_chart(
                    template, builder.chart_name, categories, series
                )
            for field_name, text in textfields or []:
                document.add_textfield(template, field_name, text)
    except (ValueError, TypeError) as exc:
        raise ChartError(str(exc)) from exc

    out_path = Path(output_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    document.save_ppttc(str(out_path))
    return str(out_path.resolve())


def write_ppttc_document(
    template: str,
    builders: list[ChartBuilder],
    output_path: str | Path,
    textfields: list[tuple[str, str]] | None = None,
) -> str:
    """Write a single-slide ``.ppttc`` file (one template entry).

    Convenience wrapper over :func:`write_ppttc_slides` for the one-slide
    case. Raises ``ChartError``/``OSError`` exactly as that function does.
    """
    return write_ppttc_slides(
        template, [(builders, textfields or [])], output_path
    )
