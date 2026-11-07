"""Read chart/table data from a local spreadsheet (.csv or .xlsx).

The first row is treated as column headers; every following row becomes one
row object with those headers as keys. That is exactly the ``data`` shape
every categorical chart builder (``create_chart``/``build_presentation``)
and the ``table`` chart type already expect -- so a spreadsheet can be
imported straight into either without reshaping.
"""
from __future__ import annotations

import csv
import datetime as dt
from pathlib import Path
from typing import Any


def _coerce_csv_value(raw: str) -> Any:
    """Best-effort type coercion for a CSV cell (everything starts as str).

    An empty cell becomes ``None`` (dropped from the row -- most builders
    treat a missing key differently from an explicit 0/empty string). A cell
    that looks numeric becomes an int or float; everything else stays a
    string.
    """
    value = raw.strip()
    if value == "":
        return None
    try:
        as_float = float(value)
    except ValueError:
        return value
    return int(as_float) if as_float.is_integer() else as_float


def _read_csv(path: Path) -> tuple[list[str], list[dict[str, Any]], list[str]]:
    errors: list[str] = []
    with path.open(newline="", encoding="utf-8-sig") as handle:
        reader = csv.reader(handle)
        try:
            header = next(reader)
        except StopIteration:
            return [], [], ["File is empty: no header row found"]

        columns = [column.strip() for column in header]
        rows: list[dict[str, Any]] = []
        for line_number, raw_row in enumerate(reader, start=2):
            if not any(cell.strip() for cell in raw_row):
                continue  # skip fully blank lines
            row: dict[str, Any] = {}
            for column, cell in zip(columns, raw_row):
                value = _coerce_csv_value(cell)
                if value is not None:
                    row[column] = value
            if len(raw_row) != len(columns):
                errors.append(
                    f"row {line_number} has {len(raw_row)} cell(s) but the "
                    f"header has {len(columns)}; extra/missing cells ignored"
                )
            rows.append(row)
    return columns, rows, errors


def _read_xlsx(
    path: Path, sheet_name: str
) -> tuple[list[str], list[dict[str, Any]], list[str], list[str]]:
    try:
        import openpyxl
    except ImportError:
        return [], [], [
            "Reading .xlsx files needs the 'openpyxl' package: "
            "python -m pip install openpyxl"
        ], []

    try:
        workbook = openpyxl.load_workbook(
            path, read_only=True, data_only=True
        )
    except Exception as exc:  # noqa: BLE001 - surface any openpyxl failure
        return [], [], [f"Could not read the workbook: {exc}"], []

    if sheet_name:
        if sheet_name not in workbook.sheetnames:
            return [], [], [
                f"Sheet '{sheet_name}' not found; available sheets: "
                f"{workbook.sheetnames}"
            ], []
        sheet = workbook[sheet_name]
    else:
        sheet = workbook.active

    rows_iter = sheet.iter_rows()
    try:
        header_cells = next(rows_iter)
    except StopIteration:
        return [], [], ["Sheet is empty: no header row found"], []

    columns = [
        str(cell.value).strip() if cell.value is not None else f"Column{i + 1}"
        for i, cell in enumerate(header_cells)
    ]

    # A column is treated as a percentage if its first populated data cell
    # is percentage-formatted in Excel -- those store the fraction (0.219
    # for 21.9%), so the value is rescaled to whole points (21.9) to match
    # what think-cell's `percentage` cells expect, and reported back so the
    # caller can also pass it as `options.percentage_columns` to `table`.
    percentage_columns: set[str] = set()
    column_format_known: set[str] = set()

    rows: list[dict[str, Any]] = []
    for raw_row in rows_iter:
        if all(cell.value is None for cell in raw_row):
            continue  # skip fully blank rows
        row: dict[str, Any] = {}
        for index, cell in enumerate(raw_row):
            if index >= len(columns):
                break
            column = columns[index]
            value = cell.value
            if value is None:
                continue
            if isinstance(value, (dt.datetime, dt.date)):
                value = value.strftime("%Y-%m-%d")
            elif isinstance(value, (int, float)) and not isinstance(value, bool):
                if column not in column_format_known:
                    column_format_known.add(column)
                    if "%" in (cell.number_format or ""):
                        percentage_columns.add(column)
                if column in percentage_columns:
                    value = value * 100
            row[column] = value
        rows.append(row)

    workbook.close()
    return columns, rows, [], sorted(percentage_columns)


def import_spreadsheet(file_path: str, sheet_name: str = "") -> dict[str, Any]:
    """Read a local ``.csv``/``.xlsx`` file into chart/table-ready row data.

    Args:
        file_path: Path to a local ``.csv`` or ``.xlsx``/``.xlsm`` file.
        sheet_name: For ``.xlsx``, the worksheet to read. Defaults to the
            active sheet. Ignored for ``.csv``.

    Returns:
        ``{"success": bool, "data": list[dict] | None, "columns": list[str],
        "row_count": int, "percentage_columns": list[str], "errors": list[str]}``.
        ``data`` is directly usable as the ``data`` argument to
        ``create_chart``/``build_presentation`` (for any categorical chart
        type) or as the ``table`` chart type's ``data``. ``percentage_columns``
        (``.xlsx`` only) lists columns Excel formatted as a percentage --
        pass it straight through as the ``table`` chart's
        ``options.percentage_columns``.
    """
    path = Path(file_path).expanduser()
    if not path.is_file():
        return {
            "success": False, "data": None, "columns": [], "row_count": 0,
            "percentage_columns": [], "errors": [f"File not found: {file_path}"],
        }

    suffix = path.suffix.lower()
    if suffix == ".csv":
        columns, rows, errors = _read_csv(path)
        percentage_columns: list[str] = []
    elif suffix in (".xlsx", ".xlsm"):
        columns, rows, errors, percentage_columns = _read_xlsx(path, sheet_name)
    else:
        return {
            "success": False, "data": None, "columns": [], "row_count": 0,
            "percentage_columns": [], "errors": [
                f"Unsupported file type '{suffix}'; expected .csv, .xlsx or .xlsm"
            ],
        }

    if not columns:
        return {
            "success": False, "data": None, "columns": [], "row_count": 0,
            "percentage_columns": [], "errors": errors or ["No columns found"],
        }

    duplicates = sorted({c for c in columns if columns.count(c) > 1})
    if duplicates:
        return {
            "success": False, "data": None, "columns": columns, "row_count": 0,
            "percentage_columns": [], "errors": [
                f"Duplicate column header(s) {duplicates}; every column "
                "needs a unique name to become a row-object key"
            ],
        }

    if not rows:
        errors.append("No data rows found (only a header row)")

    return {
        "success": len(rows) > 0,
        "data": rows or None,
        "columns": columns,
        "row_count": len(rows),
        "percentage_columns": percentage_columns,
        "errors": errors,
    }
