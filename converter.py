"""Wrapper around think-cell's ``ppttc.exe`` command-line converter.

``convert_ppttc`` runs the converter and always returns a structured dict --
it never raises for the documented failure modes (missing input file, missing
or broken ``ppttc.exe``, an unusable output path, or a timeout).
"""
from __future__ import annotations

import os
import subprocess
from pathlib import Path
from typing import Any

DEFAULT_THINKCELL_DIR = r"C:\Program Files (x86)\think-cell"


def find_ppttc_exe(folder: str | Path) -> str:
    """ppttc.exe in think-cell's install folder, or in its ``ppttc`` subfolder (where the manual puts it)."""
    root = Path(folder) / "ppttc.exe"
    sub = Path(folder) / "ppttc" / "ppttc.exe"
    return str(sub if sub.is_file() and not root.is_file() else root)


def find_thinkcell_dir() -> str:
    """think-cell's install directory: the THINKCELL_DIR environment variable,
    else the install location Windows records for think-cell (it can be
    installed anywhere, e.g. D:\\thinkcell), else the default folder."""
    if os.environ.get("THINKCELL_DIR"):
        return os.environ["THINKCELL_DIR"]
    try:
        import winreg
    except ImportError:  # not Windows
        return DEFAULT_THINKCELL_DIR
    uninstall = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    for hive, path in ((winreg.HKEY_LOCAL_MACHINE, uninstall),
                       (winreg.HKEY_LOCAL_MACHINE, uninstall.replace("SOFTWARE", r"SOFTWARE\WOW6432Node")),
                       (winreg.HKEY_CURRENT_USER, uninstall)):
        try:
            with winreg.OpenKey(hive, path) as root:
                for index in range(winreg.QueryInfoKey(root)[0]):
                    with winreg.OpenKey(root, winreg.EnumKey(root, index)) as app:
                        try:
                            name = winreg.QueryValueEx(app, "DisplayName")[0]
                            location = winreg.QueryValueEx(app, "InstallLocation")[0]
                        except OSError:
                            continue
                        if name == "think-cell" and location and Path(find_ppttc_exe(location)).is_file():
                            return location.rstrip("\\")
        except OSError:
            continue
    return DEFAULT_THINKCELL_DIR


THINKCELL_DIR = find_thinkcell_dir()

# Default location of the think-cell command-line converter.
PPTTC_EXE = find_ppttc_exe(THINKCELL_DIR)

# Conversion drives PowerPoint under the hood, so allow a generous timeout.
CONVERT_TIMEOUT_SECONDS = 180


def _failure(error: str, **extra: Any) -> dict[str, Any]:
    """Build a structured failure result with a stable shape."""
    return {"success": False, "pptx_path": None, "error": error, **extra}


def convert_ppttc(
    ppttc_path: str,
    output_path: str | None = None,
    ppttc_exe: str = PPTTC_EXE,
) -> dict[str, Any]:
    """Convert a ``.ppttc`` file to ``.pptx`` using ``ppttc.exe``.

    Args:
        ppttc_path: Path to an existing ``.ppttc`` file.
        output_path: Optional destination ``.pptx`` path. Defaults to the
            input path with a ``.pptx`` extension.
        ppttc_exe: Path to ``ppttc.exe``. Defaults to the standard think-cell
            install location.

    Returns:
        On success: ``{"success": True, "pptx_path": str, "stdout": str,
        "stderr": str, "error": None}``. On failure: ``{"success": False,
        "pptx_path": None, "error": str, ...}`` with extra diagnostic keys.
    """
    # 1. The input .ppttc file must exist.
    input_path = Path(ppttc_path).expanduser()
    if not input_path.is_file():
        return _failure(f"Input .ppttc file not found: {ppttc_path}")
    input_path = input_path.resolve()

    # 2. ppttc.exe must be installed where we expect it.
    exe_path = Path(ppttc_exe)
    if not exe_path.is_file():
        return _failure(
            f"ppttc.exe not found at '{ppttc_exe}'. Install think-cell or "
            f"pass a correct ppttc_exe path."
        )

    # 3. Resolve and sanity-check the output path.
    if output_path:
        out_path = Path(output_path).expanduser()
    else:
        out_path = input_path.with_suffix(".pptx")
    out_path = out_path.resolve()

    if out_path.suffix.lower() != ".pptx":
        return _failure(f"Output path must end in .pptx, got '{out_path}'")

    out_dir = out_path.parent
    if not out_dir.exists():
        try:
            out_dir.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            return _failure(f"Cannot create output directory '{out_dir}': {exc}")

    # 4. Run the converter, capturing stdout/stderr.
    command = [str(exe_path), str(input_path), "-o", str(out_path)]
    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            errors="replace",
            timeout=CONVERT_TIMEOUT_SECONDS,
        )
    except subprocess.TimeoutExpired:
        return _failure(
            f"ppttc.exe timed out after {CONVERT_TIMEOUT_SECONDS} seconds",
            command=" ".join(command),
        )
    except OSError as exc:
        return _failure(f"Failed to launch ppttc.exe: {exc}")

    stdout = (result.stdout or "").strip()
    stderr = (result.stderr or "").strip()

    # 5. Surface a non-zero exit code as a clear, structured error.
    if result.returncode != 0:
        return _failure(
            f"ppttc.exe failed with exit code {result.returncode}",
            returncode=result.returncode,
            stdout=stdout,
            stderr=stderr,
            command=" ".join(command),
        )

    # 6. A zero exit code with no output file is still a failure.
    if not out_path.is_file():
        return _failure(
            "ppttc.exe reported success but no .pptx file was produced",
            stdout=stdout,
            stderr=stderr,
        )

    return {
        "success": True,
        "pptx_path": str(out_path),
        "stdout": stdout,
        "stderr": stderr,
        "error": None,
    }
