"""think-cell environment diagnostics.

think-cell does not expose a general API to "control" its PowerPoint add-in --
its only supported automation is the ``.ppttc`` / ``ppttc.exe`` data pipeline.
When that pipeline does nothing, the cause is almost always environmental:
the add-in is not registered, PowerPoint disabled it after a crash, the
licence lapsed, or the template has no think-cell charts.

``run_diagnostics`` performs **read-only** Windows registry and filesystem
checks (no processes are launched) and returns a structured report that
pinpoints the problem. It uses the standard library only.
"""
from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path
from typing import Any, Callable

# winreg is Windows-only. Import lazily so the module still imports (and the
# rest of the server still works) on other platforms.
try:
    import winreg
except ImportError:  # pragma: no cover - non-Windows
    winreg = None  # type: ignore[assignment]

# think-cell install location, found the same way as converter.py (the
# THINKCELL_DIR environment variable, else the registry, else the default).
from converter import find_ppttc_exe, find_thinkcell_dir

THINKCELL_DIR = Path(find_thinkcell_dir())
PPTTC_EXE = Path(find_ppttc_exe(THINKCELL_DIR))

# PowerPoint add-in LoadBehavior values and what they mean.
LOAD_BEHAVIOR = {
    0: "Disabled - the add-in will not load",
    1: "Loaded, but not configured to load at startup",
    2: "Configured to load at startup, but not currently loaded",
    3: "Active - loads at startup and is currently connected",
    8: "Loaded on demand",
    9: "Loaded on demand and currently connected",
    16: "Connects once, then loads on demand",
}

# Office major versions to probe for the Resiliency (disabled-items) keys.
OFFICE_VERSIONS = ("16.0", "15.0", "14.0")

# Statuses ranked by severity for the overall summary.
_SEVERITY = {"ok": 0, "info": 0, "warning": 1, "error": 2}


# --------------------------------------------------------------------------
# registry helpers (all return safely when a key/value is missing)
# --------------------------------------------------------------------------
def _reg_subkeys(hive: int, path: str) -> list[str]:
    """Return the subkey names under ``hive\\path`` (empty if it is missing)."""
    if winreg is None:
        return []
    names: list[str] = []
    try:
        key = winreg.OpenKey(hive, path, 0, winreg.KEY_READ)
    except OSError:
        return names
    try:
        index = 0
        while True:
            try:
                names.append(winreg.EnumKey(key, index))
            except OSError:
                break
            index += 1
    finally:
        winreg.CloseKey(key)
    return names


def _reg_values(hive: int, path: str) -> dict[str, Any]:
    """Return ``{name: value}`` for every value under ``hive\\path``."""
    values: dict[str, Any] = {}
    if winreg is None:
        return values
    try:
        key = winreg.OpenKey(hive, path, 0, winreg.KEY_READ)
    except OSError:
        return values
    try:
        index = 0
        while True:
            try:
                name, value, _vtype = winreg.EnumValue(key, index)
            except OSError:
                break
            values[name] = value
            index += 1
    finally:
        winreg.CloseKey(key)
    return values


def _reg_key_exists(hive: int, path: str) -> bool:
    """Return True if ``hive\\path`` can be opened for reading."""
    if winreg is None:
        return False
    try:
        key = winreg.OpenKey(hive, path, 0, winreg.KEY_READ)
    except OSError:
        return False
    winreg.CloseKey(key)
    return True


def _reg_default(hive: int, path: str) -> str | None:
    """Return the default (unnamed) string value of ``hive\\path``."""
    if winreg is None:
        return None
    try:
        return str(winreg.QueryValue(hive, path))
    except OSError:
        return None


def _safe(value: Any) -> Any:
    """Coerce a registry value into something JSON-serializable."""
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, bytes):
        return f"<binary, {len(value)} bytes>"
    return str(value)


def _mentions_thinkcell(value: Any) -> bool:
    """Return True if a registry value (text or binary blob) names think-cell."""
    texts: list[str] = []
    if isinstance(value, bytes):
        for encoding in ("utf-16-le", "utf-8", "latin-1"):
            texts.append(value.decode(encoding, errors="ignore"))
    elif isinstance(value, str):
        texts.append(value)
    elif isinstance(value, (list, tuple)):
        texts.extend(str(item) for item in value)
    for text in texts:
        lowered = text.lower()
        if "thinkcell" in lowered or "think-cell" in lowered:
            return True
    return False


# --------------------------------------------------------------------------
# individual checks -- each returns
#   {"status", "summary", "details", "recommendation"}
# --------------------------------------------------------------------------
def _check_installation() -> dict[str, Any]:
    """Check that think-cell is installed and ppttc.exe is present."""
    dir_found = THINKCELL_DIR.is_dir()
    ppttc_found = PPTTC_EXE.is_file()

    binaries: list[str] = []
    if dir_found:
        try:
            binaries = sorted(
                item.name
                for item in THINKCELL_DIR.iterdir()
                if item.is_file() and item.suffix.lower() in (".exe", ".dll")
            )
        except OSError:
            binaries = []

    product = None
    uninstall_bases = [
        (winreg.HKEY_LOCAL_MACHINE if winreg else 0,
         r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_LOCAL_MACHINE if winreg else 0,
         r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_CURRENT_USER if winreg else 0,
         r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
    ]
    for hive, base in uninstall_bases:
        if product is not None:
            break
        for sub in _reg_subkeys(hive, base):
            values = _reg_values(hive, f"{base}\\{sub}")
            name = str(values.get("DisplayName", ""))
            if "think-cell" in name.lower() or "thinkcell" in name.lower():
                product = {
                    "display_name": name,
                    "version": _safe(values.get("DisplayVersion")),
                    "publisher": _safe(values.get("Publisher")),
                    "install_location": _safe(values.get("InstallLocation")),
                }
                break

    details = {
        "install_directory": str(THINKCELL_DIR),
        "install_directory_found": dir_found,
        "ppttc_exe": str(PPTTC_EXE),
        "ppttc_exe_found": ppttc_found,
        "binaries_in_install_dir": binaries,
        "registered_product": product,
    }

    if not dir_found:
        return {
            "status": "error",
            "summary": "think-cell is not installed at the expected location.",
            "details": details,
            "recommendation": "Install think-cell from think-cell.com. "
            "Without it there is no ppttc.exe and no PowerPoint add-in.",
        }
    if not ppttc_found:
        return {
            "status": "error",
            "summary": "The think-cell folder exists but ppttc.exe is missing.",
            "details": details,
            "recommendation": "Run the think-cell installer's Repair option; "
            "ppttc.exe is required by the convert_to_pptx tool.",
        }
    version = product["version"] if product else None
    return {
        "status": "ok",
        "summary": "think-cell is installed"
        + (f" (version {version})" if version else "")
        + " and ppttc.exe is present.",
        "details": details,
        "recommendation": None,
    }


def _check_powerpoint() -> dict[str, Any]:
    """Check that Microsoft PowerPoint (the add-in's host) is installed."""
    installed = _reg_key_exists(winreg.HKEY_CLASSES_ROOT, "PowerPoint.Application")
    curver = _reg_default(winreg.HKEY_CLASSES_ROOT, r"PowerPoint.Application\CurVer")
    details = {"powerpoint_installed": installed, "progid_curver": curver}

    if not installed:
        return {
            "status": "error",
            "summary": "Microsoft PowerPoint does not appear to be installed.",
            "details": details,
            "recommendation": "Install desktop Microsoft PowerPoint. think-cell "
            "is a PowerPoint add-in and cannot run without it.",
        }
    return {
        "status": "ok",
        "summary": "Microsoft PowerPoint is installed.",
        "details": details,
        "recommendation": None,
    }


def _check_com_addin() -> dict[str, Any]:
    """Check whether the think-cell COM add-in is registered with PowerPoint."""
    hives = [
        ("HKLM", winreg.HKEY_LOCAL_MACHINE),
        ("HKCU", winreg.HKEY_CURRENT_USER),
    ]
    views = [
        ("64-bit view", r"SOFTWARE\Microsoft\Office\PowerPoint\Addins"),
        ("32-bit view", r"SOFTWARE\WOW6432Node\Microsoft\Office\PowerPoint\Addins"),
    ]

    thinkcell_entries: list[dict[str, Any]] = []
    total_addins = 0
    for hive_label, hive in hives:
        for view_label, base in views:
            for progid in _reg_subkeys(hive, base):
                total_addins += 1
                values = _reg_values(hive, f"{base}\\{progid}")
                friendly = str(values.get("FriendlyName", ""))
                if "think" not in progid.lower() and "think" not in friendly.lower():
                    continue
                behavior = values.get("LoadBehavior")
                behavior = behavior if isinstance(behavior, int) else None
                thinkcell_entries.append({
                    "progid": progid,
                    "location": f"{hive_label} ({view_label})",
                    "friendly_name": friendly or None,
                    "description": str(values.get("Description", "")) or None,
                    "load_behavior": behavior,
                    "load_behavior_meaning": LOAD_BEHAVIOR.get(
                        behavior, f"Unrecognized value ({behavior})"
                    ),
                })

    details = {
        "thinkcell_addin_registrations": thinkcell_entries,
        "total_powerpoint_addins_registered": total_addins,
    }

    if not thinkcell_entries:
        return {
            "status": "error",
            "summary": "The think-cell COM add-in is NOT registered with "
            "PowerPoint.",
            "details": details,
            "recommendation": "Re-run the think-cell installer (use Repair). "
            "Until PowerPoint knows about the add-in, it cannot apply .ppttc "
            "data and ppttc.exe conversions will not work.",
        }

    behaviors = [
        entry["load_behavior"]
        for entry in thinkcell_entries
        if isinstance(entry["load_behavior"], int)
    ]
    if 3 in behaviors:
        return {
            "status": "ok",
            "summary": "The think-cell COM add-in is registered and active "
            "(it loads with PowerPoint).",
            "details": details,
            "recommendation": None,
        }
    if 2 in behaviors:
        return {
            "status": "warning",
            "summary": "The think-cell add-in is set to load at startup but is "
            "not currently marked as loaded.",
            "details": details,
            "recommendation": "Open PowerPoint once so the add-in connects; "
            "its LoadBehavior should then settle at 3 (active).",
        }
    if 0 in behaviors:
        return {
            "status": "error",
            "summary": "The think-cell add-in is registered but DISABLED "
            "(LoadBehavior 0).",
            "details": details,
            "recommendation": "In PowerPoint: File > Options > Add-ins > set "
            "'Manage' to 'COM Add-ins' > Go..., then tick 'think-cell'.",
        }
    return {
        "status": "warning",
        "summary": "The think-cell add-in is registered but has an unusual "
        "load state.",
        "details": details,
        "recommendation": "In PowerPoint: File > Options > Add-ins > set "
        "'Manage' to 'COM Add-ins' > Go..., and ensure 'think-cell' is ticked.",
    }


def _check_disabled_items() -> dict[str, Any]:
    """Check whether PowerPoint has disabled think-cell after a crash/hang."""
    resiliency_keys: list[str] = []
    hits: list[str] = []
    for version in OFFICE_VERSIONS:
        for leaf in ("DisabledItems", "CrashingAddinList"):
            path = (
                rf"SOFTWARE\Microsoft\Office\{version}\PowerPoint"
                rf"\Resiliency\{leaf}"
            )
            if not _reg_key_exists(winreg.HKEY_CURRENT_USER, path):
                continue
            resiliency_keys.append(f"HKCU\\{path}")
            for name, value in _reg_values(
                winreg.HKEY_CURRENT_USER, path
            ).items():
                if _mentions_thinkcell(value):
                    hits.append(f"HKCU\\{path} -> '{name or '(default)'}'")

    details = {
        "resiliency_keys_present": resiliency_keys,
        "entries_referencing_thinkcell": hits,
    }
    if hits:
        return {
            "status": "error",
            "summary": "PowerPoint has put think-cell on its disabled / "
            "crashing add-in list.",
            "details": details,
            "recommendation": "In PowerPoint: File > Options > Add-ins > set "
            "'Manage' to 'Disabled Items' > Go..., select think-cell and click "
            "Enable, then restart PowerPoint.",
        }
    return {
        "status": "ok",
        "summary": "PowerPoint has not disabled the think-cell add-in.",
        "details": details,
        "recommendation": None,
    }


def _check_license() -> dict[str, Any]:
    """Report think-cell registry presence (licence state is not readable)."""
    locations = [
        ("HKCU", winreg.HKEY_CURRENT_USER, r"SOFTWARE\think-cell"),
        ("HKLM", winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\think-cell"),
        ("HKLM", winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\think-cell"),
    ]
    hives_found: list[dict[str, Any]] = []
    for label, hive, path in locations:
        if not _reg_key_exists(hive, path):
            continue
        values = _reg_values(hive, path)
        hives_found.append({
            "location": f"{label}\\{path}",
            # Value NAMES only -- never dump values, they may hold a licence key.
            "value_names": sorted(str(n) for n in values if n),
            "subkeys": _reg_subkeys(hive, path),
        })

    details = {"thinkcell_registry_hives": hives_found}
    if not hives_found:
        return {
            "status": "info",
            "summary": "No think-cell registry data was found for this user; "
            "think-cell may never have been launched on this account.",
            "details": details,
            "recommendation": "Open PowerPoint once and confirm think-cell "
            "loads, then check the licence in the think-cell toolbar > More "
            "(gear icon) > 'About think-cell'.",
        }
    return {
        "status": "info",
        "summary": "think-cell registry data is present. Licence and "
        "activation state cannot be read reliably from the registry.",
        "details": details,
        "recommendation": "Confirm the licence inside PowerPoint: think-cell "
        "toolbar > More (gear icon) > 'About think-cell'. An expired licence "
        "stops automation from working.",
    }


def _check_template(template_path: str) -> dict[str, Any]:
    """Inspect a .pptx template for valid structure and think-cell content."""
    path = Path(template_path).expanduser()
    details: dict[str, Any] = {"path": str(path)}

    if not path.exists():
        return {
            "status": "error",
            "summary": f"Template file not found: {template_path}",
            "details": details,
            "recommendation": "Provide the path to an existing .pptx "
            "think-cell template.",
        }

    details["size_bytes"] = path.stat().st_size
    if path.suffix.lower() != ".pptx":
        return {
            "status": "error",
            "summary": f"Template is not a .pptx file: {path.name}",
            "details": details,
            "recommendation": "think-cell templates must be PowerPoint .pptx "
            "files.",
        }
    if not zipfile.is_zipfile(path):
        return {
            "status": "error",
            "summary": "File is not a valid PowerPoint (OOXML/.pptx) package.",
            "details": details,
            "recommendation": "Re-save the template from PowerPoint as .pptx.",
        }

    marker_parts: list[str] = []
    shape_names: set[str] = set()
    slide_parts: list[str] = []
    has_presentation = False
    try:
        with zipfile.ZipFile(path) as archive:
            names = archive.namelist()
            has_presentation = "ppt/presentation.xml" in names
            slide_parts = sorted(
                name
                for name in names
                if name.startswith("ppt/slides/slide")
                and name.endswith(".xml")
            )
            for entry in names:
                lowered_name = entry.lower()
                if "thinkcell" in lowered_name or "think-cell" in lowered_name:
                    marker_parts.append(entry)
                if not (lowered_name.endswith(".xml")
                        or lowered_name.endswith(".rels")):
                    continue
                try:
                    text = archive.read(entry).decode("utf-8", errors="ignore")
                except (KeyError, OSError):
                    continue
                lowered_text = text.lower()
                if (("thinkcell" in lowered_text or "think-cell" in lowered_text)
                        and entry not in marker_parts):
                    marker_parts.append(entry)
                if entry in slide_parts:
                    for match in re.findall(
                        r'<p:cNvPr\b[^>]*?\bname="([^"]*)"', text
                    ):
                        if match:
                            shape_names.add(match)
    except zipfile.BadZipFile:
        return {
            "status": "error",
            "summary": "The .pptx package is corrupt and cannot be read.",
            "details": details,
            "recommendation": "Re-save the template from PowerPoint.",
        }

    details["valid_pptx"] = has_presentation
    details["slide_count"] = len(slide_parts)
    details["thinkcell_markers_found"] = sorted(marker_parts)
    details["shape_names_on_slides"] = sorted(shape_names)
    details["note"] = (
        "Shape names are a best-effort hint. The names a .ppttc file targets "
        "are the think-cell element names assigned inside PowerPoint, which "
        "may differ from PowerPoint shape names."
    )

    if not has_presentation:
        return {
            "status": "error",
            "summary": "File is a zip archive but not a PowerPoint "
            "presentation (no ppt/presentation.xml).",
            "details": details,
            "recommendation": "Provide a genuine PowerPoint .pptx file.",
        }
    if not marker_parts:
        return {
            "status": "warning",
            "summary": "The .pptx opened cleanly but no think-cell content "
            "was detected. .ppttc automation only updates think-cell charts "
            "that already exist in the template.",
            "details": details,
            "recommendation": "Open the template in PowerPoint, insert the "
            "think-cell charts you need, and name each one so .ppttc data can "
            "target it by name.",
        }
    return {
        "status": "ok",
        "summary": f"Valid .pptx with {len(slide_parts)} slide(s); think-cell "
        "content was detected.",
        "details": details,
        "recommendation": "Ensure the chart_name values in your .ppttc match "
        "the think-cell element names configured in this template.",
    }


# --------------------------------------------------------------------------
# orchestration
# --------------------------------------------------------------------------
def _safe_check(check: Callable[[], dict[str, Any]]) -> dict[str, Any]:
    """Run a check, converting any unexpected error into an error result."""
    try:
        return check()
    except Exception as exc:  # defensive: registry access can be unpredictable
        return {
            "status": "error",
            "summary": f"This diagnostic check failed unexpectedly: {exc}",
            "details": {},
            "recommendation": None,
        }


def _summarize(checks: dict[str, dict[str, Any]]) -> dict[str, Any]:
    """Roll the individual checks up into an overall status and advice."""
    worst = max(
        (_SEVERITY.get(check["status"], 0) for check in checks.values()),
        default=0,
    )
    status = {0: "ok", 1: "warning", 2: "error"}[worst]
    recommendations = [
        check["recommendation"]
        for check in checks.values()
        if check.get("recommendation")
    ]
    if status == "ok" and recommendations:
        headline = (
            "think-cell is installed and its PowerPoint add-in is active. "
            "The .ppttc pipeline should work -- see the note(s) below."
        )
    elif status == "ok":
        headline = (
            "think-cell is installed and its PowerPoint add-in is active. "
            "The .ppttc automation pipeline should work."
        )
    elif status == "warning":
        headline = (
            "think-cell is installed, but something needs attention before "
            "automation will work reliably."
        )
    else:
        headline = (
            "think-cell automation cannot work until the errors below are "
            "resolved."
        )
    return {
        "status": status,
        "headline": headline,
        "recommendations": recommendations,
    }


def run_diagnostics(template_path: str | None = None) -> dict[str, Any]:
    """Run all think-cell environment checks and return a structured report.

    Args:
        template_path: Optional path to a ``.pptx`` template to inspect as
            part of the report.

    Returns:
        ``{"platform_supported": bool, "summary": {...}, "checks": {...}}``.
    """
    if winreg is None or not sys.platform.startswith("win"):
        return {
            "platform_supported": False,
            "summary": {
                "status": "error",
                "headline": "think-cell diagnostics require Windows.",
                "recommendations": [
                    "Run this tool on the Windows machine where think-cell "
                    "and PowerPoint are installed."
                ],
            },
            "checks": {},
        }

    checks: dict[str, Any] = {
        "thinkcell_installation": _safe_check(_check_installation),
        "powerpoint": _safe_check(_check_powerpoint),
        "com_addin": _safe_check(_check_com_addin),
        "office_disabled_items": _safe_check(_check_disabled_items),
        "license": _safe_check(_check_license),
    }
    if template_path:
        checks["template"] = _safe_check(lambda: _check_template(template_path))

    return {
        "platform_supported": True,
        "summary": _summarize(checks),
        "checks": checks,
    }
