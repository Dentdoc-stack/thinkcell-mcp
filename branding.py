"""Re-brand the bundled think-cell automation template.

think-cell's official automation template ships with think-cell's own sample
styling: a logo picture in the top-right corner of the slide master and a
green/blue accent palette. Neither has anything to do with how ``.ppttc``
automation locates *named* elements, so both can be changed freely.

This module rewrites two **standard PowerPoint parts** of the template:

* ``ppt/theme/theme1.xml`` -- the six theme accent colours, which drive the
  default colours think-cell charts pick for their series.
* ``ppt/slideMasters/slideMaster1.xml`` -- the visible logo picture is
  removed. think-cell's own *hidden*, zero-size "do not delete" data shapes
  are kept untouched.

The named think-cell elements live in other parts (tags / embedded data) and
are never touched here, so the automation keeps working.
"""
from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path
from typing import Any

# A professional, blue-anchored accent palette built from the Big Four
# (Deloitte, PwC, EY, KPMG) brand colours -- a sober default for advisory and
# finance decks. Order maps to theme accent1..accent6.
DEFAULT_ACCENTS: list[str] = [
    "#00338D",  # KPMG primary blue
    "#0091DA",  # KPMG light blue
    "#86BC25",  # Deloitte green
    "#E88D14",  # PwC orange
    "#797878",  # EY gray
    "#FFDB00",  # EY yellow
]

_HEX_RE = re.compile(r"^#?[0-9A-Fa-f]{6}$")
_THEME_PART = "ppt/theme/theme1.xml"
_MASTER_PART = "ppt/slideMasters/slideMaster1.xml"


def _norm_hex(value: Any) -> str | None:
    """Normalise a hex colour to 6 uppercase digits without '#', or None."""
    if not isinstance(value, str) or not _HEX_RE.match(value.strip()):
        return None
    return value.strip().lstrip("#").upper()


def recolor_theme_xml(xml: str, accents: list[str]) -> str:
    """Return ``theme1.xml`` text with accent1..accentN colours replaced.

    ``accents`` holds 6-digit hex strings (no '#'). Only ``<a:accentN>`` slots
    backed by an ``<a:srgbClr>`` are rewritten; any slot beyond ``accents`` or
    not backed by ``srgbClr`` keeps its original colour.
    """
    for index, hex6 in enumerate(accents[:6], start=1):
        xml = re.sub(
            rf'(<a:accent{index}><a:srgbClr val=")'
            rf'[0-9A-Fa-f]{{6}}'
            rf'("/></a:accent{index}>)',
            rf"\g<1>{hex6}\g<2>",
            xml,
        )
    return xml


def strip_logo_from_master(xml: str) -> str:
    """Return slide-master XML with visible ``<p:pic>`` shapes removed.

    think-cell stores its data in *hidden* (``hidden="1"``) zero-size
    pictures; those are kept. Any *visible* picture on the master is template
    branding (the think-cell logo) and is dropped.
    """
    def _keep(match: re.Match[str]) -> str:
        block = match.group(0)
        return block if 'hidden="1"' in block else ""

    return re.sub(r"<p:pic>.*?</p:pic>", _keep, xml, flags=re.DOTALL)


def load_style(template_path: str | Path, style_file: str | Path) -> dict[str, Any]:
    """Load a think-cell style file (.xml) into a template's slide master with think-cell's own ``LoadStyle``.

    A style file sets what the theme colours cannot: think-cell's fill schemes, line styles and label fonts
    (manual: Style files). PowerPoint opens the template without a window, loads the style and saves it.

    Returns ``{"success": bool, "style": str, "errors": list[str]}``; ``style`` is the name think-cell reports.
    """
    template, style = Path(template_path).resolve(), Path(style_file).expanduser().resolve()
    if not style.is_file() or style.suffix.lower() != ".xml":
        return {"success": False, "style": "", "errors": [f"Style file not found (a think-cell .xml): {style_file}"]}
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return {"success": False, "style": "", "errors": ["pywin32 is not installed: python -m pip install pywin32"]}
    pythoncom.CoInitialize()
    pres = None
    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        addin = app.COMAddIns("thinkcell.addin").Object
        pres = app.Presentations.Open(str(template), False, False, False)  # read-write, no window
        addin.LoadStyle(pres.SlideMaster, str(style))
        name = addin.GetStyleName(pres.SlideMaster)
        pres.Save()
        return {"success": True, "style": name, "errors": []}
    except Exception as exc:  # pywintypes.com_error: PowerPoint or think-cell not available, bad style file
        return {"success": False, "style": "", "errors": [f"think-cell could not load the style ({exc})."]}
    finally:
        if pres is not None:
            pres.Close()
        pythoncom.CoUninitialize()


def apply_branding(
    template_path: str | Path,
    accents: list[str] | None = None,
    remove_logo: bool = True,
) -> dict[str, Any]:
    """Re-brand a think-cell automation template ``.pptx`` in place.

    Args:
        template_path: Path to the ``.pptx`` automation template to rewrite.
        accents: Up to six ``#RRGGBB`` hex colours for the theme accent
            palette. Defaults to :data:`DEFAULT_ACCENTS`.
        remove_logo: When True (default) the visible logo picture is removed
            from the slide master.

    Returns:
        ``{"success": bool, "template": str, "accents": list[str],
        "logo_removed": bool, "errors": list[str]}``.
    """
    path = Path(template_path)
    raw = accents if accents is not None else DEFAULT_ACCENTS

    errors: list[str] = []
    normalised: list[str] = []
    if not isinstance(raw, list):
        errors.append("`accents` must be a list of #RRGGBB hex colours")
    else:
        if len(raw) > 6:
            errors.append("at most 6 accent colours are supported")
        for index, value in enumerate(raw[:6]):
            hex6 = _norm_hex(value)
            if hex6 is None:
                errors.append(
                    f"accents[{index}] is not a #RRGGBB hex colour: {value!r}"
                )
            else:
                normalised.append(hex6)
    if errors:
        return {
            "success": False, "template": str(path), "accents": [],
            "logo_removed": False, "errors": errors,
        }
    if not path.is_file():
        return {
            "success": False, "template": str(path), "accents": [],
            "logo_removed": False,
            "errors": [f"Template not found: {path}"],
        }

    try:
        with zipfile.ZipFile(path) as zin:
            parts = [(info, zin.read(info.filename)) for info in zin.infolist()]
    except (OSError, zipfile.BadZipFile) as exc:
        return {
            "success": False, "template": str(path), "accents": [],
            "logo_removed": False,
            "errors": [f"Could not read the template: {exc}"],
        }

    logo_removed = False
    rebuilt: list[tuple[zipfile.ZipInfo, bytes]] = []
    for info, data in parts:
        if info.filename == _THEME_PART and normalised:
            data = recolor_theme_xml(
                data.decode("utf-8"), normalised
            ).encode("utf-8")
        elif info.filename == _MASTER_PART and remove_logo:
            text = data.decode("utf-8")
            stripped = strip_logo_from_master(text)
            logo_removed = stripped != text
            data = stripped.encode("utf-8")
        rebuilt.append((info, data))

    tmp = path.with_name(path.name + ".tmp")
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for info, data in rebuilt:
                zout.writestr(info, data)
        shutil.move(str(tmp), str(path))
    except OSError as exc:
        tmp.unlink(missing_ok=True)
        return {
            "success": False, "template": str(path),
            "accents": ["#" + h for h in normalised],
            "logo_removed": False,
            "errors": [f"Could not write the template: {exc}"],
        }

    return {
        "success": True,
        "template": str(path),
        "accents": ["#" + h for h in normalised],
        "logo_removed": logo_removed,
        "errors": [],
    }
