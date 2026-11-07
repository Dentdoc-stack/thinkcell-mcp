"""Post-conversion single-chart widening for auto-deck output.

``templates/thinkcell_auto.pptx`` places two think-cell chart objects side by
side (``LeftChart``, ``RightChart``). When a ``create_auto_deck`` slide only
supplies one of them, the frames stay where they are and the slide shows a
half-width chart next to an empty placeholder box.

**This must only be applied to an already-converted ``.pptx``, never to the
input template.** Editing a chart graphicFrame's ``<p:xfrm>`` (even leaving
everything else byte-identical) in the *template* that ``ppttc.exe`` reads
breaks think-cell's name-to-shape automation link: ``ppttc.exe`` still exits
0 and produces a file, but the chart silently keeps its old placeholder data
instead of receiving the pushed values -- confirmed empirically by comparing
against an unmodified-template baseline. Applying the same edit to the
*output* file, after ``ppttc.exe`` has already pushed the real data, is safe:
no further automation runs against that file, so there is no link left to
break.

think-cell's element *name* (``LeftChart``/``RightChart``) lives in an
embedded LiteDB blob, not a plain XML attribute, so the two chart frames on a
given slide can't be told apart by name in the OOXML alone. Which one to
widen is instead read from the source ``.ppttc`` document (which names are
present in that slide's ``data`` entries); which *frame* is which is told
apart by geometry -- of the two real chart frames, the one closer to the
left edge is ``LeftChart``. think-cell's own tiny 1588x1588 EMU "do not
delete" automation-link placeholder is excluded by its fixed small size.
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Any

# Matches the bundled template's 16:9 canvas (12192000 x 6858000 EMU).
_SLIDE_WIDTH = 12192000

_XFRM_RE = re.compile(
    r'<p:xfrm><a:off x="(\d+)" y="(\d+)"/><a:ext cx="(\d+)" cy="(\d+)"/></p:xfrm>'
)
_FRAME_RE = re.compile(r"<p:graphicFrame>.*?</p:graphicFrame>", re.DOTALL)
_SLIDE_PART_RE = re.compile(r"^ppt/slides/slide(\d+)\.xml$")

# think-cell's hidden automation-link frame is a fixed tiny square; anything
# that small is that placeholder, not a real chart.
_AUTOMATION_LINK_MAX_EXTENT = 2000

LEFT_CHART_NAME = "LeftChart"
RIGHT_CHART_NAME = "RightChart"


def _find_chart_frames(xml: str) -> list[dict[str, Any]]:
    """Locate the two real chart graphicFrames, sorted left-to-right by x."""
    frames: list[dict[str, Any]] = []
    for match in _FRAME_RE.finditer(xml):
        block = match.group(0)
        xfrm = _XFRM_RE.search(block)
        if not xfrm:
            continue
        x, y, cx, cy = (int(value) for value in xfrm.groups())
        if cx <= _AUTOMATION_LINK_MAX_EXTENT and cy <= _AUTOMATION_LINK_MAX_EXTENT:
            continue
        frames.append({"span": match.span(0), "x": x, "y": y, "cx": cx, "cy": cy})
    frames.sort(key=lambda frame: frame["x"])
    return frames


def _apply_edits(xml: str, edits: list[tuple[tuple[int, int], str]]) -> str:
    """Apply (span, replacement) edits from the tail so earlier offsets stay valid."""
    for (start, end), replacement in sorted(edits, key=lambda e: e[0][0], reverse=True):
        xml = xml[:start] + replacement + xml[end:]
    return xml


def widen_single_chart(xml: str, keep: str) -> str:
    """Widen ``keep`` ('left' or 'right') to the full content width.

    The other chart is moved off-slide rather than deleted, so nothing about
    its shape identity changes -- it simply has no data pushed to it (per
    the caller's own selection of which slides qualify) and stays invisible.

    Returns ``xml`` unchanged (a defensive no-op) if the slide doesn't have
    exactly the two expected chart frames, rather than guessing.

    Only call this on an already-converted output ``.pptx`` -- see the
    module docstring for why editing the pre-conversion template breaks
    automation.
    """
    frames = _find_chart_frames(xml)
    if len(frames) != 2:
        return xml
    left, right = frames
    union_x = left["x"]
    union_width = (right["x"] + right["cx"]) - union_x
    target, drop = (left, right) if keep == "left" else (right, left)
    off_stage_x = _SLIDE_WIDTH + 1_000_000

    def rewrite(frame: dict[str, Any], x: int, cx: int) -> tuple[tuple[int, int], str]:
        block = xml[frame["span"][0]:frame["span"][1]]
        xfrm = _XFRM_RE.search(block)
        assert xfrm is not None  # guaranteed by _find_chart_frames
        new_xfrm = (
            f'<p:xfrm><a:off x="{x}" y="{frame["y"]}"/>'
            f'<a:ext cx="{cx}" cy="{frame["cy"]}"/></p:xfrm>'
        )
        return frame["span"], block[:xfrm.start()] + new_xfrm + block[xfrm.end():]

    edits = [
        rewrite(target, union_x, union_width),
        rewrite(drop, off_stage_x, drop["cx"]),
    ]
    return _apply_edits(xml, edits)


def _slide_keep(entry: dict[str, Any]) -> str | None:
    """Which chart to widen for one ``.ppttc`` document entry, or None."""
    names = {
        element.get("name")
        for element in entry.get("data", [])
        if isinstance(element, dict)
    }
    has_left = LEFT_CHART_NAME in names
    has_right = RIGHT_CHART_NAME in names
    if has_left and not has_right:
        return "left"
    if has_right and not has_left:
        return "right"
    return None  # both, neither, or not an auto-deck-shaped slide at all


def widen_single_charts_in_pptx(
    pptx_path: str | Path, document: list[dict[str, Any]]
) -> int:
    """Post-process a *converted* auto-deck ``.pptx``: widen any slide that
    only has ``LeftChart`` or only ``RightChart`` populated (per the source
    ``.ppttc`` ``document``) to the full content width.

    Safe to call on any converted ``.pptx``, including ones that aren't
    auto-deck output: a slide only qualifies when its corresponding
    ``.ppttc`` entry names exactly one of ``LeftChart``/``RightChart`` *and*
    the output slide's XML actually has the expected two-chart-frame shape;
    everything else is left untouched. Output slide parts are matched to
    ``document`` entries by position (``slideN.xml`` <-> ``document[N-1]``),
    which is how ``ppttc.exe`` generates them (one slide per top-level
    ``.ppttc`` entry, in order).

    Returns the number of slides actually widened. Never raises: any
    failure (unreadable file, unexpected shape, write error) is treated as
    "nothing to widen" and leaves the ``.pptx`` untouched, since this is a
    cosmetic improvement layered on top of an already-successful conversion.
    """
    path = Path(pptx_path)
    try:
        with zipfile.ZipFile(path) as zin:
            parts = [(info, zin.read(info.filename)) for info in zin.infolist()]
    except (OSError, zipfile.BadZipFile):
        return 0

    changed = 0
    rebuilt: list[tuple[zipfile.ZipInfo, bytes]] = []
    for info, data in parts:
        match = _SLIDE_PART_RE.match(info.filename)
        if match:
            index = int(match.group(1)) - 1
            if 0 <= index < len(document):
                keep = _slide_keep(document[index])
                if keep:
                    text = data.decode("utf-8")
                    new_text = widen_single_chart(text, keep)
                    if new_text != text:
                        changed += 1
                        data = new_text.encode("utf-8")
        rebuilt.append((info, data))

    if not changed:
        return 0

    tmp = path.with_name(path.name + ".tmp")
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for info, data in rebuilt:
                zout.writestr(info, data)
        tmp.replace(path)
    except OSError:
        tmp.unlink(missing_ok=True)
        return 0
    return changed
