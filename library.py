"""Chart templates from any think-cell slide: think-cell's own chart library or the user's decks.

ppttc.exe fills think-cell elements by name. Naming used to be a manual step (type the name into think-cell's
toolbar). think-cell keeps each slide's chart data in an OLE object (``ppt/embeddings/oleObject*.bin``, stream
``think-cellXML``) where every chart element carries ``<m_strName>``, so this module writes the names in code
instead. That turns any think-cell slide into a template:

* ``ensure_library()`` cuts the slides of think-cell's chart library (installed with think-cell, under
  ``templates/think-cell Charts``) into one-slide templates in ``templates/library/``. They are copied from the
  local install, never shipped with this project.
* ``template_from_slide()`` does the same for a slide of the user's own deck (``templates/user/``), keeping its
  style.
* ``fill_template()`` writes the ``.ppttc`` for a template, runs ppttc.exe and checks which charts took the data.

Charts are named ``c1``, ``c2``, ... in reading order (top to bottom, then left to right).
"""
from __future__ import annotations

import io
import json
import os
import re
import tempfile
import zipfile
from html import escape, unescape
from pathlib import Path
from typing import Any

from autodeck import OUTPUT_DIR, THINKCELL_DIR

PROJECT_ROOT = Path(__file__).resolve().parent
LIBRARY_DIR = PROJECT_ROOT / "templates" / "library"
USER_DIR = PROJECT_ROOT / "templates" / "user"
INDEX_FILE = LIBRARY_DIR / "index.json"
INSTALL_LIBRARY = THINKCELL_DIR / "templates" / "think-cell Charts"

OLE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
EMU_PER_UNIT = 12700 / 8  # think-cell frame rectangles are in 1/8 pt
# The think-cell elements ppttc.exe fills by name. Gantt charts (CGanttSE) are left out: ppttc.exe garbles
# their date axis.
CHART_CLASSES = {"CSequenceChartSE": "chart", "CPieChartSE": "pie", "CScatterChartSE": "scatter"}
_ELEMENT = re.compile(rb'<(%s) id="(\d+)"[^>]*>' % b"|".join(c.encode() for c in CHART_CLASSES))


# --------------------------------------------------------------------------
# think-cell's data object
# --------------------------------------------------------------------------
def _ole_parts(slide):
    for rel in slide.part.rels.values():
        if rel.reltype == OLE_REL and not rel.is_external:
            yield rel.target_part


def _read_xml(blob: bytes) -> bytes | None:
    import olefile

    if not olefile.isOleFile(io.BytesIO(blob)):
        return None
    ole = olefile.OleFileIO(io.BytesIO(blob))
    return ole.openstream("think-cellXML").read() if ole.exists("think-cellXML") else None


def _write_xml(blob: bytes, xml: bytes) -> bytes:
    """The OLE file with its think-cellXML stream replaced (any length), via Windows structured storage."""
    import pythoncom
    from win32com import storagecon

    fd, path = tempfile.mkstemp(suffix=".bin")
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(blob)
        mode = storagecon.STGM_READWRITE | storagecon.STGM_SHARE_EXCLUSIVE
        stg = pythoncom.StgOpenStorage(path, None, mode, None, 0)
        stm = stg.OpenStream("think-cellXML", None, mode, 0)
        stm.SetSize(len(xml))
        stm.Seek(0, 0)
        stm.Write(xml)
        stm.Commit(0)
        stg.Commit(0)
        del stm, stg
        return Path(path).read_bytes()
    finally:
        os.remove(path)


def _elements(xml: bytes):
    """(class, id, start, end) per chart element; it ends where the next one starts."""
    found = [(m.group(1).decode(), m.group(2).decode(), m.start()) for m in _ELEMENT.finditer(xml)]
    return [(c, i, s, found[k + 1][2] if k + 1 < len(found) else len(xml)) for k, (c, i, s) in enumerate(found)]


def _frame(xml: bytes, body: bytes):
    """The element's frame on the slide (EMU) from its first PowerPoint-object reference, or None."""
    for ref in re.findall(rb'<m_ppt\w+ idref="(\d+)"/>', body)[:3]:
        rect = re.search(rb'<C\w+ id="%s">\s*<m_rectPPTShape left="(-?\d+)" top="(-?\d+)" right="(-?\d+)" '
                         rb'bottom="(-?\d+)"/>' % ref, xml)
        if rect:
            left, top, right, bottom = (int(v) * EMU_PER_UNIT for v in rect.groups())
            return left, top, right - left, bottom - top
    return None


def _value(source: bytes) -> str:
    """The text or number a variable source holds ('' when it holds rich text or nothing)."""
    val = re.search(rb'<m_varval type="(\d+)"(?: val="([^"]*)")?>?([^<]*)', source)
    if not val:
        return ""
    if val.group(1) == b"5":
        return unescape(val.group(3).decode("utf-8")).replace("\r", "\n")
    try:
        number = float(val.group(2))
    except (TypeError, ValueError):
        return ""
    return str(int(number)) if number.is_integer() else f"{number:g}"


def _source_xml(xml: bytes, element: bytes, tag: bytes) -> bytes:
    """The variable source in ``<tag>`` of ``element``: a reference to a CVariableSource in think-cell's own
    files, inline after ppttc.exe has run."""
    m = re.search(rb'<%s(?: idref="(\d+)"/>|>(.*?)</%s>)' % (tag, tag), element, re.S)
    if not m:
        return b""
    if m.group(1):
        src = re.search(rb'<CVariableSource id="%s">(.*?)</CVariableSource>' % m.group(1), xml, re.S)
        return src.group(1) if src else b""
    return m.group(2)


def _source(xml: bytes, element: bytes, tag: bytes) -> str:
    return _value(_source_xml(xml, element, tag))


def _element(xml: bytes, cls: bytes, eid: bytes) -> bytes:
    m = re.search(rb'<%s id="%s">(.*?)</%s>' % (cls, eid, cls), xml, re.S)
    return m.group(1) if m else b""


def _number(text: str) -> float | None:
    try:
        return float(text)
    except ValueError:
        return None


def _datasheet(xml: bytes, body: bytes) -> tuple[list[str], int, float | None]:
    """A sequence chart's category labels, series count and, when its datasheet layout has the optional 100%=
    row (then ppttc.exe reads the second JSON row as that row), the template's first 100%= value."""
    table = re.search(rb'<m_dtable idref="(\d+)"/>', body)
    if not table:
        return [], 0, None
    sheet = re.search(rb'<CSequenceChartDataTable id="%s"[^>]*>(.*?)</CSequenceChartDataTable>' % table.group(1),
                      xml, re.S)
    vectors = sheet and re.search(rb'<ocol length="\d+">((?:<elem idref="\d+"/>)*)', sheet.group(1))
    if not vectors:
        return [], 0, None
    cats, series, total = [], 0, None
    for k, vid in enumerate(re.findall(rb'idref="(\d+)"', vectors.group(1))):
        vec = _element(xml, b"CSequenceChartDataVector", vid)
        scalars = re.match(rb'<ocol length="(\d+)">((?:<elem idref="\d+"/>)*)', vec)
        if not scalars:
            continue
        series = max(series, int(scalars.group(1)))
        cats.append(_source(xml, vec, b"m_varsrcCategory"))
        # With the 100%= row on, the column's total is a datasheet cell and carries number formatting
        # (<m_precInitial>); without it think-cell only caches the sum. Found by comparing templates on
        # think-cell 14; the manual's "Optional rows and columns" explains what the row does.
        extent = _source_xml(xml, vec, b"m_varsrcAbsoluteExtent") if k == 0 else b""
        if b"<m_precInitial" in extent:
            total = _number(_value(extent))
            if total is None:
                total = 100.0
    return cats, series, total


def charts(slide) -> list[dict[str, Any]]:
    """The think-cell charts on a slide, in reading order."""
    out = []
    for part in _ole_parts(slide):
        xml = _read_xml(part.blob)
        if not xml:
            continue
        for cls, cid, s, e in _elements(xml):
            body = xml[s:e]
            name = re.search(rb"<m_strName>(.*?)</m_strName>", body)
            if not name:
                continue
            cats, series, total = _datasheet(xml, body)
            out.append({"id": cid, "kind": CHART_CLASSES[cls], "name": unescape(name.group(1).decode("utf-8")),
                        "frame": _frame(xml, body), "categories": cats, "series": series, "total": total,
                        "part": part})
    # reading order: half-inch bands from the top, then left to right; frameless ones last
    out.sort(key=lambda c: (c["frame"] is None, round(c["frame"][1] / 457200) if c["frame"] else 0,
                            c["frame"][0] if c["frame"] else 0))
    return out


def _own_copy(slide, part):
    """The slide's own copy of an OLE part (PowerPoint shares one embedding between slides with equal objects,
    so a name written there would land on all of them)."""
    from pptx.opc.package import Part

    pkg = slide.part.package
    new = Part(pkg.next_partname("/ppt/embeddings/oleObject%d.bin"), part.content_type, pkg, part.blob)
    for rel in slide.part.rels.values():
        if rel.reltype == OLE_REL and not rel.is_external and rel.target_part is part:
            rel._target = new
            for cached in ("target_part", "target_partname"):  # python-pptx keeps them from the first look
                rel.__dict__.pop(cached, None)
    return new


def set_names(slide, names: dict[str, str]) -> list[str]:
    """Name charts by element id ({id: name}); every name slot of the element gets it. Returns the ids named."""
    done = []
    for part in list(_ole_parts(slide)):
        xml = _read_xml(part.blob)
        if not xml:
            continue
        new, last = [], 0
        for _, cid, s, e in _elements(xml):
            name = names.get(cid)
            if name is None:
                continue
            body = re.sub(rb"<m_strName>.*?</m_strName>",
                          b"<m_strName>" + escape(name).encode("utf-8") + b"</m_strName>", xml[s:e])
            new += [xml[last:s], body]
            last = e
            done.append(cid)
        if new:
            _own_copy(slide, part)._blob = _write_xml(part.blob, b"".join(new) + xml[last:])
    return done


# --------------------------------------------------------------------------
# templates
# --------------------------------------------------------------------------
def _open(path: Path):
    """A python-pptx Presentation of a .pptx, or of a .potx (its main part is relabelled as a presentation)."""
    from pptx import Presentation

    if path.suffix.lower() != ".potx":
        return Presentation(str(path))
    buf = io.BytesIO()
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(b"presentationml.template.main+xml", b"presentationml.presentation.main+xml")
            dst.writestr(item, data)
    buf.seek(0)
    return Presentation(buf)


def _walk(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == 6:  # group
            yield from _walk(shape.shapes)


def _is_tc(shape, part) -> bool:
    """A think-cell shape (tagged THINKCELL...): ppttc.exe redraws these, so they are never edited here."""
    rids = re.findall(r'<p:tags r:id="(rId\d+)"', shape._element.xml)
    return any(b"THINKCELL" in part.related_part(r).blob for r in rids)


def text_shapes(slide) -> list:
    """The slide's own text shapes (not think-cell's labels) that hold text."""
    return [sh for sh in _walk(slide.shapes) if sh.has_text_frame and sh.text_frame.text.strip()
            and not _is_tc(sh, slide.part)]


def _drop_library_notes(slide) -> None:
    """think-cell's library slides carry how-to notes beside the slide and a keywords footer: not slide content."""
    for shape in text_shapes(slide):
        if shape.top is not None and (shape.top < 0 or shape.text_frame.text.strip().startswith("Keywords:")):
            shape._element.getparent().remove(shape._element)


def _set_text(shape, text: str) -> None:
    """Replace a shape's text line by line; line k keeps paragraph k's formatting (the last one's after that)."""
    import copy

    body = shape.text_frame._txBody
    paragraphs = body.findall("{http://schemas.openxmlformats.org/drawingml/2006/main}p")
    for p in paragraphs:
        body.remove(p)
    for k, line in enumerate(text.split("\n")):
        p = copy.deepcopy(paragraphs[min(k, len(paragraphs) - 1)])
        runs = p.findall("{http://schemas.openxmlformats.org/drawingml/2006/main}r")
        for extra in runs[1:] + p.findall("{http://schemas.openxmlformats.org/drawingml/2006/main}br"):
            p.remove(extra)
        if runs:
            runs[0].find("{http://schemas.openxmlformats.org/drawingml/2006/main}t").text = line
        else:
            from pptx.text.text import _Paragraph
            _Paragraph(p, shape.text_frame).text = line
        body.append(p)


def replace_texts(pptx: Path, texts: dict[str, str]) -> list[str]:
    """In slide 1 of ``pptx``, give each text shape whose whole text is a key the new text ('' deletes the
    shape). Returns the keys that matched no shape."""
    from pptx import Presentation

    prs = Presentation(str(pptx))
    slide = prs.slides[0]
    wanted = {old.strip(): new for old, new in texts.items()}
    hit = set()
    for shape in text_shapes(slide):
        old = shape.text_frame.text.strip()
        if old in wanted:
            hit.add(old)
            if wanted[old] == "":
                shape._element.getparent().remove(shape._element)
            else:
                _set_text(shape, wanted[old])
    prs.save(str(pptx))
    return sorted(set(wanted) - hit)


def _title(slide) -> str:
    try:
        if slide.shapes.title is not None and slide.shapes.title.text_frame.text.strip():
            return slide.shapes.title.text_frame.text.strip()
    except (AttributeError, KeyError):
        pass
    texts = [sh.text_frame.text.strip() for sh in slide.shapes if sh.has_text_frame and sh.text_frame.text.strip()]
    return max(texts, key=len)[:120] if texts else ""


def cut_template(source: Path, slide_number: int, dest: Path, library_slide: bool = False) -> dict[str, Any]:
    """Save slide ``slide_number`` (1-based) of ``source`` alone as ``dest``, its charts named c1, c2, ...

    Returns ``{"title", "charts": [{"name", "kind", "categories", "series"}], "texts": [...]}``.
    """
    prs = _open(source)
    ids = prs.slides._sldIdLst
    if not 1 <= slide_number <= len(ids):
        raise ValueError(f"slide {slide_number} does not exist (the deck has {len(ids)} slides)")
    for i, sld in reversed(list(enumerate(list(ids), 1))):
        if i != slide_number:
            prs.part.drop_rel(sld.rId)
            ids.remove(sld)
    slide = prs.slides[0]
    if library_slide:
        _drop_library_notes(slide)
    found = charts(slide)
    set_names(slide, {c["id"]: f"c{k}" for k, c in enumerate(found, 1)})
    dest.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(dest))  # python-pptx leaves out the parts only the dropped slides used
    texts = list(dict.fromkeys(sh.text_frame.text.strip() for sh in text_shapes(slide)))
    return {"title": _title(slide), "charts": [
        {"name": f"c{k}", "kind": c["kind"], "categories": c["categories"], "series": c["series"],
         **({"total": c["total"]} if c["total"] is not None else {}),
         **({"was_named": c["name"]} if c["name"] else {})} for k, c in enumerate(found, 1)], "texts": texts}


def _slug(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-") or "template"


def ensure_library(rebuild: bool = False) -> list[dict[str, Any]]:
    """Cut think-cell's installed chart library into named templates once; returns the index."""
    if INDEX_FILE.is_file() and not rebuild:
        return json.loads(INDEX_FILE.read_text(encoding="utf-8"))
    if not INSTALL_LIBRARY.is_dir():
        raise FileNotFoundError(f"think-cell's chart library was not found at {INSTALL_LIBRARY}")
    index = []
    for potx in sorted(INSTALL_LIBRARY.rglob("*.potx")):
        group = _slug(potx.stem)
        for number in range(1, len(_open(potx).slides) + 1):
            template_id = f"{group}-{number}"
            info = cut_template(potx, number, LIBRARY_DIR / f"{template_id}.pptx", library_slide=True)
            if info["charts"]:
                index.append({"id": template_id, "group": potx.stem, **info})
            else:
                (LIBRARY_DIR / f"{template_id}.pptx").unlink()
    INDEX_FILE.write_text(json.dumps(index, indent=1, ensure_ascii=False), encoding="utf-8")
    return index


def user_templates() -> list[dict[str, Any]]:
    out = []
    for meta in sorted(USER_DIR.glob("*.json")):
        out.append(json.loads(meta.read_text(encoding="utf-8")))
    return out


def template_path(template_id: str) -> Path | None:
    """``user/<name>`` or a library id -> the template file, if it exists."""
    if template_id.startswith("user/"):
        path = USER_DIR / f"{_slug(template_id[5:])}.pptx"
    else:
        path = LIBRARY_DIR / f"{_slug(template_id)}.pptx"
    return path if path.is_file() else None


def save_user_template(source: Path, slide_number: int, name: str) -> dict[str, Any]:
    slug = _slug(name)
    info = cut_template(source, slide_number, USER_DIR / f"{slug}.pptx")
    entry = {"id": f"user/{slug}", "group": "My slides", "source": source.name, "slide": slide_number, **info}
    (USER_DIR / f"{slug}.json").write_text(json.dumps(entry, indent=1, ensure_ascii=False), encoding="utf-8")
    return entry


# --------------------------------------------------------------------------
# filling
# --------------------------------------------------------------------------
_ISO_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def _cell(value: Any, number_key: str) -> Any:
    if value is None:
        return None
    if isinstance(value, bool):
        raise ValueError(f"{value!r} is not a number or text")
    if isinstance(value, (int, float)):
        return {number_key: value}
    return {"string": str(value)}


def chart_table(spec: dict[str, Any]) -> list[list[Any]]:
    """A ppttc table from ``{"categories": [...], "series": [{"name", "values", "fill"?}], "percent"?}``,
    or ``{"table": [...]}`` passed through as-is."""
    if "table" in spec:
        if not isinstance(spec["table"], list):
            raise ValueError("'table' must be a list of rows")
        return spec["table"]
    cats, series = spec.get("categories"), spec.get("series")
    if not isinstance(cats, list) or not cats:
        raise ValueError("'categories' must be a non-empty list")
    if not isinstance(series, list) or not series:
        raise ValueError("'series' must be a non-empty list of {name, values}")
    key = "percentage" if spec.get("percent") else "number"
    rows = [[None] + [{"date": c} if isinstance(c, str) and _ISO_DATE.match(c) else {"string": str(c)}
                      for c in cats]]
    total = spec.get("total")
    if total is not None:  # the 100%= row: one value for every category, or one per category
        totals = total if isinstance(total, list) else [total] * len(cats)
        if len(totals) != len(cats):
            raise ValueError(f"'total' must be one number or one per category ({len(cats)})")
        rows.append([None] + [_cell(t, "number") for t in totals])
    for k, s in enumerate(series):
        values = s.get("values") if isinstance(s, dict) else None
        if not isinstance(values, list) or len(values) != len(cats):
            raise ValueError(f"series[{k}].values must list one value per category ({len(cats)})")
        row = [_cell(s.get("name"), key) if s.get("name") else None] + [_cell(v, key) for v in values]
        if s.get("fill"):
            row = [row[0]] + [dict(c, fill=s["fill"]) if c else c for c in row[1:]]
        rows.append(row)
    return rows


def _set_title(pptx: Path, title: str) -> bool:
    """Put ``title`` in the slide's title placeholder, keeping its formatting."""
    from pptx import Presentation

    prs = Presentation(str(pptx))
    shape = prs.slides[0].shapes.title
    if shape is None:
        return False
    _set_text(shape, title)
    prs.save(str(pptx))
    return True


def fill_template(template: Path, charts_data: dict[str, dict[str, Any]], title: str = "",
                  output_name: str = "", texts: dict[str, str] | None = None) -> dict[str, Any]:
    """Fill a named template through ppttc.exe; reports which charts show the new categories."""
    from pptx import Presentation

    from converter import convert_ppttc

    if not isinstance(charts_data, dict) or not charts_data:
        return {"success": False, "pptx_path": None, "errors": ["Give data for at least one chart"]}
    known = {c["name"]: c for c in charts(Presentation(str(template)).slides[0])}
    unknown = sorted(set(charts_data) - set(known))
    if unknown:
        return {"success": False, "pptx_path": None,
                "errors": [f"No chart named {', '.join(unknown)} in this template (it has {', '.join(sorted(known))})"]}
    errors, data = [], []
    for name, spec in charts_data.items():
        spec = dict(spec) if isinstance(spec, dict) else {}
        if "table" not in spec and "total" not in spec and known[name]["total"] is not None:
            spec["total"] = known[name]["total"]  # keep the template's typed-in 100%= value
        try:
            data.append({"name": name, "table": chart_table(spec)})
        except ValueError as exc:
            errors.append(f"{name}: {exc}")
    if errors:
        return {"success": False, "pptx_path": None, "errors": errors}
    stem = _slug(output_name or template.stem)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    job = OUTPUT_DIR / f"{stem}.ppttc"
    job.write_text(json.dumps([{"template": str(template), "data": data}], indent=1), encoding="utf-8")
    result = convert_ppttc(str(job))
    if not result.get("success"):
        return {"success": False, "pptx_path": None, "ppttc_path": str(job),
                "errors": [result.get("error") or "ppttc.exe failed"]}
    pptx = Path(result["pptx_path"])
    titled = _set_title(pptx, title) if title else False
    unmatched = replace_texts(pptx, texts) if texts else []
    filled, unchecked, missed = [], [], []
    after = {c["name"]: c for c in charts(Presentation(str(pptx)).slides[0])}
    for name, spec in charts_data.items():
        want = [str(c) for c in spec.get("categories", [])] if isinstance(spec, dict) else []
        got = after.get(name, {}).get("categories")
        if not want or not got:
            unchecked.append(name)
        elif [g.strip() for g in got[:len(want)]] == [w.strip() for w in want]:
            filled.append(name)
        else:
            missed.append(name)
    warnings = [f"{', '.join(missed)} did not take the new data: think-cell skipped it. Open that chart's "
                f"datasheet in the template once in PowerPoint, save, and try again."] if missed else []
    if unmatched:
        warnings.append(f"No text shape reads exactly {unmatched}; see the template's 'texts'.")
    return {"success": True, "pptx_path": str(pptx), "ppttc_path": str(job), "filled": filled,
            "not_checked": unchecked, "not_filled": missed, "title_set": titled, "errors": [], "warnings": warnings}
