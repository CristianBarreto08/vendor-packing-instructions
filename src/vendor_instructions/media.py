"""Put back the template images that openpyxl drops on save -- and prove it.

The template carries sample images (what the box label and the product label
must look like) that the written instructions refer to by name. openpyxl does
not round-trip drawings reliably: depending on the environment they are
dropped or re-encoded, and the vendor receives blank sample tabs. That is half
of the labeling explanation, not a cosmetic detail.

After saving, the drawing parts are copied back into the .xlsx (a zip) from
the template:

    xl/media/*                     the image bytes
    xl/drawings/drawingN.xml       anchors (where and how big)
    xl/drawings/_rels/...          drawing -> image links
    xl/worksheets/_rels/sheetN...  sheet -> drawing link
    <drawing r:id="..."/>          reference inside the sheet XML
    [Content_Types].xml            content types for the new parts

Sheets are matched by *name*, never by ``sheetN.xml``: the part number of a
sheet is not stable between the template and what openpyxl writes.
"""

from __future__ import annotations

import hashlib
import posixpath
import re
import shutil
import zipfile
from pathlib import Path

_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_DRAWING_TYPE = _NS_R + "/drawing"
_EMPTY_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships '
               'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               "</Relationships>")
# Per the OOXML schema <drawing> must precede these elements, if present.
_AFTER_DRAWING = ("<legacyDrawing", "<legacyDrawingHF", "<picture", "<oleObjects",
                  "<controls", "<webPublishItems", "<tableParts", "<extLst")


def _resolve(target: str, base: str = "xl") -> str:
    """Zip path of a relationship target.

    Both forms occur: relative (``worksheets/sheet1.xml``) and package-absolute
    (``/xl/worksheets/sheet1.xml``, which openpyxl writes). Treating the
    absolute one as relative yields ``xl/xl/...`` and silently restores nothing.
    """
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(base, target))


def _rels_path(part: str) -> str:
    return posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")


def _relationships(xml: str) -> list[dict[str, str]]:
    return [dict(re.findall(r'(\w+)="([^"]*)"', m)) for m in
            re.findall(r"<Relationship\b([^>]*)/>", xml)]


def sheet_parts(z: zipfile.ZipFile) -> dict[str, str]:
    """``{sheet name: 'xl/worksheets/sheetN.xml'}``."""
    workbook = z.read("xl/workbook.xml").decode("utf8", "replace")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf8", "replace")
    targets = {r["Id"]: _resolve(r["Target"]) for r in _relationships(rels)
               if "Id" in r and "Target" in r}
    parts = {}
    for attrs in re.findall(r"<sheet\b([^>]*)/?>", workbook):
        name = re.search(r'name="([^"]*)"', attrs)
        rid = re.search(r'r:id="([^"]+)"', attrs)
        if name and rid and rid.group(1) in targets:
            parts[name.group(1).replace("&amp;", "&")] = targets[rid.group(1)]
    return parts


def _drawings_of(z: zipfile.ZipFile, sheet_part: str) -> list[str]:
    rels = _rels_path(sheet_part)
    if rels not in z.namelist():
        return []
    xml = z.read(rels).decode("utf8", "replace")
    return [_resolve(r["Target"], posixpath.dirname(sheet_part))
            for r in _relationships(xml) if r.get("Type") == _DRAWING_TYPE]


def _collect(z: zipfile.ZipFile, part: str, out: dict[str, bytes]) -> None:
    """Copy a part and everything it references, recursively."""
    if part in out or part not in z.namelist():
        return
    out[part] = z.read(part)
    rels = _rels_path(part)
    if rels in z.namelist():
        out[rels] = z.read(rels)
        for r in _relationships(out[rels].decode("utf8", "replace")):
            _collect(z, _resolve(r.get("Target", ""), posixpath.dirname(part)), out)


def restore(template: Path, output: Path) -> list[str]:
    """Copy drawings from ``template`` into ``output``. Returns sheets touched.

    Never raises on content mismatch: a file without the sample images is
    better than a corrupt file, and the self-check reports the former.
    """
    with zipfile.ZipFile(template) as zt:
        pending = {name: d for name, part in sheet_parts(zt).items()
                   if (d := _drawings_of(zt, part))}
        if not pending:
            return []
        extra: dict[str, bytes] = {}
        for drawings in pending.values():
            for d in drawings:
                _collect(zt, d, extra)
        template_types = zt.read("[Content_Types].xml").decode("utf8", "replace")

    with zipfile.ZipFile(output) as zo:
        out_parts = sheet_parts(zo)
        content = {n: zo.read(n) for n in zo.namelist()}

    touched = []
    for name, drawings in pending.items():
        part = out_parts.get(name)
        if not part or part not in content:
            continue
        sheet_xml = content[part].decode("utf8", "replace")
        if "<drawing " in sheet_xml:
            continue  # already has one; never duplicate
        rels_name = _rels_path(part)
        rels = content.get(rels_name, _EMPTY_RELS.encode()).decode("utf8", "replace")
        new_ids = []
        for d in drawings:
            rid = _free_rid(rels)
            target = posixpath.relpath(d, posixpath.dirname(part))
            rels = rels.replace("</Relationships>", f'<Relationship Id="{rid}" '
                                f'Type="{_DRAWING_TYPE}" Target="{target}"/></Relationships>')
            new_ids.append(rid)
        tags = "".join(f'<drawing r:id="{r}"/>' for r in new_ids)
        content[part] = _insert_drawing(_declare_r(sheet_xml), tags).encode("utf8")
        content[rels_name] = rels.encode("utf8")
        touched.append(name)

    if not touched:
        return []
    for key, data in extra.items():
        content.setdefault(key, data)
    content["[Content_Types].xml"] = _content_types(
        content["[Content_Types].xml"].decode("utf8", "replace"), template_types, extra
    ).encode("utf8")

    tmp = output.with_suffix(output.suffix + ".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in content.items():
            z.writestr(name, data)
    shutil.move(tmp, output)
    return touched


def _declare_r(sheet_xml: str) -> str:
    """openpyxl does not declare ``xmlns:r`` on sheets it writes; Excel needs it."""
    if "xmlns:r=" in sheet_xml[:2000]:
        return sheet_xml
    m = re.search(r"<worksheet\b", sheet_xml)
    if not m:
        return sheet_xml
    return sheet_xml[:m.end()] + f' xmlns:r="{_NS_R}"' + sheet_xml[m.end():]


def _insert_drawing(sheet_xml: str, tags: str) -> str:
    positions = [i for tag in _AFTER_DRAWING if (i := sheet_xml.find(tag)) != -1]
    cut = min(positions) if positions else sheet_xml.rfind("</worksheet>")
    return sheet_xml[:cut] + tags + sheet_xml[cut:]


def _free_rid(rels: str) -> str:
    used = set(re.findall(r'Id="rId(\d+)"', rels))
    n = 1
    while str(n) in used:
        n += 1
    return f"rId{n}"


def _content_types(types: str, template_types: str, parts: dict[str, bytes]) -> str:
    """Declare the copied parts, reusing the content type the template used.

    Without this Excel reports "a problem with some content" and opens a
    repaired copy -- without the images, which defeats the purpose.
    """
    defaults = set(re.findall(r'<Default Extension="([^"]+)"', types))
    overrides = set(re.findall(r'<Override PartName="([^"]+)"', types))
    additions = []
    for part in parts:
        name = "/" + part
        ext = posixpath.splitext(part)[1].lstrip(".").lower()
        override = re.search(rf'<Override PartName="{re.escape(name)}" ContentType="([^"]+)"',
                             template_types)
        if override and name not in overrides:
            additions.append(f'<Override PartName="{name}" ContentType="{override.group(1)}"/>')
            overrides.add(name)
        elif not override and ext and ext not in defaults:
            default = re.search(rf'<Default Extension="{re.escape(ext)}" ContentType="([^"]+)"',
                                template_types)
            if default:
                additions.append(f'<Default Extension="{ext}" ContentType="{default.group(1)}"/>')
                defaults.add(ext)
    m = re.search(r"<Types\b[^>]*>", types)
    if not additions or not m:
        return types
    return types[:m.end()] + "".join(additions) + types[m.end():]


# ---------------------------------------------------------------------------
# Fingerprints used by the self-check
# ---------------------------------------------------------------------------

def image_hashes(path: Path) -> dict[str, str]:
    """``{image file name: md5}`` -- bytes, not names, so truncation is caught."""
    with zipfile.ZipFile(path) as z:
        return {n.rsplit("/", 1)[-1]: hashlib.md5(z.read(n)).hexdigest()
                for n in z.namelist() if n.startswith("xl/media/")}


def anchor_count(path: Path) -> int:
    """Anchored objects across all drawings: an unreferenced image is invisible."""
    total = 0
    with zipfile.ZipFile(path) as z:
        referenced = set()
        for part in sheet_parts(z).values():
            referenced.update(_drawings_of(z, part))
        for name in referenced:
            if name in z.namelist():
                xml = z.read(name).decode("utf8", "replace")
                total += len(re.findall(r"<xdr:(?:twoCell|oneCell|absolute)Anchor", xml))
    return total
