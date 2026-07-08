"""
Full-metadata Excel serializer / reconstructor.

Serialize: Excel → dict with per-cell fill, font, merge, alignment, numfmt,
           column widths, row heights, borders.
Reconstruct: dict → Excel (.xlsx) preserving all of the above.
"""

import json
import os
from pathlib import Path
from urllib.parse import urlparse, unquote

# lxml can hang inside the PyInstaller onefile build on this toolchain.
# Force openpyxl to use the stdlib/et_xmlfile XML path before any openpyxl import.
os.environ.setdefault("OPENPYXL_LXML", "False")


# ── URI → Path ────────────────────────────────────────────────────────────────

def uri_to_path(uri: str) -> Path:
    if uri.startswith("file://"):
        path = unquote(urlparse(uri).path)
        # Windows: /D:/foo/bar.xlsx → D:/foo/bar.xlsx
        if path.startswith("/") and len(path) > 2 and path[2] == ":":
            path = path[1:]
        return Path(path)
    return Path(uri)


# ── Color helpers ─────────────────────────────────────────────────────────────

# Default Office theme base colors (indices 0-11: dk1,lt1,dk2,lt2,accent1-6,hlink,folHlink)
_OFFICE_THEME_COLORS: list[str] = [
    "000000", "FFFFFF", "44546A", "E7E6E6",
    "4472C4", "ED7D31", "A5A5A5", "FFC000",
    "5B9BD5", "70AD47", "0563C1", "954F72",
]


def _apply_tint(hex6: str, tint: float) -> str:
    """Apply Excel luminance tint to a 6-char hex RGB and return 8-char ARGB."""
    import colorsys
    r, g, b = int(hex6[0:2], 16) / 255, int(hex6[2:4], 16) / 255, int(hex6[4:6], 16) / 255
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    l = l * (1 + tint) if tint < 0 else l + (1 - l) * tint
    l = max(0.0, min(1.0, l))
    r, g, b = colorsys.hls_to_rgb(h, l, s)
    return f"FF{int(r * 255):02X}{int(g * 255):02X}{int(b * 255):02X}"


def _wb_theme_colors(wb) -> list[str]:
    """
    Extract the 12 base theme RGB values from the workbook's theme XML.
    Falls back to Office 2016 defaults if the theme is absent or unparseable.
    """
    raw = getattr(wb, "loaded_theme", None)
    if not raw:
        return list(_OFFICE_THEME_COLORS)
    try:
        from xml.etree import ElementTree as ET
        ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
        root = ET.fromstring(raw if isinstance(raw, (bytes, str)) else raw)
        scheme = root.find(f".//{{{ns}}}clrScheme")
        if scheme is None:
            return list(_OFFICE_THEME_COLORS)
        result = list(_OFFICE_THEME_COLORS)
        for i, child in enumerate(scheme):
            if i >= 12:
                break
            srgb = child.find(f"{{{ns}}}srgbClr")
            if srgb is not None:
                val = srgb.get("val", "")
                if len(val) == 6:
                    result[i] = val.upper()
                continue
            sys_clr = child.find(f"{{{ns}}}sysClr")
            if sys_clr is not None:
                last = sys_clr.get("lastClr", "")
                if len(last) == 6:
                    result[i] = last.upper()
        return result
    except Exception:
        return list(_OFFICE_THEME_COLORS)


def _resolve_color(fc, theme_colors: list[str]) -> str | None:
    """
    Resolve an openpyxl Color to an 8-char ARGB hex string, or None if transparent/absent.
    Handles rgb, indexed, and theme types.
    Only excludes "00000000" (fully transparent) — all real colors including black and white
    are preserved so fill round-trips correctly.
    """
    if fc is None:
        return None
    try:
        if fc.type == "rgb":
            rgb = fc.rgb
            if isinstance(rgb, str) and rgb != "00000000":
                return rgb
        elif fc.type == "indexed":
            from openpyxl.styles.colors import COLOR_INDEX
            idx = fc.indexed
            if idx is not None and 0 <= idx < len(COLOR_INDEX):
                argb = COLOR_INDEX[idx]
                if argb != "00000000":
                    return argb
        elif fc.type == "theme":
            idx = fc.theme
            if idx is not None and 0 <= idx < len(theme_colors):
                base = theme_colors[idx]
                tint = fc.tint or 0.0
                return _apply_tint(base, tint) if tint else f"FF{base}"
    except Exception:
        pass
    return None


def _parse_xml_attrs(attr_text: str) -> dict[str, str]:
    """Parse a simple XML attribute string into {name: value}."""
    import re
    return {m.group(1): m.group(2) for m in re.finditer(r'([\w:.-]+)="([^"]*)"', attr_text)}


def _extract_xml_children(parent_xml: str, child_tag: str) -> list[str]:
    """Return raw child XML elements for a simple OOXML container."""
    import re
    return [m.group(0) for m in re.finditer(
        rf"<{child_tag}\b[^>]*/>|<{child_tag}\b[^>]*>.*?</{child_tag}>",
        parent_xml,
        re.DOTALL,
    )]


def _color_ref_from_openpyxl(fc) -> dict | None:
    """Serialize an openpyxl color reference for later reconstruction."""
    if fc is None:
        return None
    ctype = getattr(fc, "type", None)
    try:
        if ctype == "rgb" and isinstance(fc.rgb, str) and fc.rgb != "00000000":
            ref = {"type": "rgb", "rgb": fc.rgb}
        elif ctype == "theme" and fc.theme is not None:
            ref = {"type": "theme", "theme": int(fc.theme)}
        elif ctype == "indexed" and fc.indexed is not None:
            ref = {"type": "indexed", "indexed": int(fc.indexed)}
        elif ctype == "auto" and fc.auto is not None:
            ref = {"type": "auto", "auto": bool(fc.auto)}
        else:
            return None
        tint = getattr(fc, "tint", None)
        if tint:
            ref["tint"] = float(tint)
        return ref
    except Exception:
        return None


def _make_color_from_ref(ref: dict):
    from openpyxl.styles.colors import Color

    ctype = ref.get("type")
    kw = {}
    if ctype == "theme":
        kw["theme"] = int(ref["theme"])
    elif ctype == "indexed":
        kw["indexed"] = int(ref["indexed"])
    elif ctype == "auto":
        kw["auto"] = bool(ref.get("auto", True))
    elif ctype == "rgb":
        kw["rgb"] = ref["rgb"]
    else:
        return None
    if ref.get("tint") is not None:
        kw["tint"] = float(ref["tint"])
    return Color(**kw)


def _usable_raw_fill(cd: dict) -> dict | None:
    raw = cd.get("_fill_raw")
    if not raw:
        return None
    # Keep the raw fill only while the public resolved fill has not changed.
    if cd.get("fill") != raw.get("rgb"):
        return None
    return raw


def _usable_raw_font(cd: dict) -> dict | None:
    raw = cd.get("_font_raw")
    if not raw:
        return None
    checks = (
        ("font", "name"),
        ("size", "size"),
        ("bold", "bold"),
        ("italic", "italic"),
        ("uline", "underline"),
        ("strike", "strike"),
        ("vAlign", "vertAlign"),
    )
    for public_key, raw_key in checks:
        if raw.get(raw_key) != cd.get(public_key):
            return None
    return raw


def _usable_raw_color_ref(ref: dict | None, current_rgb: str | None) -> dict | None:
    if not ref:
        return None
    raw_rgb = ref.get("rgb")
    if current_rgb == raw_rgb:
        return ref
    if current_rgb is None and raw_rgb in (None, "FF000000", "00000000"):
        return ref
    return None


def _font_raw_from_openpyxl(font, fcolor: str | None) -> dict | None:
    if font is None:
        return None
    raw = {
        "name":      font.name,
        "size":      font.size,
        "bold":      bool(font.bold),
        "italic":    bool(font.italic),
        "underline": font.underline,
        "strike":    bool(font.strike),
        "vertAlign": font.vertAlign,
    }
    color_ref = _color_ref_from_openpyxl(font.color)
    if color_ref:
        if color_ref.get("type") != "rgb":
            color_ref["rgb"] = fcolor
        raw["color"] = color_ref
    for attr in ("charset", "family", "scheme", "outline", "shadow", "condense", "extend"):
        value = getattr(font, attr, None)
        if value is not None:
            raw[attr] = value
    return raw if len(raw) > 7 else None


def _apply_raw_font_kwargs(fk: dict, cd: dict) -> None:
    raw = _usable_raw_font(cd)
    if not raw:
        if cd.get("fcolor"):
            fk["color"] = cd["fcolor"]
        return

    color_ref = _usable_raw_color_ref(raw.get("color"), cd.get("fcolor"))
    if color_ref:
        color = _make_color_from_ref(color_ref)
        if color is not None:
            fk["color"] = color
    elif cd.get("fcolor"):
        fk["color"] = cd["fcolor"]

    for attr in ("charset", "family", "scheme", "outline", "shadow", "condense", "extend"):
        if attr in raw:
            fk[attr] = raw[attr]


def _make_pattern_fill_from_raw(raw: dict):
    from openpyxl.styles import PatternFill

    if raw.get("is_gradient"):
        import hashlib
        # Placeholder only. _inject_raw_fills patches the saved fill XML back to
        # the original gradient fill.
        color = "FF" + hashlib.sha1((raw.get("xml") or "").encode("utf-8")).hexdigest()[:6].upper()
        return PatternFill("solid", fgColor=color)

    pattern_type = raw.get("patternType") or "solid"
    fg = _make_color_from_ref(raw.get("fgColor") or {})
    bg = _make_color_from_ref(raw.get("bgColor") or {})
    kwargs = {"fill_type": pattern_type}
    if fg is not None:
        kwargs["fgColor"] = fg
    if bg is not None:
        kwargs["bgColor"] = bg
    return PatternFill(**kwargs)


def _fill_xml_has_color_reference(fill_xml: str) -> bool:
    import re
    return bool(re.search(r"<(?:fgColor|bgColor)\b[^>]*(?:\btheme=|\bindexed=|\bauto=|\btint=)", fill_xml))


def _fill_xml_should_preserve(fill_xml: str) -> bool:
    import re

    if _fill_xml_has_color_reference(fill_xml):
        return True
    if re.search(r"<gradientFill\b", fill_xml):
        return True
    pattern_m = re.search(r"<patternFill\b([^>]*)", fill_xml)
    if not pattern_m:
        return False
    pattern_type = _parse_xml_attrs(pattern_m.group(1)).get("patternType")
    if not pattern_type or pattern_type in ("none", "solid"):
        return False
    if pattern_type == "gray125":
        # Bare gray125 is openpyxl's built-in default fill slot; gray125 with
        # explicit colors is a real user fill and must be preserved.
        return "<fgColor" in fill_xml or "<bgColor" in fill_xml
    return True


def _extract_raw_fill_data(xlsx_path, sheet_file_map: dict) -> dict:
    """
    Extract raw fill XML per cell for fills that use theme/indexed/auto/tint
    references. The public API still exposes resolved RGB; this metadata is used
    only to keep the original OOXML color reference on save.
    """
    import zipfile, re

    result = {}
    try:
        with zipfile.ZipFile(str(xlsx_path), "r") as zf:
            if "xl/styles.xml" not in zf.namelist():
                return result
            styles_xml = zf.read("xl/styles.xml").decode("utf-8")
            fills_m = re.search(r"<fills\b[^>]*>.*?</fills>", styles_xml, re.DOTALL)
            xfs_m = re.search(r"<cellXfs\b[^>]*>.*?</cellXfs>", styles_xml, re.DOTALL)
            if not fills_m or not xfs_m:
                return result

            fills = _extract_xml_children(fills_m.group(0), "fill")
            style_to_fill: dict[int, int] = {}
            for idx, xf_m in enumerate(re.finditer(r"<xf\b([^>]*)/?>", xfs_m.group(0))):
                attrs = _parse_xml_attrs(xf_m.group(1))
                fill_id = attrs.get("fillId")
                if fill_id is not None:
                    style_to_fill[idx] = int(fill_id)

            fill_raw: dict[int, str] = {
                idx: xml for idx, xml in enumerate(fills)
                if _fill_xml_should_preserve(xml)
            }
            if not fill_raw:
                return result

            for sname, sheet_file in sheet_file_map.items():
                if sheet_file not in zf.namelist():
                    continue
                sheet_xml = zf.read(sheet_file).decode("utf-8")
                cells = {}
                for cell_m in re.finditer(r"<c\b([^>]*)", sheet_xml):
                    attrs = _parse_xml_attrs(cell_m.group(1))
                    coord = attrs.get("r")
                    style_idx = attrs.get("s")
                    if coord is None or style_idx is None:
                        continue
                    fill_id = style_to_fill.get(int(style_idx))
                    if fill_id in fill_raw:
                        cells[coord] = {"xml": fill_raw[fill_id]}
                if cells:
                    result[sname] = cells
    except Exception:
        pass
    return result


def _extract_sheet_view_attrs(xlsx_path, sheet_file_map: dict) -> dict:
    """Extract raw sheetView opening-tag attributes per sheet."""
    import zipfile, re

    result = {}
    try:
        with zipfile.ZipFile(str(xlsx_path), "r") as zf:
            for sname, sheet_file in sheet_file_map.items():
                if sheet_file not in zf.namelist():
                    continue
                sheet_xml = zf.read(sheet_file).decode("utf-8")
                m = re.search(r"<sheetView\b([^>]*)>", sheet_xml)
                if not m:
                    continue
                raw_attrs = m.group(1).strip()
                if raw_attrs.endswith("/"):
                    raw_attrs = raw_attrs[:-1].rstrip()
                result[sname] = {
                    "raw": raw_attrs,
                    "attrs": _parse_xml_attrs(raw_attrs),
                }
    except Exception:
        pass
    return result


def _extract_sheet_format_data(xlsx_path, sheet_file_map: dict) -> dict:
    """Extract raw worksheet root attrs, sheetFormatPr XML, and cols XML per sheet."""
    import zipfile, re

    result = {}
    try:
        with zipfile.ZipFile(str(xlsx_path), "r") as zf:
            for sname, sheet_file in sheet_file_map.items():
                if sheet_file not in zf.namelist():
                    continue
                sheet_xml = zf.read(sheet_file).decode("utf-8")
                root_m = re.search(r"<worksheet\b([^>]*)>", sheet_xml)
                sf_m = re.search(
                    r"<sheetFormatPr\b[^>]*/>|<sheetFormatPr\b[^>]*>.*?</sheetFormatPr>",
                    sheet_xml,
                    re.DOTALL,
                )
                cols_m = re.search(r"<cols\b[^>]*>.*?</cols>", sheet_xml, re.DOTALL)
                if root_m or sf_m or cols_m:
                    result[sname] = {
                        "root_attrs": root_m.group(1).strip() if root_m else "",
                        "sheetFormatPr": sf_m.group(0) if sf_m else None,
                        "cols": cols_m.group(0) if cols_m else None,
                    }
    except Exception:
        pass
    return result


def _serialize_sheet_view(sv, raw_view: dict | None = None) -> dict:
    attrs = {}
    for attr in getattr(type(sv), "__attrs__", ()):
        try:
            value = getattr(sv, attr)
        except Exception:
            continue
        if value is not None:
            attrs[attr] = value

    # Backwards-compatible alias used by older session data.
    if "zoomScale" in attrs:
        attrs["zoom"] = attrs["zoomScale"]
    if raw_view:
        attrs["_raw_attrs"] = raw_view.get("raw")
        attrs["_raw_attr_values"] = raw_view.get("attrs")
    return attrs


# ── Helpers ───────────────────────────────────────────────────────────────────

def _ser_border_side(side, theme_colors: list) -> dict | None:
    """Return compact border side dict, or None if no border style set."""
    if side is None or not side.border_style:
        return None
    d = {"style": side.border_style}
    color = _resolve_color(side.color, theme_colors)
    color_ref = _color_ref_from_openpyxl(side.color)
    if color_ref:
        color_ref["rgb"] = color
        d["_color_raw"] = color_ref
    # Suppress default black — it is the implicit border color
    if color and color not in ("FF000000", "00000000"):
        d["color"] = color
    return d


def _ser_validations(ws) -> list:
    result = []
    for dv in ws.data_validations.dataValidation:
        item = {"formula1": dv.formula1, "formula2": dv.formula2}
        for attr in getattr(type(dv), "__attrs__", ()):
            value = getattr(dv, attr, None)
            if attr == "sqref":
                value = str(value)
            if value is not None:
                item[attr] = value
        result.append(item)
    return result


def _ser_column_dimensions(ws) -> tuple[dict, dict, dict]:
    import openpyxl.utils

    widths = {}
    hidden = {}
    outline = {}
    for key, cd in ws.column_dimensions.items():
        try:
            start = cd.min or openpyxl.utils.column_index_from_string(key)
            end = cd.max or start
        except Exception:
            start = end = openpyxl.utils.column_index_from_string(key)
        for idx in range(start, end + 1):
            letter = openpyxl.utils.get_column_letter(idx)
            widths[letter] = cd.width
            if cd.hidden:
                hidden[letter] = True
            if cd.outlineLevel:
                outline[letter] = cd.outlineLevel
    return widths, hidden, outline


def _dimension_state(cw: dict | None, ch: dict | None, co: dict | None) -> dict:
    def _widths(values):
        result = {}
        for key, value in (values or {}).items():
            if value is None:
                result[str(key)] = None
            else:
                try:
                    result[str(key)] = round(float(value), 10)
                except Exception:
                    result[str(key)] = value
        return result

    return {
        "cw": _widths(cw),
        "ch": {str(k): bool(v) for k, v in (ch or {}).items() if v},
        "co": {str(k): int(v) for k, v in (co or {}).items() if v},
    }


def _normalize_dimension_state(state: dict | None) -> dict:
    state = state or {}
    return _dimension_state(state.get("cw"), state.get("ch"), state.get("co"))


def _make_border_side(sd):
    from openpyxl.styles import Side
    if not sd:
        return Side()
    from openpyxl.styles.colors import Color
    kw = {"border_style": sd["style"]}
    raw_color = _usable_raw_color_ref(sd.get("_color_raw"), sd.get("color"))
    if raw_color:
        color = _make_color_from_ref(raw_color)
        if color is not None:
            kw["color"] = color
    elif sd.get("color"):
        kw["color"] = Color(rgb=sd["color"])
    return Side(**kw)


def _xlsx_sheet_file_map(wb_xml: str, rels_xml: str) -> dict:
    """Return {sheet_name: zip_path} from workbook XML and its .rels XML."""
    import re
    from html import unescape
    rel_map = {}
    for m in re.finditer(r'<Relationship\b([^>]+)/>', rels_xml):
        attrs = m.group(1)
        id_m  = re.search(r'\bId="([^"]+)"', attrs)
        tgt_m = re.search(r'\bTarget="([^"]+)"', attrs)
        if id_m and tgt_m:
            rel_map[id_m.group(1)] = tgt_m.group(1)
    result = {}
    for m in re.finditer(r'<sheet\b[^>]+\bname="([^"]+)"[^>]+\br:id="([^"]+)"', wb_xml):
        raw_name, rid = m.group(1), m.group(2)
        sname = unescape(raw_name)
        if rid not in rel_map:
            continue
        t = rel_map[rid]
        # Normalize: /xl/worksheets/sheet1.xml or ../worksheets/sheet1.xml
        if t.startswith("/"):
            t = t.lstrip("/")           # /xl/... → xl/...
        elif t.startswith("../"):
            t = "xl/" + t[3:]           # ../worksheets/... → xl/worksheets/...
        elif not t.startswith("xl/"):
            t = "xl/" + t
        result[sname] = t
    return result


def _normalize_rel_target(target: str, prefix: str = "xl/") -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    if target.startswith("../"):
        # Both worksheet and drawing rels live one directory below xl/
        # (xl/worksheets/, xl/drawings/), so ../foo always resolves to xl/foo.
        return "xl/" + target[3:]
    if not target.startswith(prefix):
        return prefix + target
    return target


def _xlsx_parts(path: str) -> set[str]:
    import zipfile
    with zipfile.ZipFile(str(path), "r") as zf:
        return set(zf.namelist())


def inspect_xlsx_package(path: str) -> dict:
    """Return a compact, validation-oriented summary of an .xlsx package."""
    import re
    import xml.etree.ElementTree as ET
    import zipfile

    required = {"[Content_Types].xml", "xl/workbook.xml", "xl/_rels/workbook.xml.rels"}
    errors: list[str] = []
    warnings: list[str] = []
    features: dict[str, int] = {}
    parts: list[str] = []

    try:
        with zipfile.ZipFile(str(path), "r") as zf:
            bad = zf.testzip()
            if bad:
                errors.append(f"corrupt zip member: {bad}")
            parts = sorted(zf.namelist())
            missing = sorted(required - set(parts))
            if missing:
                errors.append(f"missing required parts: {missing}")

            xml_parts = [p for p in parts if p.endswith((".xml", ".rels"))]
            for part in xml_parts:
                try:
                    ET.fromstring(zf.read(part))
                except Exception as exc:
                    errors.append(f"invalid XML in {part}: {exc}")

            prefixes = {
                "drawings": "xl/drawings/",
                "charts": "xl/charts/",
                "media": "xl/media/",
                "vml_drawings": "xl/drawings/vmlDrawing",
                "pivot_tables": "xl/pivotTables/",
                "slicers": "xl/slicers/",
                "external_links": "xl/externalLinks/",
                "custom_xml": "customXml/",
            }
            for name, prefix in prefixes.items():
                features[name] = sum(1 for p in parts if p.startswith(prefix))
            features["vba_project"] = int("xl/vbaProject.bin" in parts)
            features["unknown_parts"] = sum(
                1 for p in parts
                if not p.startswith(("_rels/", "docProps/", "xl/", "customXml/"))
                and p != "[Content_Types].xml"
            )

            referenced_parts: set[str] = set()
            for rel_part in [p for p in parts if p.endswith(".rels")]:
                rel_base = ""
                if rel_part == "_rels/.rels":
                    rel_base = ""
                elif "/_rels/" in rel_part:
                    folder, rel_name = rel_part.split("/_rels/", 1)
                    rel_base = folder.rsplit("/", 1)[0] + "/" if "/" in folder else ""
                    rel_base += rel_name[:-5]
                try:
                    rel_xml = zf.read(rel_part).decode("utf-8")
                    for match in re.finditer(r'\bTarget="([^"]+)"', rel_xml):
                        target = match.group(1)
                        if target.startswith(("http://", "https://", "mailto:")):
                            continue
                        if target.startswith("/"):
                            referenced_parts.add(target.lstrip("/"))
                        else:
                            import posixpath
                            referenced_parts.add(posixpath.normpath(posixpath.join(posixpath.dirname(rel_base), target)))
                except Exception:
                    pass
            orphan_advanced = [
                p for p in parts
                if p.startswith(("xl/pivotTables/", "xl/externalLinks/"))
                and p not in referenced_parts
            ]
            if orphan_advanced:
                warnings.append(f"advanced parts are present but not referenced by relationships: {orphan_advanced[:8]}")
            if any(features.get(k, 0) for k in ("vml_drawings", "pivot_tables", "slicers", "external_links", "vba_project")):
                warnings.append("workbook contains advanced parts that are preserved best-effort only")
    except Exception as exc:
        errors.append(f"cannot read xlsx package: {exc}")

    return {
        "path": str(path),
        "valid": not errors,
        "errors": errors,
        "warnings": warnings,
        "part_count": len(parts),
        "features": features,
    }


def validate_xlsx(path: str) -> list[str]:
    """Raise ValueError if the saved workbook package is structurally broken."""
    report = inspect_xlsx_package(path)
    if not report["valid"]:
        raise ValueError("Saved workbook failed validation: " + "; ".join(report["errors"]))
    return report["warnings"]


def diff_xlsx_package(before: str, after: str) -> dict:
    """Compare ZIP package manifests for diagnostics after a save."""
    before_parts = _xlsx_parts(before)
    after_parts = _xlsx_parts(after)
    removed = sorted(before_parts - after_parts)
    added = sorted(after_parts - before_parts)
    return {
        "before_part_count": len(before_parts),
        "after_part_count": len(after_parts),
        "added": added,
        "removed": removed,
        "changed": bool(added or removed),
    }


def _restore_missing_package_parts(source_path: str | None, xlsx_path: str) -> str | None:
    """Best-effort copy of advanced OOXML parts that openpyxl dropped."""
    if not source_path or str(source_path) == str(xlsx_path):
        return None
    import os
    import zipfile

    if not os.path.exists(str(source_path)):
        return None

    tmp = str(xlsx_path) + ".~parts.tmp"
    def merge_relationships(current_xml: str, source_xml: str, restored_parts: list[str]) -> str:
        import posixpath
        import re

        restored = set(restored_parts)
        existing_ids = set(re.findall(r'\bId="([^"]+)"', current_xml))
        existing_targets = set(re.findall(r'\bTarget="([^"]+)"', current_xml))
        additions = []
        for match in re.finditer(r'<Relationship\b([^>]*)/>', source_xml):
            rel = match.group(0)
            attrs = match.group(1)
            target_m = re.search(r'\bTarget="([^"]+)"', attrs)
            id_m = re.search(r'\bId="([^"]+)"', attrs)
            if not target_m:
                continue
            target = target_m.group(1)
            norm = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
            if norm not in restored and "vbaProject.bin" not in norm:
                continue
            if target in existing_targets:
                continue
            if id_m and id_m.group(1) in existing_ids:
                next_id = 1
                while f"rId{next_id}" in existing_ids:
                    next_id += 1
                new_id = f"rId{next_id}"
                existing_ids.add(new_id)
                rel = re.sub(r'\bId="[^"]+"', f'Id="{new_id}"', rel, count=1)
            elif id_m:
                existing_ids.add(id_m.group(1))
            existing_targets.add(target)
            additions.append(rel)
        if additions:
            current_xml = current_xml.replace("</Relationships>", "".join(additions) + "</Relationships>", 1)
        return current_xml

    def merge_content_types(current_xml: str, source_xml: str, restored_parts: list[str]) -> str:
        import re
        workbook_override = re.search(r'<Override\b[^>]*\bPartName="/xl/workbook.xml"[^>]*/>', source_xml)
        if workbook_override and (
            "macroEnabled" in workbook_override.group(0)
            or "template" in workbook_override.group(0)
        ):
            current_xml = re.sub(
                r'<Override\b[^>]*\bPartName="/xl/workbook.xml"[^>]*/>',
                workbook_override.group(0),
                current_xml,
                count=1,
            )
        needed_exts = {p.rsplit(".", 1)[-1].lower() for p in restored_parts if "." in p}
        for match in re.finditer(r'<Default\b[^>]*\bExtension="([^"]+)"[^>]*/>', source_xml):
            ext = match.group(1).lower()
            if ext in needed_exts and f'Extension="{ext}"' not in current_xml:
                current_xml = current_xml.replace("</Types>", match.group(0) + "</Types>", 1)
        for part in restored_parts:
            part_name = "/" + part
            if f'PartName="{part_name}"' in current_xml:
                continue
            match = re.search(rf'<Override\b[^>]*\bPartName="{re.escape(part_name)}"[^>]*/>', source_xml)
            if match:
                current_xml = current_xml.replace("</Types>", match.group(0) + "</Types>", 1)
        return current_xml

    try:
        with zipfile.ZipFile(str(source_path), "r") as src, zipfile.ZipFile(str(xlsx_path), "r") as cur:
            current = set(cur.namelist())
            missing = [p for p in src.namelist() if p not in current]
            source_names = set(src.namelist())
            if not missing and "xl/vbaProject.bin" not in source_names:
                return None
            unsafe_prefixes = (
                "xl/worksheets/",      # regenerated sheets own their rel ids
                "xl/drawings/",        # handled by _inject_drawing_data
                "xl/media/",           # handled through drawing relationships
                "xl/charts/",          # handled through drawing relationships
                "xl/printerSettings/", # old sheet rels/pageSetup ids are not stable
            )
            restored = [
                p for p in missing
                if p not in {"xl/workbook.xml", "xl/styles.xml", "xl/sharedStrings.xml", "[Content_Types].xml", "xl/calcChain.xml"}
                and not p.startswith(unsafe_prefixes)
            ]
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in cur.infolist():
                    raw = cur.read(item.filename)
                    if item.filename == "[Content_Types].xml":
                        raw = merge_content_types(
                            raw.decode("utf-8"),
                            src.read("[Content_Types].xml").decode("utf-8"),
                            restored,
                        ).encode("utf-8")
                    elif item.filename == "xl/_rels/workbook.xml.rels" and "xl/_rels/workbook.xml.rels" in source_names:
                        raw = merge_relationships(
                            raw.decode("utf-8"),
                            src.read("xl/_rels/workbook.xml.rels").decode("utf-8"),
                            restored,
                        ).encode("utf-8")
                    zout.writestr(item, raw)
                for part in restored:
                    zout.writestr(part, src.read(part))
        os.replace(tmp, str(xlsx_path))
    except Exception as exc:
        if os.path.exists(tmp):
            os.remove(tmp)
        return f"advanced package part passthrough failed: {exc}"
    return None


def _extract_drawing_data(xlsx_path, sheet_file_map: dict) -> dict:
    """
    Extract drawing/chart/image files per sheet name.
    Returns {sname: {"drawing_xml": str, "drawing_rels": str, "files": {path: b64str}}}.
    """
    import zipfile, re, base64
    result = {}
    try:
        with zipfile.ZipFile(str(xlsx_path), "r") as zf:
            namelist = set(zf.namelist())
            for sname, sheet_file in sheet_file_map.items():
                # xl/worksheets/sheetN.xml → xl/worksheets/_rels/sheetN.xml.rels
                parts = sheet_file.rsplit("/", 1)
                sheet_rels_file = parts[0] + "/_rels/" + parts[1] + ".rels"
                if sheet_rels_file not in namelist:
                    continue
                rels_content = zf.read(sheet_rels_file).decode("utf-8")
                drawing_file = None
                for rm in re.finditer(r'<Relationship\b([^>]+)/>', rels_content):
                    attrs = rm.group(1)
                    type_m = re.search(r'\bType="([^"]+)"', attrs)
                    # Only DrawingML drawings — NOT vmlDrawing (comment shapes),
                    # which openpyxl regenerates itself on save.
                    if not type_m or not type_m.group(1).rstrip("/").endswith("/drawing"):
                        continue
                    tgt_m = re.search(r'\bTarget="([^"]+)"', attrs)
                    if tgt_m:
                        drawing_file = _normalize_rel_target(tgt_m.group(1), "xl/drawings/")
                        break
                if not drawing_file or drawing_file not in namelist:
                    continue

                sd = {
                    "drawing_file": drawing_file,
                    "drawing_xml": zf.read(drawing_file).decode("utf-8"),
                    "drawing_rels": None,
                    "files": {},
                }

                dr_rels_path = drawing_file.rsplit("/", 1)
                dr_rels_file = dr_rels_path[0] + "/_rels/" + dr_rels_path[1] + ".rels"
                if dr_rels_file in namelist:
                    dr_rels = zf.read(dr_rels_file).decode("utf-8")
                    sd["drawing_rels"] = dr_rels
                    for rm in re.finditer(r'<Relationship\b([^>]+)/>', dr_rels):
                        attrs = rm.group(1)
                        tgt_m = re.search(r'\bTarget="([^"]+)"', attrs)
                        if tgt_m:
                            tgt = _normalize_rel_target(tgt_m.group(1), "xl/")
                            if tgt in namelist:
                                sd["files"][tgt] = base64.b64encode(zf.read(tgt)).decode()

                result[sname] = sd
    except Exception:
        pass
    return result


def _shape_texts_from_drawing_xml(xml: str) -> list[str]:
    import re
    from html import unescape
    texts = []
    for block in re.findall(r"<a:t(?:\s[^>]*)?>(.*?)</a:t>", xml, flags=re.DOTALL):
        texts.append(unescape(re.sub(r"<[^>]+>", "", block)))
    return texts


def _extract_shape_inventory(drawing_data: dict) -> dict:
    """Build lightweight shape metadata from preserved DrawingML XML."""
    import re
    from html import unescape
    result: dict[str, list[dict]] = {}
    for sname, sd in drawing_data.items():
        xml = sd.get("drawing_xml") or ""
        shapes = []
        anchor_re = r"<(?:(?:xdr:)?)(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b.*?</(?:(?:xdr:)?)(?:twoCellAnchor|oneCellAnchor|absoluteAnchor)>"
        for idx, match in enumerate(re.finditer(anchor_re, xml, re.DOTALL), 1):
            anchor_xml = match.group(0)
            name_m = re.search(r'<(?:xdr:)?cNvPr\b[^>]*\bname="([^"]*)"', anchor_xml)
            id_m = re.search(r'<(?:xdr:)?cNvPr\b[^>]*\bid="([^"]*)"', anchor_xml)
            rel_m = re.search(r'\br:embed="([^"]+)"|\br:link="([^"]+)"', anchor_xml)
            kind = "shape"
            if re.search(r"<(?:xdr:)?pic\b", anchor_xml):
                kind = "picture"
            elif "graphicData" in anchor_xml and "/chart" in anchor_xml:
                kind = "chart"
            elif re.search(r"<(?:xdr:)?sp\b", anchor_xml):
                kind = "shape"
            shapes.append({
                "index": idx,
                "id": id_m.group(1) if id_m else None,
                "name": unescape(name_m.group(1)) if name_m else None,
                "type": kind,
                "text": "".join(_shape_texts_from_drawing_xml(anchor_xml)) or None,
                "relationship_id": next((g for g in (rel_m.groups() if rel_m else ()) if g), None),
            })
        if shapes:
            result[sname] = shapes
    return result


def _inject_drawing_data(xlsx_path: str, drawing_data: dict,
                         sheet_name_to_new_file: dict) -> str | None:
    """
    Inject preserved DrawingML drawings, charts, and media into a saved xlsx file.
    sheet_name_to_new_file: {sheet_name: "xl/worksheets/sheetN.xml"} in the NEW file.
    """
    import zipfile, re, os, base64
    if not drawing_data:
        return None
    tmp = str(xlsx_path) + ".~draw.tmp"
    try:
        with zipfile.ZipFile(str(xlsx_path), "r") as zin:
            existing = set(zin.namelist())
            extra_files: dict[str, bytes] = {}
            sheet_xml_patches: dict[str, dict] = {}

            for sname, sd in drawing_data.items():
                new_sheet_file = sheet_name_to_new_file.get(sname)
                old_drawing = sd.get("drawing_file")
                drawing_xml = sd.get("drawing_xml")
                if not new_sheet_file or not old_drawing or not drawing_xml:
                    continue

                # Keep original DrawingML part names/relationships. Excel desktop
                # is stricter than XML validators; remapping complex drawing
                # packages can break hidden cross-part references.
                extra_files[old_drawing] = drawing_xml.encode("utf-8")

                old_rels_content = sd.get("drawing_rels")
                if old_rels_content:
                    dr_rels_path = old_drawing.rsplit("/", 1)
                    extra_files[dr_rels_path[0] + "/_rels/" + dr_rels_path[1] + ".rels"] = old_rels_content.encode("utf-8")

                for fp, payload in (sd.get("files") or {}).items():
                    extra_files[fp] = base64.b64decode(payload)

                parts = new_sheet_file.rsplit("/", 1)
                new_sheet_rels = parts[0] + "/_rels/" + parts[1] + ".rels"
                rid = "rId1"
                if new_sheet_rels in existing:
                    rels_now = zin.read(new_sheet_rels).decode("utf-8")
                    used = [int(m) for m in re.findall(r'\bId="rId(\d+)"', rels_now)]
                    rid = f"rId{max(used, default=0) + 1}"

                sheet_xml_patches[new_sheet_file] = {
                    "drawing_rId": rid,
                    "sheet_rels_file": new_sheet_rels,
                    "rel_target": "../drawings/" + old_drawing.rsplit("/", 1)[1],
                }

            if not sheet_xml_patches:
                return None

            ct_additions = set()
            for fp in extra_files:
                ext = fp.rsplit(".", 1)[-1].lower()
                if ext == "xml":
                    if "/charts/" in fp:
                        ct_additions.add(("Override", fp, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"))
                    elif "/drawings/drawing" in fp:
                        ct_additions.add(("Override", fp, "application/vnd.openxmlformats-officedocument.drawing+xml"))
                elif ext == "png":
                    ct_additions.add(("Default", "png", "image/png"))
                elif ext in {"jpg", "jpeg"}:
                    ct_additions.add(("Default", ext, "image/jpeg"))
                elif ext == "gif":
                    ct_additions.add(("Default", "gif", "image/gif"))
                elif ext == "emf":
                    ct_additions.add(("Default", "emf", "image/x-emf"))
                elif ext == "wmf":
                    ct_additions.add(("Default", "wmf", "image/x-wmf"))

            draw_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
            rels_patches = {
                patch["sheet_rels_file"]: (
                    f'<Relationship Id="{patch["drawing_rId"]}" '
                    f'Type="{draw_type}" Target="{patch["rel_target"]}"/>'
                )
                for patch in sheet_xml_patches.values()
            }

            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    raw = zin.read(item.filename)

                    if item.filename in sheet_xml_patches:
                        patch = sheet_xml_patches[item.filename]
                        content = raw.decode("utf-8")
                        if "<drawing" not in content:
                            rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                            tag = f'<drawing xmlns:r="{rns}" r:id="{patch["drawing_rId"]}"/>'
                            anchor = re.search(
                                r"<(?:legacyDrawing|legacyDrawingHF|picture|oleObjects"
                                r"|controls|webPublishItems|tableParts|extLst)\b",
                                content,
                            )
                            pos = anchor.start() if anchor else content.rfind("</worksheet>")
                            content = content[:pos] + tag + content[pos:]
                        raw = content.encode("utf-8")

                    elif item.filename in rels_patches:
                        content = raw.decode("utf-8")
                        if rels_patches[item.filename] not in content:
                            content = content.replace("</Relationships>", rels_patches[item.filename] + "</Relationships>", 1)
                        raw = content.encode("utf-8")

                    elif item.filename == "[Content_Types].xml" and ct_additions:
                        content = raw.decode("utf-8")
                        for kind, part_or_ext, content_type in ct_additions:
                            if kind == "Default":
                                if f'Extension="{part_or_ext}"' not in content:
                                    content = content.replace(
                                        "</Types>",
                                        f'<Default Extension="{part_or_ext}" ContentType="{content_type}"/></Types>', 1)
                            elif f'PartName="/{part_or_ext}"' not in content:
                                content = content.replace(
                                    "</Types>",
                                    f'<Override PartName="/{part_or_ext}" ContentType="{content_type}"/></Types>', 1)
                        raw = content.encode("utf-8")

                    if item.filename in extra_files:
                        raw = extra_files[item.filename]
                    zout.writestr(item, raw)

                for fp, fb in extra_files.items():
                    if fp not in existing:
                        zout.writestr(fp, fb)

                ns = "http://schemas.openxmlformats.org/package/2006/relationships"
                for rels_file, rel_entry in rels_patches.items():
                    if rels_file not in existing:
                        zout.writestr(rels_file, f'<Relationships xmlns="{ns}">{rel_entry}</Relationships>'.encode("utf-8"))

        os.replace(tmp, str(xlsx_path))
    except Exception as e:
        if os.path.exists(tmp):
            os.remove(tmp)
        return f"drawings/charts/images passthrough failed: {e}"
    return None

def _extract_cf_xml(xlsx_path, sheet_names: list) -> dict:
    """
    Extract raw <conditionalFormatting> XML blocks per sheet, plus the workbook's
    <dxfs> section (differential styles referenced by CF rules).
    Returns {sname: [block, ...], "__dxfs__": "<dxfs>...</dxfs>"}.
    """
    import zipfile, re
    result = {}
    try:
        with zipfile.ZipFile(str(xlsx_path), "r") as zf:
            wb_xml   = zf.read("xl/workbook.xml").decode("utf-8")
            rels_xml = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
            sheet_file_map = _xlsx_sheet_file_map(wb_xml, rels_xml)
            for sname in sheet_names:
                fp = sheet_file_map.get(sname)
                if not fp or fp not in zf.namelist():
                    continue
                content = zf.read(fp).decode("utf-8")
                blocks = re.findall(
                    r"<conditionalFormatting(?:\s[^>]*)?>.*?</conditionalFormatting>",
                    content, re.DOTALL)
                if blocks:
                    result[sname] = blocks
            # Extract dxfs section from styles.xml (needed for dxfId refs in CF rules)
            if result and "xl/styles.xml" in zf.namelist():
                styles_xml = zf.read("xl/styles.xml").decode("utf-8")
                dxfs_m = re.search(r"<dxfs\b[^>]*>.*?</dxfs>", styles_xml, re.DOTALL)
                if dxfs_m:
                    result["__dxfs__"] = dxfs_m.group(0)
    except Exception:
        pass
    return result


def _extract_data_validations_xml(xlsx_path, sheet_file_map: dict) -> dict:
    """Extract raw <dataValidations> XML per sheet."""
    import zipfile, re
    result = {}
    try:
        with zipfile.ZipFile(str(xlsx_path), "r") as zf:
            for sname, sheet_file in sheet_file_map.items():
                if sheet_file not in zf.namelist():
                    continue
                content = zf.read(sheet_file).decode("utf-8")
                m = re.search(
                    r"<dataValidations\b[^>]*>.*?</dataValidations>",
                    content,
                    re.DOTALL,
                )
                if m:
                    result[sname] = m.group(0)
    except Exception:
        pass
    return result


def _inject_cf_xml(xlsx_path: str, sheet_cf: dict) -> str | None:
    """
    Patch a saved xlsx file by:
    1. Injecting stored CF XML into each sheet's XML.
    2. Replacing the <dxfs/> section in styles.xml with the original one (needed for dxfId refs).
    """
    import zipfile, re, os
    if not sheet_cf:
        return
    dxfs_xml = sheet_cf.pop("__dxfs__", None)
    sheet_cf_only = {k: v for k, v in sheet_cf.items() if v}
    if not sheet_cf_only and not dxfs_xml:
        return
    tmp = str(xlsx_path) + ".~tmp"
    try:
        with zipfile.ZipFile(str(xlsx_path), "r") as zin:
            wb_xml   = zin.read("xl/workbook.xml").decode("utf-8")
            rels_xml = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
            sheet_file_map = _xlsx_sheet_file_map(wb_xml, rels_xml)
            to_patch = {sheet_file_map[sn]: blocks
                        for sn, blocks in sheet_cf_only.items()
                        if sn in sheet_file_map}
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    raw = zin.read(item.filename)
                    if item.filename in to_patch:
                        content = raw.decode("utf-8")
                        injection = "".join(to_patch[item.filename])
                        # CT_Worksheet order: conditionalFormatting must precede
                        # dataValidations/hyperlinks/pageMargins/… — insert before
                        # the first such element, not at the end of the sheet.
                        anchor = re.search(
                            r"<(?:dataValidations|hyperlinks|printOptions|pageMargins"
                            r"|pageSetup\b|headerFooter|rowBreaks|colBreaks|drawing\b"
                            r"|legacyDrawing|tableParts|extLst)\b",
                            content,
                        )
                        pos = anchor.start() if anchor else content.rfind("</worksheet>")
                        content = content[:pos] + injection + content[pos:]
                        raw = content.encode("utf-8")
                    elif item.filename == "xl/styles.xml" and dxfs_xml:
                        content = raw.decode("utf-8")
                        # Replace existing <dxfs> block, or inject at its schema
                        # position: dxfs must precede tableStyles/colors/extLst
                        # (CT_Stylesheet order) — Excel refuses the file otherwise.
                        if re.search(r"<dxfs\b", content):
                            content = re.sub(
                                r"<dxfs\b[^>]*/?>|<dxfs\b[^>]*>.*?</dxfs>",
                                dxfs_xml, content, count=1, flags=re.DOTALL)
                        else:
                            anchor = re.search(r"<(?:tableStyles|colors|extLst)\b", content)
                            pos = anchor.start() if anchor else content.rfind("</styleSheet>")
                            content = content[:pos] + dxfs_xml + content[pos:]
                        raw = content.encode("utf-8")
                    zout.writestr(item, raw)
        os.replace(tmp, str(xlsx_path))
    except Exception as e:
        if os.path.exists(tmp):
            os.remove(tmp)
        return f"conditional formatting passthrough failed: {e}"
    return None


def _inject_data_validations_xml(xlsx_path: str, sheet_validations: dict) -> str | None:
    """Patch saved worksheet XML with original dataValidations XML."""
    import zipfile, re, os
    if not sheet_validations:
        return
    tmp = str(xlsx_path) + ".~validations.tmp"
    try:
        with zipfile.ZipFile(str(xlsx_path), "r") as zin:
            wb_xml = zin.read("xl/workbook.xml").decode("utf-8")
            rels_xml = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
            sheet_file_map = _xlsx_sheet_file_map(wb_xml, rels_xml)
            file_to_xml = {
                sheet_file_map[sname]: xml
                for sname, xml in sheet_validations.items()
                if sname in sheet_file_map
            }
            if not file_to_xml:
                return

            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    raw = zin.read(item.filename)
                    dv_xml = file_to_xml.get(item.filename)
                    if dv_xml:
                        content = raw.decode("utf-8")
                        if re.search(r"<dataValidations\b", content):
                            content = re.sub(
                                r"<dataValidations\b[^>]*>.*?</dataValidations>",
                                dv_xml,
                                content,
                                count=1,
                                flags=re.DOTALL,
                            )
                        else:
                            content = re.sub(
                                r"(<hyperlinks\b|<pageMargins\b|</worksheet>)",
                                dv_xml + r"\1",
                                content,
                                count=1,
                            )
                        raw = content.encode("utf-8")
                    zout.writestr(item, raw)
        os.replace(tmp, str(xlsx_path))
    except Exception as e:
        if os.path.exists(tmp):
            os.remove(tmp)
        return f"data validations passthrough failed: {e}"
    return None


def _inject_raw_fills(xlsx_path: str, data: dict) -> str | None:
    """Patch saved styles.xml fill entries back to original raw OOXML fills."""
    import zipfile, re, os
    from openpyxl.utils import get_column_letter

    sheet_targets = {}
    for sd in data.get("sheets", []):
        targets = {}
        for r_idx, row_data in enumerate(sd.get("rows", []), 1):
            for c_idx, cd in enumerate(row_data.get("cells", []), 1):
                if cd.get("merge") == "slave":
                    continue
                raw = _usable_raw_fill(cd)
                if raw and raw.get("xml"):
                    targets[f"{get_column_letter(c_idx)}{r_idx}"] = raw["xml"]
        if targets:
            sheet_targets[sd["name"]] = targets

    if not sheet_targets:
        return

    tmp = str(xlsx_path) + ".~fills.tmp"
    try:
        with zipfile.ZipFile(str(xlsx_path), "r") as zin:
            namelist = set(zin.namelist())
            if "xl/styles.xml" not in namelist:
                return
            wb_xml = zin.read("xl/workbook.xml").decode("utf-8")
            rels_xml = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
            sheet_file_map = _xlsx_sheet_file_map(wb_xml, rels_xml)

            styles_xml = zin.read("xl/styles.xml").decode("utf-8")
            fills_m = re.search(r"<fills\b[^>]*>.*?</fills>", styles_xml, re.DOTALL)
            xfs_m = re.search(r"<cellXfs\b[^>]*>.*?</cellXfs>", styles_xml, re.DOTALL)
            if not fills_m or not xfs_m:
                return
            fills = _extract_xml_children(fills_m.group(0), "fill")
            style_to_fill: dict[int, int] = {}
            for idx, xf_m in enumerate(re.finditer(r"<xf\b([^>]*)/?>", xfs_m.group(0))):
                attrs = _parse_xml_attrs(xf_m.group(1))
                if attrs.get("fillId") is not None:
                    style_to_fill[idx] = int(attrs["fillId"])

            fill_patches: dict[int, str] = {}
            conflicts: set[int] = set()
            for sname, targets in sheet_targets.items():
                sheet_file = sheet_file_map.get(sname)
                if not sheet_file or sheet_file not in namelist:
                    continue
                sheet_xml = zin.read(sheet_file).decode("utf-8")
                coord_to_style = {}
                for cell_m in re.finditer(r"<c\b([^>]*)", sheet_xml):
                    attrs = _parse_xml_attrs(cell_m.group(1))
                    if attrs.get("r") and attrs.get("s") is not None:
                        coord_to_style[attrs["r"]] = int(attrs["s"])
                for coord, raw_xml in targets.items():
                    style_idx = coord_to_style.get(coord)
                    if style_idx is None:
                        continue
                    fill_id = style_to_fill.get(style_idx)
                    if fill_id is None or not (0 <= fill_id < len(fills)):
                        continue
                    existing = fill_patches.get(fill_id)
                    if existing is not None and existing != raw_xml:
                        conflicts.add(fill_id)
                        continue
                    fill_patches[fill_id] = raw_xml

            for fill_id in conflicts:
                fill_patches.pop(fill_id, None)
            if not fill_patches:
                return

            patched_fills = [
                fill_patches.get(idx, fill_xml)
                for idx, fill_xml in enumerate(fills)
            ]

            def _replace_fills(match):
                open_tag = re.match(r"<fills\b[^>]*>", match.group(0)).group(0)
                return open_tag + "".join(patched_fills) + "</fills>"

            patched_styles = re.sub(
                r"<fills\b[^>]*>.*?</fills>",
                _replace_fills,
                styles_xml,
                count=1,
                flags=re.DOTALL,
            )

            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    raw = zin.read(item.filename)
                    if item.filename == "xl/styles.xml":
                        raw = patched_styles.encode("utf-8")
                    zout.writestr(item, raw)
        os.replace(tmp, str(xlsx_path))
    except Exception as e:
        if os.path.exists(tmp):
            os.remove(tmp)
        return f"raw fills passthrough failed: {e}"
    return None


def _inject_sheet_view_attrs(xlsx_path: str, data: dict) -> str | None:
    """Patch saved worksheet XML with original sheetView attributes."""
    import zipfile, re, os

    sheet_attrs = {
        sd["name"]: (sd.get("sheet_view") or {}).get("_raw_attrs")
        for sd in data.get("sheets", [])
        if (sd.get("sheet_view") or {}).get("_raw_attrs") is not None
    }
    if not sheet_attrs:
        return

    tmp = str(xlsx_path) + ".~sheetviews.tmp"
    try:
        with zipfile.ZipFile(str(xlsx_path), "r") as zin:
            wb_xml = zin.read("xl/workbook.xml").decode("utf-8")
            rels_xml = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
            sheet_file_map = _xlsx_sheet_file_map(wb_xml, rels_xml)
            file_to_attrs = {
                sheet_file_map[sname]: attrs
                for sname, attrs in sheet_attrs.items()
                if sname in sheet_file_map
            }
            if not file_to_attrs:
                return

            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    raw = zin.read(item.filename)
                    attrs = file_to_attrs.get(item.filename)
                    if attrs is not None:
                        content = raw.decode("utf-8")

                        def _replace(match):
                            current_attrs = match.group(1).rstrip()
                            self_closing = current_attrs.endswith("/")
                            slash = "/" if self_closing else ""
                            sep = " " if attrs else ""
                            return f"<sheetView{sep}{attrs}{slash}>"

                        content = re.sub(
                            r"<sheetView\b([^>]*)>",
                            _replace,
                            content,
                            count=1,
                        )
                        raw = content.encode("utf-8")
                    zout.writestr(item, raw)
        os.replace(tmp, str(xlsx_path))
    except Exception as e:
        if os.path.exists(tmp):
            os.remove(tmp)
        return f"sheetView attrs passthrough failed: {e}"
    return None


def _inject_raw_cols(xlsx_path: str, data: dict) -> str | None:
    """Patch saved worksheet XML with original <cols> ranges when dimensions are unchanged."""
    import zipfile, re, os

    sheet_cols = {}
    for sd in data.get("sheets", []):
        raw = sd.get("_cols_raw") or {}
        raw_xml = raw.get("xml")
        if not raw_xml:
            continue
        current_state = _dimension_state(sd.get("cw"), sd.get("ch"), sd.get("co"))
        original_state = _normalize_dimension_state(raw.get("state"))
        if current_state == original_state:
            sheet_cols[sd["name"]] = raw_xml

    if not sheet_cols:
        return

    tmp = str(xlsx_path) + ".~cols.tmp"
    try:
        with zipfile.ZipFile(str(xlsx_path), "r") as zin:
            wb_xml = zin.read("xl/workbook.xml").decode("utf-8")
            rels_xml = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
            sheet_file_map = _xlsx_sheet_file_map(wb_xml, rels_xml)
            file_to_cols = {
                sheet_file_map[sname]: cols_xml
                for sname, cols_xml in sheet_cols.items()
                if sname in sheet_file_map
            }
            if not file_to_cols:
                return

            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    raw = zin.read(item.filename)
                    cols_xml = file_to_cols.get(item.filename)
                    if cols_xml is not None:
                        content = raw.decode("utf-8")
                        if re.search(r"<cols\b[^>]*>.*?</cols>", content, re.DOTALL):
                            content = re.sub(
                                r"<cols\b[^>]*>.*?</cols>",
                                cols_xml,
                                content,
                                count=1,
                                flags=re.DOTALL,
                            )
                        else:
                            content = re.sub(
                                r"(<sheetData\b)",
                                cols_xml + r"\1",
                                content,
                                count=1,
                            )
                        raw = content.encode("utf-8")
                    zout.writestr(item, raw)
        os.replace(tmp, str(xlsx_path))
    except Exception as e:
        if os.path.exists(tmp):
            os.remove(tmp)
        return f"cols passthrough failed: {e}"
    return None


def _xml_prefixed_attrs(xml: str) -> set[str]:
    import re
    return set(re.findall(r"\b([A-Za-z_][\w.-]*):[A-Za-z_][\w.-]*=", xml))


def _inject_missing_root_attrs(content: str, needed_attrs: dict[str, str]) -> str:
    import re

    root_m = re.search(r"<worksheet\b([^>]*)>", content)
    if not root_m:
        return content
    current = root_m.group(1)
    additions = []
    for key, value in needed_attrs.items():
        if re.search(rf"\b{re.escape(key)}=", current):
            continue
        additions.append(f'{key}="{value}"')
    if not additions:
        return content
    insert = " " + " ".join(additions)
    return content[:root_m.end() - 1] + insert + content[root_m.end() - 1:]


def _inject_sheet_format_pr(xlsx_path: str, data: dict) -> str | None:
    """Restore raw sheetFormatPr XML, including extension attrs like x14ac:dyDescent."""
    import zipfile, re, os

    sheet_data = {
        sd["name"]: sd.get("_sheet_format_pr")
        for sd in data.get("sheets", [])
        if sd.get("_sheet_format_pr")
    }
    if not sheet_data:
        return

    tmp = str(xlsx_path) + ".~sheetformat.tmp"
    try:
        with zipfile.ZipFile(str(xlsx_path), "r") as zin:
            wb_xml = zin.read("xl/workbook.xml").decode("utf-8")
            rels_xml = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
            sheet_file_map = _xlsx_sheet_file_map(wb_xml, rels_xml)
            file_to_data = {
                sheet_file_map[sname]: sf_data
                for sname, sf_data in sheet_data.items()
                if sname in sheet_file_map
            }
            if not file_to_data:
                return

            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    raw = zin.read(item.filename)
                    sf_data = file_to_data.get(item.filename)
                    if sf_data:
                        content = raw.decode("utf-8")
                        raw_sf = sf_data.get("sheetFormatPr")
                        root_attrs = _parse_xml_attrs(sf_data.get("root_attrs") or "")
                        needed = {}
                        root_prefixes = {
                            p for p in _xml_prefixed_attrs(sf_data.get("root_attrs") or "")
                            if p not in {"xmlns", "mc"}
                        }
                        for prefix in root_prefixes:
                            ns_key = f"xmlns:{prefix}"
                            if root_attrs.get(ns_key):
                                needed[ns_key] = root_attrs[ns_key]
                            attr_key = next((k for k in root_attrs if k.startswith(prefix + ":")), None)
                            if attr_key:
                                needed[attr_key] = root_attrs[attr_key]
                        if raw_sf:
                            prefixes = _xml_prefixed_attrs(raw_sf)
                            for prefix in prefixes:
                                ns_key = f"xmlns:{prefix}"
                                if root_attrs.get(ns_key):
                                    needed[ns_key] = root_attrs[ns_key]
                            ignorable_prefixes = prefixes | root_prefixes
                            if ignorable_prefixes and root_attrs.get("xmlns:mc") and root_attrs.get("mc:Ignorable"):
                                # mc:Ignorable may only list prefixes that are
                                # actually declared in the new root — Excel
                                # refuses to open the file otherwise.
                                current_m = re.search(r"<worksheet\b([^>]*)>", content)
                                current_attrs = _parse_xml_attrs(current_m.group(1)) if current_m else {}
                                declared = {k[6:] for k in needed if k.startswith("xmlns:")}
                                declared |= {k[6:] for k in current_attrs if k.startswith("xmlns:")}
                                keep = [t for t in root_attrs["mc:Ignorable"].split()
                                        if t in declared]
                                if keep:
                                    needed["xmlns:mc"] = root_attrs["xmlns:mc"]
                                    needed["mc:Ignorable"] = " ".join(keep)
                            if re.search(r"<sheetFormatPr\b", content):
                                content = re.sub(
                                    r"<sheetFormatPr\b[^>]*/>|<sheetFormatPr\b[^>]*>.*?</sheetFormatPr>",
                                    raw_sf,
                                    content,
                                    count=1,
                                    flags=re.DOTALL,
                                )
                            else:
                                content = re.sub(
                                    r"(<sheetData\b)",
                                    raw_sf + r"\1",
                                    content,
                                    count=1,
                                )
                        if needed:
                            content = _inject_missing_root_attrs(content, needed)
                        raw = content.encode("utf-8")
                    zout.writestr(item, raw)
        os.replace(tmp, str(xlsx_path))
    except Exception as e:
        if os.path.exists(tmp):
            os.remove(tmp)
        return f"sheetFormatPr passthrough failed: {e}"
    return None


# ── Serialize ─────────────────────────────────────────────────────────────────

def serialize_excel(uri: str, sheet_name: str | None = None) -> dict:
    """
    Serialize an Excel file to a metadata dict.

    Each cell carries: value, fill RGB, bold/italic/size/font-color,
    wrap/halign/valign, number format, merge info, border sides.
    Sheet carries: freeze_panes ref, data_validations.
    Merge origin → {rowspan, colspan, r1,c1,r2,c2}.
    Merge slave  → "slave" (skipped during reconstruct).
    """
    import openpyxl

    path = uri_to_path(uri)
    keep_vba = path.suffix.lower() in {".xlsm", ".xltm"}
    wb = openpyxl.load_workbook(str(path), keep_vba=keep_vba)
    raw_theme = getattr(wb, "loaded_theme", None)
    theme_xml = None
    if raw_theme:
        import base64
        if isinstance(raw_theme, str):
            raw_theme = raw_theme.encode("utf-8")
        theme_xml = base64.b64encode(raw_theme).decode("ascii")

    # Extract theme colors once — used to resolve all theme-type cell colors
    theme_colors = _wb_theme_colors(wb)

    try:
        import zipfile as _zf
        with _zf.ZipFile(str(path), "r") as _z:
            _wb_xml = _z.read("xl/workbook.xml").decode("utf-8")
            _rels_xml = _z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        _sfm = _xlsx_sheet_file_map(_wb_xml, _rels_xml)
    except Exception:
        _sfm = {}
    raw_fill_data = _extract_raw_fill_data(path, _sfm)
    raw_sheet_views = _extract_sheet_view_attrs(path, _sfm)
    raw_sheet_formats = _extract_sheet_format_data(path, _sfm)
    raw_data_validations = _extract_data_validations_xml(path, _sfm)

    names = [sheet_name] if sheet_name else wb.sheetnames
    sheets = []

    for sname in names:
        if sname not in wb.sheetnames:
            raise ValueError(f"Sheet '{sname}' not found. Available: {wb.sheetnames}")
        ws = wb[sname]
        raw_fills_for_sheet = raw_fill_data.get(sname, {})

        # Build merge map
        merged_map: dict = {}
        for mg in ws.merged_cells.ranges:
            merged_map[(mg.min_row, mg.min_col)] = {
                "r1": mg.min_row - 1, "c1": mg.min_col - 1,  # 0-based
                "r2": mg.max_row - 1, "c2": mg.max_col - 1,  # 0-based
                "rowspan": mg.max_row - mg.min_row + 1,
                "colspan": mg.max_col - mg.min_col + 1,
            }
            for r in range(mg.min_row, mg.max_row + 1):
                for c in range(mg.min_col, mg.max_col + 1):
                    if not (r == mg.min_row and c == mg.min_col):
                        merged_map[(r, c)] = "slave"

        rows = []
        for row in ws.iter_rows():
            rh = ws.row_dimensions[row[0].row].height
            cells = []
            for cell in row:
                mi = merged_map.get((cell.row, cell.column), {})

                fill_rgb = None
                fill_raw = None
                if cell.fill:
                    if cell.fill.fill_type == "solid":
                        fill_rgb = _resolve_color(cell.fill.fgColor, theme_colors)
                    raw_cell_fill = raw_fills_for_sheet.get(cell.coordinate)
                    if raw_cell_fill:
                        fill_raw = dict(raw_cell_fill)
                        fill_raw["rgb"] = fill_rgb
                        fill_raw["is_gradient"] = "gradientFill" in (fill_raw.get("xml") or "")
                        fill_raw["patternType"] = (
                            cell.fill.fill_type
                            if not fill_raw["is_gradient"]
                            else None
                        )
                        fg_ref = _color_ref_from_openpyxl(getattr(cell.fill, "fgColor", None))
                        bg_ref = _color_ref_from_openpyxl(getattr(cell.fill, "bgColor", None))
                        if fg_ref:
                            fill_raw["fgColor"] = fg_ref
                        if bg_ref:
                            fill_raw["bgColor"] = bg_ref

                fcolor = _resolve_color(
                    cell.font.color if cell.font else None, theme_colors
                )
                # Suppress default black font color (no need to store it)
                if fcolor in ("FF000000", "00000000"):
                    fcolor = None
                font_raw = _font_raw_from_openpyxl(cell.font, fcolor)

                bdr = {}
                if cell.border:
                    for attr in ("top", "bottom", "left", "right", "diagonal"):
                        sd = _ser_border_side(getattr(cell.border, attr), theme_colors)
                        if sd:
                            bdr[attr] = sd
                    if cell.border.diagonalUp:
                        bdr["diagonalUp"] = True
                    if cell.border.diagonalDown:
                        bdr["diagonalDown"] = True

                aln = cell.alignment
                cell_data = {
                    "v":       cell.value,
                    "fill":    fill_rgb,
                    "bold":    bool(cell.font.bold)           if cell.font else False,
                    "italic":  bool(cell.font.italic)         if cell.font else False,
                    "size":    cell.font.size                 if cell.font else None,
                    "font":    cell.font.name                 if cell.font else None,
                    "fcolor":  fcolor,
                    "uline":   cell.font.underline            if cell.font else None,
                    "strike":  bool(cell.font.strike)         if cell.font else False,
                    "vAlign":  cell.font.vertAlign            if cell.font else None,
                    "wrap":    bool(aln.wrap_text)            if aln else False,
                    "halign":  aln.horizontal                 if aln else None,
                    "valign":  aln.vertical                   if aln else None,
                    "rot":     aln.text_rotation              if aln else None,
                    "indent":  aln.indent                     if aln else None,
                    "shrink":  bool(aln.shrink_to_fit)        if aln else False,
                    "numfmt":  cell.number_format,
                    "merge":   mi,
                    "border":  bdr,
                    "locked":  bool(cell.protection.locked)   if cell.protection else True,
                    "hidden_cell": bool(cell.protection.hidden) if cell.protection else False,
                }
                # Disambiguate literal text that LOOKS like a formula ("=…"):
                # without this marker a text cell would silently turn into a
                # broken formula on reconstruct.
                if (isinstance(cell.value, str) and cell.value.startswith("=")
                        and cell.data_type != "f"):
                    cell_data["dt"] = "s"
                # quotePrefix style flag (Excel's leading-apostrophe text marker)
                try:
                    if cell._style is not None and cell._style.quotePrefix:
                        cell_data["qp"] = True
                except Exception:
                    pass
                if fill_raw:
                    cell_data["_fill_raw"] = fill_raw
                if font_raw:
                    cell_data["_font_raw"] = font_raw
                cells.append(cell_data)
            rd = ws.row_dimensions[row[0].row]
            rows.append({
                "h":       rh,
                "hidden":  bool(rd.hidden),
                "outline": rd.outlineLevel or 0,
                "cells":   cells,
            })

        sv = ws.sheet_view
        col_widths, col_hidden, col_outline = _ser_column_dimensions(ws)
        col_state = _dimension_state(col_widths, col_hidden, col_outline)

        # Tab color
        tc = None
        if ws.sheet_properties and ws.sheet_properties.tabColor:
            tc = _resolve_color(ws.sheet_properties.tabColor, theme_colors)

        # Print settings
        ps = ws.page_setup
        pm = ws.page_margins
        def _safe(obj, attr):
            try: return getattr(obj, attr)
            except Exception: return None
        # fitToPage lives in sheetPr/pageSetUpPr, not in pageSetup
        fit_to_page = None
        try:
            pspr = ws.sheet_properties.pageSetUpPr
            fit_to_page = pspr.fitToPage if pspr else None
        except Exception:
            pass
        page_setup = {k: v for k, v in {
            "orientation": _safe(ps, "orientation"),
            "paperSize":   _safe(ps, "paperSize"),
            "fitToPage":   fit_to_page,
            "fitToWidth":  _safe(ps, "fitToWidth"),
            "fitToHeight": _safe(ps, "fitToHeight"),
            "scale":       _safe(ps, "scale"),
        }.items() if v is not None}
        page_margins = {k: getattr(pm, k) for k in
                        ("left", "right", "top", "bottom", "header", "footer")
                        if getattr(pm, k, None) is not None}

        # Sheet protection
        prot = ws.protection
        protection = None
        if prot.sheet:
            protection = {
                "password":            prot.password,
                "selectLockedCells":   prot.selectLockedCells,
                "selectUnlockedCells": prot.selectUnlockedCells,
                "formatCells":         prot.formatCells,
                "formatColumns":       prot.formatColumns,
                "formatRows":          prot.formatRows,
                "insertColumns":       prot.insertColumns,
                "insertRows":          prot.insertRows,
                "deleteColumns":       prot.deleteColumns,
                "deleteRows":          prot.deleteRows,
                "sort":                prot.sort,
                "autoFilter":          prot.autoFilter,
                "objects":             prot.objects,
                "scenarios":           prot.scenarios,
                "insertHyperlinks":    prot.insertHyperlinks,
                "pivotTables":         prot.pivotTables,
            }

        # Print titles
        print_titles = None
        ptr = ws.print_title_rows
        ptc = ws.print_title_cols
        if ptr or ptc:
            print_titles = {"rows": ptr, "cols": ptc}

        # Header / footer
        hf_data = {}
        try:
            oh = ws.oddHeader
            if oh:
                if oh.left   and oh.left.text:   hf_data["hl"] = oh.left.text
                if oh.center and oh.center.text: hf_data["hc"] = oh.center.text
                if oh.right  and oh.right.text:  hf_data["hr"] = oh.right.text
            of_ = ws.oddFooter
            if of_:
                if of_.left   and of_.left.text:   hf_data["fl"] = of_.left.text
                if of_.center and of_.center.text: hf_data["fc"] = of_.center.text
                if of_.right  and of_.right.text:  hf_data["fr"] = of_.right.text
        except Exception:
            pass

        # Hyperlinks (external target and/or internal location like "Sheet2!A1")
        hyperlinks = {}
        for _row in ws.iter_rows():
            for _cell in _row:
                if _cell.hyperlink:
                    hl = _cell.hyperlink
                    target = getattr(hl, "target", None) or None
                    location = getattr(hl, "location", None) or None
                    if target or location:
                        hyperlinks[_cell.coordinate] = {
                            "target":   target,
                            "location": location,
                            "tooltip":  getattr(hl, "tooltip", None),
                        }

        # Comments
        comments = {}
        for _row in ws.iter_rows():
            for _cell in _row:
                if _cell.comment:
                    comments[_cell.coordinate] = {
                        "text":   _cell.comment.text or "",
                        "author": _cell.comment.author or "",
                    }

        # Tables
        tables = []
        try:
            for _t in ws.tables.values():
                ts = _t.tableStyleInfo
                tables.append({
                    "name": _t.displayName,
                    "ref":  _t.ref,
                    "style": {
                        "name":            ts.name            if ts else None,
                        "showRowStripes":  ts.showRowStripes  if ts else None,
                        "showColStripes":  ts.showColumnStripes if ts else None,
                        "showFirstCol":    ts.showFirstColumn  if ts else None,
                        "showLastCol":     ts.showLastColumn   if ts else None,
                    } if ts else None,
                })
        except Exception:
            pass

        sheets.append({
            "name":          sname,
            "cw":            col_widths,
            "ch":            col_hidden,
            "co":            col_outline or None,
            "rows":          rows,
            "freeze":        ws.freeze_panes,
            "validations":   _ser_validations(ws),
            "sheet_view":    _serialize_sheet_view(sv, raw_sheet_views.get(sname)),
            "tab_color":     tc,
            "auto_filter":   str(ws.auto_filter.ref) if ws.auto_filter.ref else None,
            "page_setup":    page_setup or None,
            "page_margins":  page_margins or None,
            "protection":    protection,
            "print_titles":  print_titles,
            "header_footer": hf_data or None,
            "hyperlinks":    hyperlinks or None,
            "comments":      comments or None,
            "tables":        tables or None,
        })
        raw_sheet_xml = raw_sheet_formats.get(sname) or {}
        if raw_sheet_xml.get("root_attrs") or raw_sheet_xml.get("sheetFormatPr"):
            sheets[-1]["_sheet_format_pr"] = {
                "root_attrs": raw_sheet_xml.get("root_attrs", ""),
                "sheetFormatPr": raw_sheet_xml.get("sheetFormatPr"),
            }
        if raw_sheet_xml.get("cols"):
            sheets[-1]["_cols_raw"] = {
                "xml": raw_sheet_xml["cols"],
                "state": col_state,
            }
        if raw_data_validations.get(sname):
            sheets[-1]["data_validations_xml"] = raw_data_validations[sname]

    # Named ranges (workbook level)
    named_ranges = []
    try:
        for name in wb.defined_names:
            dn = wb.defined_names[name]
            named_ranges.append({
                "name":     dn.name,
                "value":    dn.attr_text,
                "sheet_id": dn.localSheetId,
            })
    except Exception:
        pass

    # Extract conditional formatting as raw XML for passthrough
    cf_xml = _extract_cf_xml(path, [sd["name"] for sd in sheets])
    dxfs_xml = cf_xml.pop("__dxfs__", None)
    for sd in sheets:
        if sd["name"] in cf_xml:
            sd["cf_xml"] = cf_xml[sd["name"]]

    # Extract drawing/chart/image data for passthrough
    drawing_data = _extract_drawing_data(path, _sfm)
    shape_inventory = _extract_shape_inventory(drawing_data)
    for sd in sheets:
        if sd["name"] in drawing_data:
            sd["drawing_data"] = drawing_data[sd["name"]]
        if sd["name"] in shape_inventory:
            sd["shapes"] = shape_inventory[sd["name"]]

    # Document properties (docProps/core.xml)
    doc_props = {}
    try:
        props = wb.properties
        for attr in ("creator", "title", "subject", "description",
                     "keywords", "category", "lastModifiedBy"):
            value = getattr(props, attr, None)
            if value:
                doc_props[attr] = value
        if props.created:
            doc_props["created"] = props.created.isoformat()
    except Exception:
        pass

    # Workbook view (active tab, window geometry)
    wb_view = {}
    try:
        view = wb.views[0]
        for attr in getattr(type(view), "__attrs__", ()):
            value = getattr(view, attr, None)
            if value is not None:
                wb_view[attr] = value
    except Exception:
        pass

    # Release the zip handle now — openpyxl workbooks have reference cycles,
    # so waiting for GC can leave the file locked on Windows.
    try:
        wb.close()
    except Exception:
        pass

    return {"source": str(path), "sheets": sheets, "named_ranges": named_ranges,
            "dxfs_xml": dxfs_xml, "theme_xml": theme_xml,
            "doc_props": doc_props or None, "wb_view": wb_view or None}


# ── Reconstruct ───────────────────────────────────────────────────────────────

def reconstruct_excel(data: dict, output_path: str) -> list[str]:
    """Reconstruct an Excel file from metadata dict produced by serialize_excel.

    Writes atomically: everything is assembled in a temp file that replaces
    output_path only on success, so a failure never corrupts an existing file.
    Returns warning strings for passthrough features that could not be restored.
    """
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side, Protection
    from openpyxl.worksheet.datavalidation import DataValidation

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    if data.get("theme_xml"):
        try:
            import base64
            wb.loaded_theme = base64.b64decode(data["theme_xml"])
        except Exception:
            pass

    for sd in data["sheets"]:
        ws = wb.create_sheet(sd["name"])

        if sd.get("freeze"):
            ws.freeze_panes = sd["freeze"]

        sv_data = sd.get("sheet_view") or {}
        sv_aliases = {"zoom": "zoomScale"}
        sv_supported = set(getattr(type(ws.sheet_view), "__attrs__", ()))
        for key, value in sv_data.items():
            if key.startswith("_") or value is None:
                continue
            attr = sv_aliases.get(key, key)
            if attr not in sv_supported:
                continue
            try:
                setattr(ws.sheet_view, attr, value)
            except Exception:
                pass

        for col_letter, width in sd["cw"].items():
            if width:
                ws.column_dimensions[col_letter].width = width
        for col_letter in sd.get("ch", {}):
            ws.column_dimensions[col_letter].hidden = True
        for col_letter, level in (sd.get("co") or {}).items():
            ws.column_dimensions[col_letter].outlineLevel = level

        if sd.get("tab_color"):
            from openpyxl.styles.colors import Color as _Color
            ws.sheet_properties.tabColor = _Color(rgb=sd["tab_color"])

        if sd.get("auto_filter"):
            ws.auto_filter.ref = sd["auto_filter"]

        ps_data = sd.get("page_setup") or {}
        if ps_data:
            for key in ("orientation", "paperSize", "fitToWidth", "fitToHeight", "scale"):
                if ps_data.get(key) is not None:
                    setattr(ws.page_setup, key, ps_data[key])
            if ps_data.get("fitToPage"):
                try:
                    from openpyxl.worksheet.properties import PageSetupProperties
                    if ws.sheet_properties.pageSetUpPr is None:
                        ws.sheet_properties.pageSetUpPr = PageSetupProperties()
                    ws.sheet_properties.pageSetUpPr.fitToPage = True
                except Exception:
                    pass

        pm_data = sd.get("page_margins") or {}
        if pm_data:
            for key in ("left", "right", "top", "bottom", "header", "footer"):
                if pm_data.get(key) is not None:
                    setattr(ws.page_margins, key, pm_data[key])

        prot_data = sd.get("protection")
        if prot_data:
            ws.protection.sheet = True
            if prot_data.get("password"):
                try:
                    ws.protection.set_password(prot_data["password"], already_hashed=True)
                except Exception:
                    ws.protection.password = prot_data["password"]
            for key in ("selectLockedCells", "selectUnlockedCells", "formatCells",
                        "formatColumns", "formatRows", "insertColumns", "insertRows",
                        "deleteColumns", "deleteRows", "sort", "autoFilter",
                        "objects", "scenarios", "insertHyperlinks", "pivotTables"):
                if prot_data.get(key) is not None:
                    setattr(ws.protection, key, prot_data[key])

        # Print titles
        pt = sd.get("print_titles") or {}
        if pt.get("rows"):
            ws.print_title_rows = pt["rows"]
        if pt.get("cols"):
            ws.print_title_cols = pt["cols"]

        # Header / footer
        hf = sd.get("header_footer") or {}
        if hf.get("hl"): ws.oddHeader.left.text   = hf["hl"]
        if hf.get("hc"): ws.oddHeader.center.text = hf["hc"]
        if hf.get("hr"): ws.oddHeader.right.text  = hf["hr"]
        if hf.get("fl"): ws.oddFooter.left.text   = hf["fl"]
        if hf.get("fc"): ws.oddFooter.center.text = hf["fc"]
        if hf.get("fr"): ws.oddFooter.right.text  = hf["fr"]

        # Tables
        try:
            from openpyxl.worksheet.table import Table, TableStyleInfo as TSI
            for t_data in (sd.get("tables") or []):
                t = Table(displayName=t_data["name"], ref=t_data["ref"])
                s = t_data.get("style")
                if s:
                    t.tableStyleInfo = TSI(
                        name=s.get("name"),
                        showFirstColumn=bool(s.get("showFirstCol")),
                        showLastColumn=bool(s.get("showLastCol")),
                        showRowStripes=bool(s.get("showRowStripes", True)),
                        showColumnStripes=bool(s.get("showColStripes", False)),
                    )
                ws.add_table(t)
        except Exception:
            pass

        for r_idx, row_data in enumerate(sd["rows"], 1):
            if row_data.get("h"):
                ws.row_dimensions[r_idx].height = row_data["h"]
            if row_data.get("hidden"):
                ws.row_dimensions[r_idx].hidden = True
            if row_data.get("outline"):
                ws.row_dimensions[r_idx].outlineLevel = row_data["outline"]

            for c_idx, cd in enumerate(row_data["cells"], 1):
                if cd["merge"] == "slave":
                    continue

                cell = ws.cell(row=r_idx, column=c_idx, value=cd["v"])
                if cd.get("dt") == "s" and isinstance(cd["v"], str):
                    cell.data_type = "s"  # literal text, not a formula

                raw_fill = _usable_raw_fill(cd)
                if raw_fill:
                    try:
                        cell.fill = _make_pattern_fill_from_raw(raw_fill)
                    except Exception:
                        pass
                elif cd["fill"]:
                    try:
                        cell.fill = PatternFill("solid", fgColor=cd["fill"])
                    except Exception:
                        pass

                fk: dict = {}
                if cd["bold"]:              fk["bold"]      = True
                if cd["italic"]:            fk["italic"]    = True
                if cd.get("size"):          fk["size"]      = cd["size"]
                if cd.get("font"):          fk["name"]      = cd["font"]
                if cd.get("uline"):
                    fk["underline"] = "single" if cd["uline"] is True else cd["uline"]
                if cd.get("strike"):        fk["strike"]    = True
                if cd.get("vAlign"):        fk["vertAlign"] = cd["vAlign"]
                _apply_raw_font_kwargs(fk, cd)
                if fk:
                    cell.font = Font(**fk)

                ak: dict = {}
                if cd.get("wrap"):   ak["wrap_text"]    = True
                if cd.get("halign"): ak["horizontal"]   = cd["halign"]
                if cd.get("valign"): ak["vertical"]     = cd["valign"]
                if cd.get("rot"):    ak["text_rotation"] = cd["rot"]
                if cd.get("indent"): ak["indent"]        = cd["indent"]
                if cd.get("shrink"): ak["shrink_to_fit"] = True
                if ak:
                    cell.alignment = Alignment(**ak)

                if cd.get("numfmt"):
                    cell.number_format = cd["numfmt"]

                bdr = cd.get("border", {})
                if bdr:
                    cell.border = Border(
                        top=_make_border_side(bdr.get("top")),
                        bottom=_make_border_side(bdr.get("bottom")),
                        left=_make_border_side(bdr.get("left")),
                        right=_make_border_side(bdr.get("right")),
                        diagonal=_make_border_side(bdr.get("diagonal")),
                        diagonalUp=bool(bdr.get("diagonalUp")),
                        diagonalDown=bool(bdr.get("diagonalDown")),
                    )

                locked = cd.get("locked", True)
                hidden_p = cd.get("hidden_cell", False)
                if locked is False or hidden_p:
                    cell.protection = Protection(locked=bool(locked), hidden=bool(hidden_p))

                if cd.get("qp"):
                    try:
                        from openpyxl.styles.cell_style import StyleArray
                        if cell._style is None:
                            cell._style = StyleArray()
                        cell._style.quotePrefix = 1
                    except Exception:
                        pass

                mi = cd["merge"]
                if isinstance(mi, dict) and (mi.get("rowspan", 1) > 1 or mi.get("colspan", 1) > 1):
                    ws.merge_cells(
                        start_row=r_idx, start_column=c_idx,
                        end_row=r_idx + mi["rowspan"] - 1,
                        end_column=c_idx + mi["colspan"] - 1,
                    )

        for vd in sd.get("validations", []):
            dv_kwargs = {
                key: vd.get(key)
                for key in (
                    "type", "formula1", "formula2", "showErrorMessage",
                    "showInputMessage", "showDropDown",
                    "promptTitle", "errorStyle", "error", "prompt",
                    "errorTitle", "imeMode", "operator",
                )
                if vd.get(key) is not None
            }
            if vd.get("allowBlank") is not None:
                dv_kwargs["allow_blank"] = vd["allowBlank"]
            elif vd.get("allow_blank") is not None:
                dv_kwargs["allow_blank"] = vd["allow_blank"]
            dv = DataValidation(**dv_kwargs)
            for sqref_part in vd["sqref"].split():
                dv.add(sqref_part)
            ws.add_data_validation(dv)

        # Hyperlinks
        try:
            from openpyxl.worksheet.hyperlink import Hyperlink
            for coord, hl_data in (sd.get("hyperlinks") or {}).items():
                if not (hl_data.get("target") or hl_data.get("location")):
                    continue
                cell = ws[coord]
                hl = Hyperlink(ref=coord, target=hl_data.get("target"))
                if hl_data.get("location"):
                    hl.location = hl_data["location"]
                if hl_data.get("tooltip"):
                    hl.tooltip = hl_data["tooltip"]
                cell.hyperlink = hl
        except Exception:
            pass

        # Comments
        try:
            from openpyxl.comments import Comment
            for coord, cm in (sd.get("comments") or {}).items():
                ws[coord].comment = Comment(cm["text"], cm.get("author", ""))
        except Exception:
            pass

    # Named ranges
    try:
        from openpyxl.workbook.defined_name import DefinedName
        for nr in data.get("named_ranges") or []:
            dn = DefinedName(nr["name"], attr_text=nr["value"])
            if nr.get("sheet_id") is not None:
                dn.localSheetId = nr["sheet_id"]
            wb.defined_names[nr["name"]] = dn
    except Exception:
        pass

    # Document properties
    dp = data.get("doc_props") or {}
    if dp:
        try:
            from datetime import datetime
            for key, value in dp.items():
                if key == "created":
                    wb.properties.created = datetime.fromisoformat(value)
                else:
                    setattr(wb.properties, key, value)
        except Exception:
            pass

    # Workbook view (active tab, window geometry)
    wv = data.get("wb_view") or {}
    if wv:
        try:
            view = wb.views[0]
            supported = set(getattr(type(view), "__attrs__", ()))
            for key, value in wv.items():
                if key not in supported:
                    continue
                if key == "activeTab":
                    value = max(0, min(int(value), len(wb.worksheets) - 1))
                try:
                    setattr(view, key, value)
                except Exception:
                    pass
            if wv.get("activeTab") is not None:
                # openpyxl's writer takes activeTab from wb.active, not the view
                wb.active = max(0, min(int(wv["activeTab"]), len(wb.worksheets) - 1))
        except Exception:
            pass

    # Atomic write: assemble everything in a temp file, replace target on success.
    import os as _os
    out_str = str(output_path)
    tmp_out = out_str + ".~saving.tmp"
    warnings: list[str] = []
    try:
        wb.save(tmp_out)

        # Restore raw theme/indexed/tint fill XML and sheetView attributes after save.
        for w in (
            _inject_raw_fills(tmp_out, data),
            _inject_sheet_view_attrs(tmp_out, data),
            _inject_sheet_format_pr(tmp_out, data),
            _inject_raw_cols(tmp_out, data),
            _inject_data_validations_xml(
                tmp_out,
                {
                    sd["name"]: sd["data_validations_xml"]
                    for sd in data["sheets"]
                    if sd.get("data_validations_xml")
                },
            ),
        ):
            if w:
                warnings.append(w)

        # Inject conditional formatting (XML passthrough, must be after save)
        cf_map = {sd["name"]: sd["cf_xml"] for sd in data["sheets"] if sd.get("cf_xml")}
        dxfs_xml = data.get("dxfs_xml")
        if cf_map or dxfs_xml:
            if dxfs_xml:
                cf_map["__dxfs__"] = dxfs_xml
            w = _inject_cf_xml(tmp_out, cf_map)
            if w:
                warnings.append(w)

        # Inject drawings/charts/images (XML passthrough, must be after save)
        drawing_sheets = {sd["name"]: sd["drawing_data"] for sd in data["sheets"] if sd.get("drawing_data")}
        if drawing_sheets:
            import zipfile as _zf2
            try:
                with _zf2.ZipFile(tmp_out, "r") as _z2:
                    _wb2   = _z2.read("xl/workbook.xml").decode("utf-8")
                    _rels2 = _z2.read("xl/_rels/workbook.xml.rels").decode("utf-8")
                new_sfm = _xlsx_sheet_file_map(_wb2, _rels2)
            except Exception:
                new_sfm = {}
            w = _inject_drawing_data(tmp_out, drawing_sheets, new_sfm)
            if w:
                warnings.append(w)

        w = _restore_missing_package_parts(data.get("source"), tmp_out)
        if w:
            warnings.append(w)

        warnings.extend(validate_xlsx(tmp_out))

        # Windows: a stale GC-held handle or AV scan can briefly lock the
        # target — retry the swap a few times before giving up.
        import gc as _gc
        import time as _time
        for attempt in range(5):
            try:
                _os.replace(tmp_out, out_str)
                break
            except PermissionError:
                if attempt == 4:
                    raise
                _gc.collect()
                _time.sleep(0.2)
    except Exception:
        if _os.path.exists(tmp_out):
            try:
                _os.remove(tmp_out)
            except OSError:
                pass
        raise
    return warnings

