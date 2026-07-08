"""
excel-tools — MCP server for round-trip Excel editing.
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

import copy
import json
import re
import xml.etree.ElementTree as ET

from mcp.server.fastmcp import FastMCP
from mcp.types import TextContent

from core import (
    diff_xlsx_package,
    inspect_xlsx_package,
    reconstruct_excel,
    serialize_excel,
    uri_to_path,
)

# Session cache lives at module level so it persists across tool calls
_sessions: dict[str, dict] = {}

# Style keys accepted by excel_set_style
_STYLE_KEYS = frozenset({"fill", "fcolor", "bold", "italic", "strike", "underline", "uline", "size", "font", "wrap", "halign", "valign", "numfmt"})

# Blank cell template used when inserting empty cells
_EMPTY_CELL: dict = {
    "v": None, "fill": None, "bold": False, "italic": False,
    "size": None, "font": None, "fcolor": None, "wrap": False,
    "halign": None, "valign": None, "numfmt": "General", "merge": {}, "border": {},
}


def _normalize_underline(value):
    if value is True:
        return "single"
    if value in (False, "", None):
        return None
    return value


def _apply_style(cell: dict, style: dict) -> None:
    for key, value in style.items():
        if key == "underline":
            key = "uline"
        if key in _STYLE_KEYS:
            if key == "uline":
                value = _normalize_underline(value)
            cell[key] = value
            if key == "fill":
                cell.pop("_fill_raw", None)
            if key in {"fcolor", "bold", "italic", "strike", "underline", "uline", "size", "font"}:
                cell.pop("_font_raw", None)


def _drop_raw_fills(sheet: dict) -> None:
    for row in sheet.get("rows", []):
        for cell in row.get("cells", []):
            cell.pop("_fill_raw", None)
            cell.pop("_font_raw", None)
            for side in (cell.get("border") or {}).values():
                if isinstance(side, dict):
                    side.pop("_color_raw", None)


_LEGACY_OR_BINARY_EXTS = {".xls", ".xlsb"}


def _check_supported(path) -> None:
    ext = Path(str(path)).suffix.lower()
    if ext in _LEGACY_OR_BINARY_EXTS:
        raise ValueError(
            f"'{ext}' files are not supported by the edit engine. Use convert_to_markdown "
            "for read-only extraction, or convert to an OOXML workbook (.xlsx/.xlsm/.xltx/.xltm).")


def _resolve_session_key(session_key: str) -> str:
    """Return the canonical key for a session, tolerating path-case differences."""
    if session_key in _sessions:
        return session_key
    try:
        alt = str(Path(session_key).resolve())
    except (OSError, ValueError):
        alt = None
    if alt and alt in _sessions:
        return alt
    raise ValueError(f"Session '{session_key}' not found. Call excel_load first.")


def _get_session(session_key: str) -> dict:
    return _sessions[_resolve_session_key(session_key)]


def _find_sheet(data: dict, name: str) -> dict:
    sheet = next((s for s in data["sheets"] if s["name"] == name), None)
    if sheet is None:
        available = [s["name"] for s in data["sheets"]]
        raise ValueError(f"Sheet '{name}' not found. Available: {available}")
    return sheet



def _excel_family_ext(path: str | Path) -> str:
    return Path(path).suffix.lower()


def _check_save_extension_compatible(source: str, dest: str) -> None:
    src_ext = _excel_family_ext(source)
    dst_ext = _excel_family_ext(dest)
    macro_exts = {".xlsm", ".xltm"}
    non_macro_exts = {".xlsx", ".xltx"}
    if src_ext in macro_exts and dst_ext in non_macro_exts:
        raise ValueError(f"Refusing to save macro-enabled {src_ext} workbook as {dst_ext}; choose a macro-enabled extension.")
    if src_ext in non_macro_exts and dst_ext in macro_exts:
        raise ValueError(f"Refusing to save non-macro {src_ext} workbook as {dst_ext}; choose a non-macro extension.")

def _excel_range_to_indices(range_ref: str) -> tuple[int, int, int, int]:
    """Convert an Excel A1 range to 0-based inclusive row/column bounds."""
    import openpyxl.utils

    normalized = range_ref.split("!", 1)[-1].replace("$", "")
    if ":" not in normalized:
        normalized = f"{normalized}:{normalized}"
    min_col, min_row, max_col, max_row = openpyxl.utils.range_boundaries(normalized)
    return min_row - 1, max_row - 1, min_col - 1, max_col - 1

def _range_from_args(
    range_ref: str | None,
    start_row: int | None,
    end_row: int | None,
    start_col: int | None,
    end_col: int | None,
    max_row: int,
    max_col: int,
) -> tuple[int, int, int, int]:
    if range_ref:
        return _excel_range_to_indices(range_ref)
    r1 = 0 if start_row is None else start_row
    r2 = max_row - 1 if end_row is None else end_row - 1
    c1 = 0 if start_col is None else start_col
    c2 = max_col - 1 if end_col is None else end_col - 1
    if min(r1, r2, c1, c2) < 0 or r1 > r2 or c1 > c2:
        raise ValueError("Invalid range bounds. Use 0-based start and exclusive end indexes.")
    return r1, r2, c1, c2

def _limit_workbook_data(data: dict, max_rows: int | None = None, max_cols: int | None = None) -> dict:
    """Return a copy of serialized workbook data trimmed for Markdown export."""
    if max_rows is None and max_cols is None:
        return data
    limited = copy.deepcopy(data)
    for sheet in limited.get("sheets", []):
        rows = sheet.get("rows", [])
        if max_rows is not None:
            sheet["rows"] = rows[:max_rows]
        if max_cols is not None:
            for row in sheet.get("rows", []):
                row["cells"] = row.get("cells", [])[:max_cols]
    return limited


# ── Structural-shift machinery ────────────────────────────────────────────────
# When rows/columns are inserted or deleted, every piece of coordinate-anchored
# metadata must follow: merges, hyperlinks, comments, validations, conditional
# formatting, tables, auto filter, freeze panes, print titles, drawing anchors.

def _shift_col_dims(sheet: dict, at_col: int, delta: int) -> None:
    """Shift column width/hidden/outline maps after a column insert (+N) / delete (-1)."""
    import openpyxl.utils

    def shift(d: dict) -> dict:
        result = {}
        for letter, value in (d or {}).items():
            idx = openpyxl.utils.column_index_from_string(letter) - 1  # 0-based
            if delta > 0:
                new_idx = idx + delta if idx >= at_col else idx
            else:
                if idx == at_col:
                    continue  # drop the deleted column's entry
                new_idx = idx + delta if idx > at_col else idx
            result[openpyxl.utils.get_column_letter(new_idx + 1)] = value
        return result

    sheet["cw"] = shift(sheet.get("cw") or {})
    if sheet.get("ch"):
        sheet["ch"] = shift(sheet["ch"])
    if sheet.get("co"):
        sheet["co"] = shift(sheet["co"])


def _insert_maps(pos: int, count: int):
    """(strict, clamp) index maps for inserting `count` slots at `pos`."""
    def shift(i):
        return i + count if i >= pos else i
    return shift, shift


def _delete_maps(deleted: set):
    """(strict, clamp) index maps for deleting an index set.

    strict: deleted index → None; clamp: deleted index → position of the
    element that takes its place (for anchors that must survive).
    """
    import bisect
    sd = sorted(deleted)

    def strict(i):
        if i in deleted:
            return None
        return i - bisect.bisect_left(sd, i)

    def clamp(i):
        return i - bisect.bisect_left(sd, i)

    return strict, clamp


_COORD_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")
_RANGE_RE = re.compile(
    r"^(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?::(\$?)([A-Za-z]{1,3})(\$?)(\d+))?$")
_ROWS_RE = re.compile(r"^(\$?)(\d+):(\$?)(\d+)$")
_COLS_RE = re.compile(r"^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})$")


def _coord_to_rc(coord: str):
    import openpyxl.utils
    m = _COORD_RE.match(str(coord))
    if not m:
        return None
    return int(m.group(2)) - 1, openpyxl.utils.column_index_from_string(m.group(1).upper()) - 1


def _rc_to_coord(r: int, c: int) -> str:
    import openpyxl.utils
    return f"{openpyxl.utils.get_column_letter(c + 1)}{r + 1}"


def _shift_span(lo: int, hi: int, smap, cmap):
    """Map an inclusive index span; None if nothing survives a deletion."""
    a, b = smap(lo), smap(hi)
    if a is not None and b is not None:
        return a, b
    if hi - lo <= 50000:
        vals = [v for v in (smap(i) for i in range(lo, hi + 1)) if v is not None]
        if not vals:
            return None
        return min(vals), max(vals)
    a, b = cmap(lo), cmap(hi)
    return (a, b) if a <= b else None


def _shift_ref(ref: str, row_maps=None, col_maps=None):
    """Shift one A1-style ref/range ($ preserved). None = entirely removed."""
    import openpyxl.utils as U
    ref = str(ref)
    m = _RANGE_RE.match(ref)
    if m:
        d1, col1, d2, row1, d3, col2, d4, row2 = m.groups()
        c1 = U.column_index_from_string(col1.upper()) - 1
        r1 = int(row1) - 1
        c2 = U.column_index_from_string((col2 or col1).upper()) - 1
        r2 = int(row2 or row1) - 1
        if row_maps:
            span = _shift_span(min(r1, r2), max(r1, r2), *row_maps)
            if span is None:
                return None
            r1, r2 = span
        if col_maps:
            span = _shift_span(min(c1, c2), max(c1, c2), *col_maps)
            if span is None:
                return None
            c1, c2 = span
        first = f"{d1}{U.get_column_letter(c1 + 1)}{d2}{r1 + 1}"
        if m.group(6) is None:
            return first
        second = f"{d3}{U.get_column_letter(c2 + 1)}{d4}{r2 + 1}"
        return f"{first}:{second}"
    m = _ROWS_RE.match(ref)
    if m:
        if not row_maps:
            return ref
        d1, row1, d2, row2 = m.groups()
        span = _shift_span(int(row1) - 1, int(row2) - 1, *row_maps)
        if span is None:
            return None
        return f"{d1}{span[0] + 1}:{d2}{span[1] + 1}"
    m = _COLS_RE.match(ref)
    if m:
        if not col_maps:
            return ref
        d1, col1, d2, col2 = m.groups()
        span = _shift_span(U.column_index_from_string(col1.upper()) - 1,
                           U.column_index_from_string(col2.upper()) - 1, *col_maps)
        if span is None:
            return None
        return f"{d1}{U.get_column_letter(span[0] + 1)}:{d2}{U.get_column_letter(span[1] + 1)}"
    return ref


def _shift_sqref(sqref: str, row_maps=None, col_maps=None):
    parts = []
    for token in str(sqref).split():
        s = _shift_ref(token, row_maps, col_maps)
        if s:
            parts.append(s)
    return " ".join(parts) or None


def _shift_coord_map(d, row_maps=None, col_maps=None):
    if not d:
        return d
    out = {}
    for coord, value in d.items():
        rc = _coord_to_rc(coord)
        if rc is None:
            out[coord] = value
            continue
        r, c = rc
        if row_maps:
            r = row_maps[0](r)
        if col_maps:
            c = col_maps[0](c) if c is not None else None
        if r is None or c is None:
            continue
        out[_rc_to_coord(r, c)] = value
    return out or None


def _shift_dv_xml(xml: str, row_maps=None, col_maps=None):
    def repl(m):
        block = m.group(0)
        sq = re.search(r'sqref="([^"]+)"', block)
        if not sq:
            return block
        new = _shift_sqref(sq.group(1), row_maps, col_maps)
        if new is None:
            return ""
        return block.replace(sq.group(0), f'sqref="{new}"', 1)

    xml = re.sub(r"<dataValidation\b[^>]*/>|<dataValidation\b[^>]*>.*?</dataValidation>",
                 repl, xml, flags=re.DOTALL)
    n = len(re.findall(r"<dataValidation\b", xml)) - len(re.findall(r"<dataValidations\b", xml))
    if n <= 0:
        return None
    return re.sub(r'(<dataValidations\b[^>]*?count=")\d+(")', rf"\g<1>{n}\g<2>", xml, count=1)


def _shift_cf_blocks(blocks: list, row_maps=None, col_maps=None):
    out = []
    for block in blocks:
        m = re.search(r'sqref="([^"]+)"', block)
        if not m:
            out.append(block)
            continue
        new = _shift_sqref(m.group(1), row_maps, col_maps)
        if new is None:
            continue
        out.append(block.replace(m.group(0), f'sqref="{new}"', 1))
    return out or None


def _shift_drawing_anchors(xml: str, row_maps=None, col_maps=None) -> str:
    if row_maps:
        clamp = row_maps[1]
        xml = re.sub(r"(<\w+:row>)(\d+)(</\w+:row>)",
                     lambda m: f"{m.group(1)}{clamp(int(m.group(2)))}{m.group(3)}", xml)
    if col_maps:
        clamp = col_maps[1]
        xml = re.sub(r"(<\w+:col>)(\d+)(</\w+:col>)",
                     lambda m: f"{m.group(1)}{clamp(int(m.group(2)))}{m.group(3)}", xml)
    return xml


def _shift_sheet_meta(sheet: dict, row_maps=None, col_maps=None) -> None:
    """Shift all coordinate-anchored sheet metadata after a structural change."""
    if sheet.get("hyperlinks"):
        sheet["hyperlinks"] = _shift_coord_map(sheet["hyperlinks"], row_maps, col_maps)
    if sheet.get("comments"):
        sheet["comments"] = _shift_coord_map(sheet["comments"], row_maps, col_maps)

    kept = []
    for vd in sheet.get("validations") or []:
        new_sq = _shift_sqref(vd.get("sqref", ""), row_maps, col_maps)
        if new_sq is None:
            continue
        vd["sqref"] = new_sq
        kept.append(vd)
    if sheet.get("validations") is not None:
        sheet["validations"] = kept

    if sheet.get("data_validations_xml"):
        new_xml = _shift_dv_xml(sheet["data_validations_xml"], row_maps, col_maps)
        if new_xml is None:
            sheet.pop("data_validations_xml", None)
        else:
            sheet["data_validations_xml"] = new_xml

    if sheet.get("cf_xml"):
        new_blocks = _shift_cf_blocks(sheet["cf_xml"], row_maps, col_maps)
        if new_blocks is None:
            sheet.pop("cf_xml", None)
        else:
            sheet["cf_xml"] = new_blocks

    if sheet.get("auto_filter"):
        sheet["auto_filter"] = _shift_ref(sheet["auto_filter"], row_maps, col_maps)

    if sheet.get("tables"):
        kept_tables = []
        for t in sheet["tables"]:
            new_ref = _shift_ref(t.get("ref", ""), row_maps, col_maps)
            if new_ref is None:
                continue
            t["ref"] = new_ref
            kept_tables.append(t)
        sheet["tables"] = kept_tables or None

    if sheet.get("freeze"):
        rc = _coord_to_rc(str(sheet["freeze"]))
        if rc:
            r, c = rc
            if row_maps:
                r = row_maps[1](r)
            if col_maps:
                c = col_maps[1](c)
            sheet["freeze"] = _rc_to_coord(r, c) if (r > 0 or c > 0) else None

    pt = sheet.get("print_titles")
    if pt:
        if pt.get("rows") and row_maps:
            pt["rows"] = _shift_ref(pt["rows"], row_maps, None)
        if pt.get("cols") and col_maps:
            pt["cols"] = _shift_ref(pt["cols"], None, col_maps)

    dd = sheet.get("drawing_data")
    if dd and dd.get("drawing_xml"):
        dd["drawing_xml"] = _shift_drawing_anchors(dd["drawing_xml"], row_maps, col_maps)


def _capture_merge_regions(rows: list) -> list:
    """Merge regions [(r1,c1,r2,c2)] from actual grid positions + spans."""
    regions = []
    for r, row in enumerate(rows):
        for c, cd in enumerate(row.get("cells") or []):
            mi = cd.get("merge")
            if isinstance(mi, dict) and (mi.get("rowspan", 1) > 1 or mi.get("colspan", 1) > 1):
                regions.append([r, c,
                                r + mi.get("rowspan", 1) - 1,
                                c + mi.get("colspan", 1) - 1])
    return regions


def _restamp_merges(rows: list, regions: list) -> None:
    """Clear all merge markers and re-stamp them from a region list."""
    for row in rows:
        for cd in row.get("cells") or []:
            cd["merge"] = {}
    n_rows = len(rows)
    for r1, c1, r2, c2 in regions:
        if not (0 <= r1 < n_rows):
            continue
        r2 = min(r2, n_rows - 1)
        if r2 == r1 and c2 == c1:
            continue  # collapsed to a single cell — no merge left
        cells = rows[r1].get("cells") or []
        if c1 >= len(cells):
            continue
        cells[c1]["merge"] = {"r1": r1, "c1": c1, "r2": r2, "c2": c2,
                              "rowspan": r2 - r1 + 1, "colspan": c2 - c1 + 1}
        for r in range(r1, r2 + 1):
            row_cells = rows[r].get("cells") or []
            for c in range(c1, min(c2 + 1, len(row_cells))):
                if r == r1 and c == c1:
                    continue
                row_cells[c]["merge"] = "slave"


def _regions_after_insert(regions: list, pos: int, count: int, axis: str) -> list:
    out = []
    for r1, c1, r2, c2 in regions:
        lo, hi = (r1, r2) if axis == "row" else (c1, c2)
        if pos <= lo:
            lo, hi = lo + count, hi + count
        elif pos <= hi:
            hi += count  # inserted inside the merge → region extends
        out.append([lo, c1, hi, c2] if axis == "row" else [r1, lo, r2, hi])
    return out


def _regions_after_delete(regions: list, maps, axis: str) -> list:
    smap, cmap = maps
    out = []
    for r1, c1, r2, c2 in regions:
        lo, hi = (r1, r2) if axis == "row" else (c1, c2)
        span = _shift_span(lo, hi, smap, cmap)
        if span is None:
            continue
        lo, hi = span
        out.append([lo, c1, hi, c2] if axis == "row" else [r1, lo, r2, hi])
    return out


def _block_merge_regions(block_rows: list, row_offset: int) -> list:
    """Self-contained merges inside an inserted block, clamped to the block."""
    out = []
    n = len(block_rows)
    for r1, c1, r2, c2 in _capture_merge_regions(block_rows):
        out.append([r1 + row_offset, c1, min(r2, n - 1) + row_offset, c2])
    return out


def _apply_row_insert(data: dict, sheet: dict, pos: int, new_rows: list) -> None:
    regions = _capture_merge_regions(sheet["rows"])
    block_regions = _block_merge_regions(new_rows, pos)
    sheet["rows"] = sheet["rows"][:pos] + new_rows + sheet["rows"][pos:]
    regions = _regions_after_insert(regions, pos, len(new_rows), "row") + block_regions
    _restamp_merges(sheet["rows"], regions)
    maps = _insert_maps(pos, len(new_rows))
    _shift_sheet_meta(sheet, row_maps=maps)
    _shift_formulas_workbook(data, sheet["name"], row_maps=maps)


def _apply_row_delete(data: dict, sheet: dict, to_delete: set) -> None:
    regions = _capture_merge_regions(sheet["rows"])
    maps = _delete_maps(set(to_delete))
    sheet["rows"] = [row for i, row in enumerate(sheet["rows"]) if i not in to_delete]
    _restamp_merges(sheet["rows"], _regions_after_delete(regions, maps, "row"))
    _shift_sheet_meta(sheet, row_maps=maps)
    _shift_formulas_workbook(data, sheet["name"], row_maps=maps)


def _finish_col_insert(data: dict, sheet: dict, regions: list, pos: int, count: int = 1) -> None:
    _restamp_merges(sheet["rows"], _regions_after_insert(regions, pos, count, "col"))
    _shift_col_dims(sheet, pos, +count)
    maps = _insert_maps(pos, count)
    _shift_sheet_meta(sheet, col_maps=maps)
    _shift_formulas_workbook(data, sheet["name"], col_maps=maps)


def _finish_col_delete(data: dict, sheet: dict, regions: list, col: int) -> None:
    maps = _delete_maps({col})
    _restamp_merges(sheet["rows"], _regions_after_delete(regions, maps, "col"))
    _shift_col_dims(sheet, col, -1)
    _shift_sheet_meta(sheet, col_maps=maps)
    _shift_formulas_workbook(data, sheet["name"], col_maps=maps)


def _strip_private(obj):
    """Drop internal keys (_fill_raw, _font_raw, …) from tool output JSON."""
    if isinstance(obj, dict):
        return {k: _strip_private(v) for k, v in obj.items()
                if not (isinstance(k, str) and k.startswith("_"))}
    if isinstance(obj, list):
        return [_strip_private(x) for x in obj]
    return obj


_SHEETNAME_SIMPLE_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_.]*$")


def _formula_sheet_prefix(name: str) -> str:
    """Sheet prefix for formulas — quoted when Excel requires quoting
    (spaces/specials, or names that look like cell references such as D2)."""
    looks_like_ref = bool(re.match(r"^[A-Za-z]{1,3}\d+$", name)) or \
        bool(re.match(r"^[Rr]\d|^[Cc]\d", name))
    if _SHEETNAME_SIMPLE_RE.match(name) and not looks_like_ref:
        return f"{name}!"
    return "'" + name.replace("'", "''") + "'!"


def _rename_sheet_in_formula(text, old: str, new: str):
    """Rewrite 'Old Name'! / OldName! references in a formula/defined-name."""
    if not text or not isinstance(text, str):
        return text
    new_ref = _formula_sheet_prefix(new)
    text = text.replace("'" + old.replace("'", "''") + "'!", new_ref)
    if _SHEETNAME_SIMPLE_RE.match(old):
        text = re.sub(rf"(?<![A-Za-z0-9_.!'\"]){re.escape(old)}!",
                      lambda m: new_ref, text)
    return text


def _remap_named_range_sheet_ids(data: dict, mapper) -> None:
    """Remap localSheetId indices after sheet add/copy/move/delete; None = drop."""
    kept = []
    for nr in data.get("named_ranges") or []:
        sid = nr.get("sheet_id")
        if sid is None:
            kept.append(nr)
            continue
        new = mapper(int(sid))
        if new is None:
            continue
        nr["sheet_id"] = new
        kept.append(nr)
    data["named_ranges"] = kept


def _dedupe_table_names(data: dict, sheet: dict) -> None:
    """Table displayNames must be unique workbook-wide — rename clones."""
    existing = {t["name"]
                for s in data["sheets"] if s is not sheet
                for t in (s.get("tables") or [])}
    for t in (sheet.get("tables") or []):
        if t["name"] in existing:
            i = 2
            while f"{t['name']}_{i}" in existing:
                i += 1
            t["name"] = f"{t['name']}_{i}"
        existing.add(t["name"])


# ── Formula handling ──────────────────────────────────────────────────────────

def _is_formula_value(cd: dict) -> bool:
    v = cd.get("v")
    return isinstance(v, str) and v.startswith("=") and cd.get("dt") != "s"


def _normalize_input_value(value):
    """Apply the Excel-like input contract for written cell values.

    Returns (value, force_text). A leading apostrophe forces the rest to be
    stored as literal text (the apostrophe itself is NOT stored — it becomes
    the quotePrefix style flag, exactly like typing 'text in Excel).
    A leading "=" that does not even tokenize as a formula is stored as text.
    """
    if isinstance(value, str):
        if value.startswith("'"):
            return value[1:], True
        if value.startswith("="):
            from openpyxl.formula import Tokenizer
            try:
                Tokenizer(value)
            except Exception:
                return value, True
    return value, False


def _store_cell_value(cell: dict, value) -> None:
    value, force_text = _normalize_input_value(value)
    cell["v"] = value
    if force_text:
        cell["dt"] = "s"
        cell["qp"] = True
    else:
        cell.pop("dt", None)
        cell.pop("qp", None)


def _split_sheet_prefix(ref: str):
    """'My Sheet'!A5 → ("My Sheet", "A5"); Data!C3 → ("Data", "C3"); A5 → (None, "A5")."""
    m = re.match(r"^'((?:[^']|'')+)'!(.*)$", ref)
    if m:
        return m.group(1).replace("''", "'"), m.group(2)
    m = re.match(r"^([A-Za-z_][A-Za-z0-9_.]*)!(.*)$", ref)
    if m:
        return m.group(1), m.group(2)
    return None, ref


def _shift_formula_str(formula: str, current_sheet, target_sheet: str,
                       row_maps=None, col_maps=None) -> str:
    """Shift cell/range references in one formula that target target_sheet.

    Only RANGE operand tokens are touched (string literals, names and
    structured table references pass through untouched). References whose
    whole area was deleted become #REF!. Any parse problem leaves the
    formula unchanged — never worse than not shifting.
    """
    from openpyxl.formula import Tokenizer
    try:
        tok = Tokenizer(formula)
        changed = False
        for t in tok.items:
            if t.type != "OPERAND" or t.subtype != "RANGE":
                continue
            sheet, rest = _split_sheet_prefix(t.value)
            if sheet is None:
                if current_sheet != target_sheet:
                    continue
            elif sheet != target_sheet:
                continue
            new_rest = _shift_ref(rest, row_maps, col_maps)
            if new_rest is None:
                new_rest = "#REF!"
            if new_rest != rest:
                t.value = t.value[: len(t.value) - len(rest)] + new_rest
                changed = True
        return tok.render() if changed else formula
    except Exception:
        return formula


def _shift_formulas_workbook(data: dict, target_sheet: str,
                             row_maps=None, col_maps=None) -> None:
    """After a structural edit on target_sheet, rewrite formulas everywhere
    (any sheet may reference the edited sheet) and defined names."""
    for sd in data["sheets"]:
        current = sd["name"]
        for row in sd["rows"]:
            for cd in row.get("cells") or []:
                if not _is_formula_value(cd):
                    continue
                new_v = _shift_formula_str(cd["v"], current, target_sheet,
                                           row_maps, col_maps)
                if new_v != cd["v"]:
                    cd["v"] = new_v
    for nr in data.get("named_ranges") or []:
        val = nr.get("value")
        if isinstance(val, str) and val:
            # defined-name refs are always sheet-qualified; current_sheet=None
            nr["value"] = _shift_formula_str("=" + val, None, target_sheet,
                                             row_maps, col_maps)[1:]


def _rename_sheet_in_cell_formulas(data: dict, old: str, new: str) -> None:
    from openpyxl.formula import Tokenizer
    new_prefix = _formula_sheet_prefix(new)
    for sd in data["sheets"]:
        for row in sd["rows"]:
            for cd in row.get("cells") or []:
                if not _is_formula_value(cd):
                    continue
                try:
                    tok = Tokenizer(cd["v"])
                    changed = False
                    for t in tok.items:
                        if t.type == "OPERAND" and t.subtype == "RANGE":
                            sheet, rest = _split_sheet_prefix(t.value)
                            if sheet == old:
                                t.value = new_prefix + rest
                                changed = True
                    if changed:
                        cd["v"] = tok.render()
                except Exception:
                    pass


_INSTRUCTIONS = """\
excel-tools MCP — Excel round-trip editing.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
EXCEL EDITING WORKFLOW
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Standard flow:
  1. excel_get_info     — lightweight: sheet names + dimensions, no session needed
  2. excel_load         — load file into server session → returns session_key
  3. excel_to_markdown  — read full content; rows annotated with 0-based row_index,
                          columns annotated as col_0, col_1, … so you know exact
                          indices to pass to edit/delete/insert tools
  4. <edit tools>       — excel_edit_cells / excel_insert_rows / excel_delete_rows / …
  5. excel_save         — write session back to .xlsx (omit output_path to overwrite
                          the original file; pass a new path to save elsewhere)

IMPORTANT — index conventions:
  • All row_index and col_index values are 0-based.
  • excel_to_markdown shows them explicitly — use those values directly.
  • Merge info (r1 c1 r2 c2) returned by excel_get_cell is also 0-based.
  • Merged cells: origin shows [M rowspan×colspan]; slave cells are blank.
    Only edit/delete the ORIGIN cell of a merged range.

IMPORTANT — multi-insert pattern (avoid index drift):
  1. Clone ALL template rows first (excel_clone_rows) before any inserts.
  2. Insert from BOTTOM to TOP (largest after_index first).
  This ensures earlier rows are not shifted before later inserts.

IMPORTANT — automatic reference shifting:
  • Inserting/deleting rows & columns automatically shifts merges, hyperlinks,
    comments, data validations, conditional-formatting ranges, tables,
    auto-filter, freeze panes, print titles, image/chart anchors, AND cell
    formulas / defined names referencing the edited sheet (from every loaded
    sheet). References whose entire area was deleted become #REF!.
  • Caveat: with a sheet_name-filtered session, formulas on UNLOADED sheets
    are not rewritten — load the full workbook before structural edits if
    other sheets reference the edited one.
  • Editing a slave (merged) cell raises an error — edit the origin cell.

IMPORTANT — writing values that start with = + - :
  • A value starting with "=" is stored as a FORMULA.
  • To store literal text that starts with "=" (or any text), prefix it with
    a single apostrophe: "'=not a formula" → cell text «=not a formula».
    The apostrophe is not stored; it becomes Excel's quote-prefix flag,
    exactly like typing 'text in Excel.
  • Values starting with "+" or "-" are already stored as text automatically
    when they are not numbers — no prefix needed.
  • Merging over an existing merged region raises an error — unmerge first.
  • convert_to_markdown is read-only and can accept Excel-family files readable
    by the converter.
  • Session/edit/save tools support OOXML Excel packages: .xlsx, .xlsm,
    .xltx, and .xltm. Macro/template parts are preserved best-effort.
    Legacy/binary .xls and .xlsb need read-only conversion or conversion to
    OOXML before editing.
  • excel_load with sheet_name loads ONLY that sheet, but excel_save merges
    the other sheets back from disk automatically — nothing is lost.
  • excel_save validates the generated .xlsx before replacing the destination.
    If validation fails, the existing destination file is left untouched.
  • Advanced DrawingML/charts/images/unknown OOXML parts are preserved
    best-effort. Use excel_validate_workbook and excel_diff_package for risky
    files before trusting a save workflow.

IMPORTANT — session_key:
  • session_key is the absolute file path returned by excel_load.
  • It persists server-side across all tool calls in this conversation.
  • You do NOT need to reload the file between edits — just reuse the same key.
  • Call excel_reload to discard in-memory changes and re-read from disk.
  • Call excel_close when done to free server memory.

Quick reference — one-shot conversion:
  convert_to_markdown  read-only Markdown export; supports sheet/range/max limits
  excel_get_workbook_summary compact file summary without excel_load/session
  excel_get_sheet_preview compact top-left sheet preview without session

Quick reference — session lifecycle:
  excel_get_info       sheet names + dimensions (no session needed)
  excel_load           load file → session_key
  excel_save           write session back to .xlsx (output_path optional)
  excel_save_as_copy   save to a different .xlsx path without overwriting source
  excel_validate_workbook validate .xlsx ZIP/XML structure + feature summary
  excel_diff_package   compare before/after package manifests
  excel_reload         reload from disk, discard unsaved changes
  excel_close          remove session from cache, free memory

Quick reference — sheet management:
  excel_add_sheet      add a new empty sheet
  excel_delete_sheet   delete a sheet (cannot delete the only sheet)
  excel_rename_sheet   rename a sheet
  excel_copy_sheet     duplicate a sheet within the same workbook
  excel_move_sheet     reorder a sheet to a new position

Quick reference — reading:
  excel_to_markdown    annotated Markdown view; supports max_rows/max_cols
  excel_to_markdown_range Markdown table for one A1/0-based range
  excel_list_tables    list Excel table objects captured in session
  excel_list_defined_names list workbook defined names/named ranges
  excel_get_rows       row range as JSON; values_only=True for compact output
  excel_read_range     exact A1 or 0-based rectangular range; token-efficient
  excel_get_cell       single cell with full style metadata
  excel_get_column     all cells in a column
  excel_find_cells     find literal/regex values or formulas across workbook
  excel_get_shapes     list captured DrawingML shape/image/chart metadata

Quick reference — editing rows:
  excel_edit_cells     edit cell values across one or more rows
  excel_insert_rows    insert rows at one or more positions in one call
  excel_clone_rows     deep-clone rows → JSON for modification before insert
  excel_copy_row       clone a row AND insert immediately (one step)
  excel_delete_rows    delete rows by index list OR start_row+end_row range
  excel_clear_range    clear values and/or styles from a rectangular range

Quick reference — editing columns:
  excel_insert_column  insert a new empty column at a position
  excel_copy_column    copy a column to a new position
  excel_delete_column  delete a column from all rows

Quick reference — merge:
  excel_merge_cells    merge a range (unmerge=False) or unmerge origin (unmerge=True)

Quick reference — formatting:
  excel_set_style      set style on a cell/range (fill/fcolor/font/strike/align/numfmt)
  excel_set_font_color set font color on a cell or range
  excel_set_strike     enable/disable strikethrough on a cell or range
  excel_set_borders    set/remove borders on a cell range
  excel_set_dimension  set row height (axis="row") or column width (axis="col")
                       • axis="row", index=3, size=20   → set row 3 height to 20pt
                       • axis="col", index=1, size=22   → set col 1 (B) width to 22
                       • size=null resets to auto
  excel_autofit_cols   auto-fit column widths to content (heuristic)
  excel_freeze_panes   freeze header rows and/or columns
  excel_set_data_validation  add dropdown list validation to a cell range
  excel_update_shape_text    update simple DrawingML textbox/shape text
  excel_set_shape_style      set simple DrawingML shape fill/outline/text color

Quick reference — search & fill:
  excel_find_rows      find rows matching a value or regex in a column
  excel_fill_column    fill a column range with a constant or sequence
  excel_fill_rows      clone a template row N times and insert (stamp pattern)
"""

mcp = FastMCP("excel-tools", instructions=_INSTRUCTIONS)


# ── 1. Info ───────────────────────────────────────────────────────────────────

@mcp.tool()
def convert_to_markdown(file_path: str, sheet_name: str | None = None, range_ref: str | None = None, max_rows: int | None = None, max_cols: int | None = None, include_styles: bool = False) -> TextContent:
    """
    Convert an Excel-family file to Markdown in one call without creating a session.

    This read-only conversion is for quick inspection/export workflows. Use
    excel_load + excel_to_markdown when you need Markdown generated from an
    editable .xlsx in-memory session, including unsaved changes.

    Args:
        file_path: Path or file:// URI to an Excel-family file readable by the converter.
        sheet_name: Sheet to export; omit to export all sheets.
        range_ref: Optional A1 range to export from sheet_name.
        max_rows: Optional maximum rows per exported sheet/range.
        max_cols: Optional maximum columns per exported sheet/range.
        include_styles: Reserved for clients; current Markdown export remains content-focused.

    Returns:
        Markdown content (text/markdown) representing workbook sheets.
    """
    from excel_converter import convert_excel_to_markdown

    path = uri_to_path(file_path)
    data = serialize_excel(str(path), sheet_name)
    if range_ref:
        if not sheet_name:
            raise ValueError("sheet_name is required when range_ref is provided")
        r1, r2, c1, c2 = _excel_range_to_indices(range_ref)
        for sheet in data.get("sheets", []):
            sheet["rows"] = [
                {**row, "cells": row.get("cells", [])[c1:c2 + 1]}
                for row in sheet.get("rows", [])[r1:r2 + 1]
            ]
    data = _limit_workbook_data(data, max_rows=max_rows, max_cols=max_cols)
    markdown = convert_excel_to_markdown(data)
    return TextContent(type="text", text=markdown, mimeType="text/markdown")

@mcp.tool()
def excel_get_info(uri: str) -> str:
    """
    Return summary info about an Excel file: sheet names, row and column counts.

    Use this first to understand the file structure before loading.

    Args:
        uri: Local file path or file:// URI to the .xlsx file

    Returns:
        JSON: {source, sheets: [{name, max_row, max_column}]}
    """
    import openpyxl
    path = uri_to_path(uri)
    wb = openpyxl.load_workbook(
        str(path),
        read_only=False,
        data_only=False,
        keep_vba=path.suffix.lower() in {".xlsm", ".xltm"},
    )
    try:
        info = {
            "source": str(path),
            "sheets": [],
        }
        for name in wb.sheetnames:
            ws = wb[name]
            info["sheets"].append({
                "name": name,
                "max_row": ws.max_row,
                "max_column": ws.max_column,
                "state": ws.sheet_state,
                "hidden": ws.sheet_state != "visible",
                "freeze_panes": ws.freeze_panes,
                "merged_ranges": len(ws.merged_cells.ranges),
                "table_count": len(getattr(ws, "tables", {}) or {}),
            })
    finally:
        wb.close()
    return json.dumps(info, ensure_ascii=False)


# ── 2. Load ───────────────────────────────────────────────────────────────────

@mcp.tool()
def excel_load(uri: str, sheet_name: str | None = None) -> str:
    """
    Load an Excel file into the server session cache and return a session_key.

    The session_key is then passed to all other excel_* tools.

    Args:
        uri: Local file path or file:// URI to an OOXML Excel workbook
             (.xlsx/.xlsm/.xltx/.xltm). Legacy/binary .xls/.xlsb are not
             handled by the edit engine.
        sheet_name: Sheet to load; omit to load ALL sheets. When a filter is
             used, excel_save automatically merges the unloaded sheets back
             from the file on disk — they are NOT lost.

    Returns:
        session_key string to pass to other tools
    """
    _check_supported(uri_to_path(uri))
    data = serialize_excel(uri, sheet_name)
    data["_sheet_filter"] = sheet_name
    data["_loaded_disk_names"] = [s["name"] for s in data["sheets"]]
    session_key = str(Path(data["source"]).resolve())
    data["source"] = session_key
    _sessions[session_key] = data
    sheet_names = [s["name"] for s in data["sheets"]]
    total_rows = sum(len(s["rows"]) for s in data["sheets"])
    return f"Loaded: session_key={session_key!r} | sheets={sheet_names} | total_rows={total_rows}"


# ── 3. Save ───────────────────────────────────────────────────────────────────

@mcp.tool()
def excel_save(session_key: str, output_path: str | None = None) -> str:
    """
    Reconstruct an Excel file from session data and write it to disk.

    If the session was loaded with a sheet_name filter, the unloaded sheets are
    merged back from the file on disk so they are never lost.

    Args:
        session_key: Key returned by excel_load
        output_path: Absolute path where the .xlsx will be written.
                     Omit to overwrite the original file loaded by excel_load.

    Returns:
        Summary: saved path, file size, sheet count, total rows.
        Any WARNINGS line means a passthrough feature could not be restored.
    """
    data = _get_session(session_key)
    if output_path:
        _check_supported(output_path)
    dest = output_path or data["source"]
    _check_save_extension_compatible(data["source"], dest)

    to_write = data
    if data.get("_sheet_filter"):
        try:
            full = serialize_excel(data["source"])
        except Exception as e:
            raise ValueError(
                "Session was loaded with a sheet_name filter and the original file "
                f"can no longer be read to merge the unloaded sheets back ({e}). "
                "Fix the file on disk or reload without a filter.")
        loaded = set(data.get("_loaded_disk_names") or [])
        merged, spliced = [], False
        for s in full["sheets"]:
            if s["name"] in loaded:
                if not spliced:
                    merged.extend(data["sheets"])
                    spliced = True
            else:
                merged.append(s)
        if not spliced:
            merged.extend(data["sheets"])
        names = [s["name"] for s in merged]
        duplicates = sorted({n for n in names if names.count(n) > 1})
        if duplicates:
            raise ValueError(
                f"Sheet name collision while merging the filtered session back: {duplicates}. "
                "Rename the session sheet or save to a different output_path.")
        to_write = {**full, "sheets": merged}

    warnings = reconstruct_excel(to_write, dest)
    size = Path(dest).stat().st_size
    msg = (
        f"Saved: {dest} ({size // 1024} KB)\n"
        f"Sheets: {len(to_write['sheets'])}, "
        f"Total rows: {sum(len(s['rows']) for s in to_write['sheets'])}"
    )
    if data.get("_sheet_filter"):
        msg += "\nNote: sheet-filtered session — unloaded sheets were merged back from disk."
    if warnings:
        msg += "\nWARNINGS: " + "; ".join(warnings)
    return msg


@mcp.tool()
def excel_save_as_copy(session_key: str, output_path: str) -> str:
    """Save the session to a new .xlsx path without overwriting the source workbook."""
    if not output_path:
        raise ValueError("output_path is required for excel_save_as_copy.")
    source = Path(_get_session(session_key)["source"]).resolve()
    dest = Path(output_path).resolve()
    if source == dest:
        raise ValueError("excel_save_as_copy output_path must differ from the source. Use excel_save to overwrite.")
    return excel_save(session_key, str(dest))


@mcp.tool()
def excel_validate_workbook(path: str) -> str:
    """Validate an .xlsx package and report advanced features found."""
    return json.dumps(inspect_xlsx_package(str(uri_to_path(path))), ensure_ascii=False, indent=2)


@mcp.tool()
def excel_diff_package(before_path: str, after_path: str) -> str:
    """Compare two .xlsx ZIP package manifests for save diagnostics."""
    return json.dumps(
        diff_xlsx_package(str(uri_to_path(before_path)), str(uri_to_path(after_path))),
        ensure_ascii=False,
        indent=2,
    )


# ── 4. Reload / Close ─────────────────────────────────────────────────────────

@mcp.tool()
def excel_reload(session_key: str) -> str:
    """
    Reload session data from disk, discarding any unsaved in-memory changes.

    Args:
        session_key: Key returned by excel_load (must still point to a valid file)

    Returns:
        Same summary as excel_load
    """
    session_key = _resolve_session_key(session_key)
    sheet_filter = _sessions[session_key].get("_sheet_filter")
    data = serialize_excel(session_key, sheet_filter)
    data["_sheet_filter"] = sheet_filter
    data["_loaded_disk_names"] = [s["name"] for s in data["sheets"]]
    data["source"] = session_key
    _sessions[session_key] = data
    sheet_names = [s["name"] for s in data["sheets"]]
    total_rows = sum(len(s["rows"]) for s in data["sheets"])
    return f"Reloaded: session_key={session_key!r} | sheets={sheet_names} | total_rows={total_rows}"


@mcp.tool()
def excel_close(session_key: str) -> str:
    """
    Remove a session from the server cache to free memory.

    Args:
        session_key: Key returned by excel_load

    Returns:
        Confirmation
    """
    session_key = _resolve_session_key(session_key)
    del _sessions[session_key]
    return f"Closed session '{session_key}'."


# ── 5. To Markdown ────────────────────────────────────────────────────────────

@mcp.tool()
def excel_to_markdown(session_key: str, sheet_name: str | None = None, max_rows: int | None = None, max_cols: int | None = None) -> TextContent:
    """
    Export session data as Markdown tables annotated with 0-based row/column indices.

    Column headers show col_N (header text) using row 0 as the header row.
    When a sheet contains merged cells a "merge" column is inserted showing RxC
    for origin cells. Slave cells are shown as blank.

    Args:
        session_key: Key returned by excel_load
        sheet_name: Sheet to export; omit to export all sheets
        max_rows: Optional maximum rows per exported sheet
        max_cols: Optional maximum columns per exported sheet

    Returns:
        Markdown content (text/markdown) — one table per sheet
    """
    data = _limit_workbook_data(_get_session(session_key), max_rows=max_rows, max_cols=max_cols)
    from excel_converter import convert_excel_to_markdown
    markdown = convert_excel_to_markdown(data, sheet_name=sheet_name)
    return TextContent(type="text", text=markdown, mimeType="text/markdown")


# ── LibreOffice capture ───────────────────────────────────────────────────────

def _find_soffice(hint: str | None = None) -> str:
    import shutil
    if hint:
        return hint
    found = shutil.which("soffice")
    if found:
        return found
    from pathlib import Path
    for candidate in (
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ):
        if Path(candidate).exists():
            return candidate
    raise FileNotFoundError(
        "LibreOffice (soffice) not found. Install LibreOffice or pass soffice_path explicitly."
    )


@mcp.tool()
def excel_capture(
    session_key: str,
    sheet_name: str,
    output_path: str,
    soffice_path: str | None = None,
) -> str:
    """
    Render a sheet as a PNG image using LibreOffice.

    The target sheet is exported from the current in-memory session (unsaved
    edits are included). LibreOffice renders the used range and saves a PNG.

    Args:
        session_key: Key returned by excel_load
        sheet_name: Sheet to render
        output_path: Where to save the PNG (e.g. "C:/tmp/sheet.png")
        soffice_path: Path to soffice.exe; auto-detected if omitted

    Returns:
        Confirmation with output path and image dimensions
    """
    import subprocess
    import tempfile
    import shutil
    from pathlib import Path

    lo = _find_soffice(soffice_path)
    data = _get_session(session_key)
    sheet = _find_sheet(data, sheet_name)

    tmp_dir = Path(tempfile.mkdtemp())
    try:
        # Write only the target sheet so LibreOffice produces exactly one PNG
        tmp_xlsx = tmp_dir / "capture.xlsx"
        reconstruct_excel({"source": "", "sheets": [sheet]}, str(tmp_xlsx))

        # Isolated user profile so a running LibreOffice instance cannot
        # silently block the headless conversion.
        profile = tmp_dir / "lo_profile"
        result = subprocess.run(
            [lo, "--headless", "--norestore", "--nofirststartwizard",
             f"-env:UserInstallation=file:///{profile.as_posix()}",
             "--convert-to", "png", "--outdir", str(tmp_dir), str(tmp_xlsx)],
            capture_output=True, text=True, timeout=120,
        )

        pngs = sorted(tmp_dir.glob("*.png"))
        if not pngs:
            detail = (result.stdout + result.stderr).strip()
            raise RuntimeError(f"LibreOffice produced no PNG. {detail}")

        out = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy(str(pngs[0]), str(out))

        # Read PNG dimensions from header (no extra library needed)
        with open(out, "rb") as f:
            f.seek(16)
            w = int.from_bytes(f.read(4), "big")
            h = int.from_bytes(f.read(4), "big")

        return f"Saved {w}×{h}px → {out}"
    finally:
        shutil.rmtree(str(tmp_dir), ignore_errors=True)


@mcp.tool()
def excel_extract_images(
    session_key: str,
    sheet_name: str,
    output_dir: str,
) -> str:
    """
    Extract all embedded images from a sheet and save them to a directory.

    Images are read directly from the source file on disk (not from in-memory
    edits). Call excel_save first if you want the saved state to be reflected.

    Args:
        session_key: Key returned by excel_load (used as the source file path)
        sheet_name: Name of the sheet to extract images from
        output_dir: Directory to save images into (created if it does not exist)

    Returns:
        List of saved image paths with their anchor cell references
    """
    import openpyxl
    import openpyxl.utils
    from pathlib import Path

    def _ext(data: bytes) -> str:
        if data[:8] == b'\x89PNG\r\n\x1a\n':
            return 'png'
        if data[:2] == b'\xff\xd8':
            return 'jpg'
        if data[:4] == b'GIF8':
            return 'gif'
        if data[:4] == b'RIFF' and data[8:12] == b'WEBP':
            return 'webp'
        return 'bin'

    def _anchor_cell(img) -> str:
        a = img.anchor
        if isinstance(a, str):
            return a
        try:
            fr = a._from
            cell = f"{openpyxl.utils.get_column_letter(fr.col + 1)}{fr.row + 1}"
            if hasattr(a, 'to'):
                t = a.to
                cell += f":{openpyxl.utils.get_column_letter(t.col + 1)}{t.row + 1}"
            return cell
        except Exception:
            return str(a)

    _get_session(session_key)  # validate session exists
    wb = openpyxl.load_workbook(session_key)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found.")
    ws = wb[sheet_name]

    images = ws._images
    if not images:
        return f"No images found in sheet '{sheet_name}'."

    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)

    saved = []
    for i, img in enumerate(images, 1):
        data = img._data()
        ext = _ext(data)
        name = f"image_{i:02d}.{ext}"
        path = out / name
        path.write_bytes(data)
        saved.append(f"{path}  [{_anchor_cell(img)}]")

    lines = [f"Extracted {len(saved)} image(s) from sheet '{sheet_name}':"] + saved
    return "\n".join(lines)


# ── 6. Get rows ───────────────────────────────────────────────────────────────

@mcp.tool()
def excel_get_rows(
    session_key: str,
    sheet_name: str,
    start_row: int = 0,
    end_row: int | None = None,
    values_only: bool = False,
) -> str:
    """
    Get a range of rows from session data as JSON.

    Use start_row + end_row for pagination (e.g. start_row=0, end_row=100).

    Args:
        session_key: Key returned by excel_load
        sheet_name: Name of the sheet
        start_row: 0-based start index (inclusive), default 0
        end_row: 0-based end index (exclusive); omit for all rows
        values_only: If True, return [[value, ...], ...] without style metadata

    Returns:
        JSON — full row objects [{h, cells: [{v, fill, bold, ...}]}] by default,
        or [[value, ...], ...] when values_only=True (slave cells → null)
    """
    data = _get_session(session_key)
    sheet = _find_sheet(data, sheet_name)
    rows = sheet["rows"][start_row:end_row]
    if values_only:
        result = [
            [cd["v"] if cd["merge"] != "slave" else None for cd in row["cells"]]
            for row in rows
        ]
    else:
        result = _strip_private(rows)
    return json.dumps(result, default=str, ensure_ascii=False)


@mcp.tool()
def excel_read_range(
    session_key: str,
    sheet_name: str,
    range_ref: str | None = None,
    start_row: int | None = None,
    end_row: int | None = None,
    start_col: int | None = None,
    end_col: int | None = None,
    values_only: bool = True,
) -> str:
    """
    Read an exact rectangular range from a loaded worksheet.

    Provide either range_ref (for example "A1:D20") or 0-based bounds where
    end_row/end_col are exclusive. This is more token-efficient than reading a
    whole sheet when only a small area is needed.
    """
    data = _get_session(session_key)
    sheet = _find_sheet(data, sheet_name)
    rows = sheet.get("rows", [])
    max_row = len(rows)
    max_col = max((len(row.get("cells", [])) for row in rows), default=0)
    if max_row == 0 or max_col == 0:
        return json.dumps({"sheet_name": sheet_name, "range": None, "values": []}, ensure_ascii=False)
    r1, r2, c1, c2 = _range_from_args(range_ref, start_row, end_row, start_col, end_col, max_row, max_col)
    r2 = min(r2, max_row - 1)
    c2 = min(c2, max_col - 1)
    values = []
    for row_index in range(r1, r2 + 1):
        cells = rows[row_index].get("cells", []) if row_index < len(rows) else []
        row_values = []
        for col_index in range(c1, c2 + 1):
            cell = cells[col_index] if col_index < len(cells) else None
            if values_only:
                row_values.append(None if not cell or cell.get("merge") == "slave" else cell.get("v"))
            else:
                row_values.append(_strip_private(cell) if cell else None)
        values.append(row_values)
    result = {
        "sheet_name": sheet_name,
        "range": {"start_row": r1, "end_row": r2, "start_col": c1, "end_col": c2},
        "values_only": values_only,
        "values": values,
    }
    return json.dumps(result, default=str, ensure_ascii=False)

@mcp.tool()
def excel_find_cells(
    session_key: str,
    query: str,
    sheet_name: str | None = None,
    regex: bool = False,
    case_sensitive: bool = False,
    match_in: str = "value",
    max_results: int = 100,
) -> str:
    """Find cells by literal text or regex across one sheet or the whole workbook."""
    data = _get_session(session_key)
    flags = 0 if case_sensitive else re.IGNORECASE
    pattern = re.compile(query if regex else re.escape(query), flags)
    sheets = [_find_sheet(data, sheet_name)] if sheet_name else data.get("sheets", [])
    results = []
    for sheet in sheets:
        for row_index, row in enumerate(sheet.get("rows", [])):
            for col_index, cell in enumerate(row.get("cells", [])):
                if cell.get("merge") == "slave":
                    continue
                haystack = cell.get("v")
                if match_in == "formula":
                    if not isinstance(haystack, str) or not haystack.startswith("="):
                        continue
                elif match_in != "value":
                    raise ValueError("match_in must be 'value' or 'formula'")
                text = "" if haystack is None else str(haystack)
                if pattern.search(text):
                    results.append({"sheet_name": sheet["name"], "row_index": row_index, "col_index": col_index, "value": haystack})
                    if len(results) >= max_results:
                        return json.dumps({"query": query, "truncated": True, "count": len(results), "matches": results}, default=str, ensure_ascii=False)
    return json.dumps({"query": query, "truncated": False, "count": len(results), "matches": results}, default=str, ensure_ascii=False)

@mcp.tool()
def excel_get_workbook_summary(file_path: str) -> str:
    """Return a compact read-only workbook summary without creating a session."""
    path = uri_to_path(file_path)
    data = serialize_excel(str(path))
    sheets = []
    for sheet in data.get("sheets", []):
        rows = sheet.get("rows", [])
        max_cols = max((len(row.get("cells", [])) for row in rows), default=0)
        formula_count = 0
        merge_origins = 0
        non_empty = 0
        for row in rows:
            for cell in row.get("cells", []):
                value = cell.get("v")
                if value not in (None, ""):
                    non_empty += 1
                if isinstance(value, str) and value.startswith("="):
                    formula_count += 1
                merge = cell.get("merge")
                if isinstance(merge, dict) and merge:
                    merge_origins += 1
        sheets.append({
            "name": sheet.get("name"),
            "rows": len(rows),
            "columns": max_cols,
            "non_empty_cells": non_empty,
            "formula_count": formula_count,
            "merged_ranges": merge_origins,
            "freeze": sheet.get("freeze"),
            "validations": len(sheet.get("validations") or []),
        })
    return json.dumps({"source": str(path), "sheet_count": len(sheets), "sheets": sheets}, default=str, ensure_ascii=False)
@mcp.tool()
def excel_to_markdown_range(session_key: str, sheet_name: str, range_ref: str | None = None, start_row: int | None = None, end_row: int | None = None, start_col: int | None = None, end_col: int | None = None) -> TextContent:
    """Export one worksheet range as a compact Markdown table."""
    data = json.loads(excel_read_range(session_key, sheet_name, range_ref, start_row, end_row, start_col, end_col, True))
    values = data["values"]
    if not values:
        return TextContent(type="text", text="", mimeType="text/markdown")
    headers = ["" if value is None else str(value) for value in values[0]]
    if not headers:
        return TextContent(type="text", text="", mimeType="text/markdown")
    rows = values[1:] if len(values) > 1 else []
    def cell(value):
        return "" if value is None else str(value).replace("|", "\\|").replace("\n", "<br>")
    lines = ["| " + " | ".join(cell(value) for value in headers) + " |"]
    lines.append("| " + " | ".join("---" for _ in headers) + " |")
    lines.extend("| " + " | ".join(cell((row + [None] * len(headers))[index]) for index in range(len(headers))) + " |" for row in rows)
    return TextContent(type="text", text="\n".join(lines), mimeType="text/markdown")

@mcp.tool()
def excel_list_tables(session_key: str, sheet_name: str | None = None) -> str:
    """List Excel table objects captured in the loaded workbook session."""
    data = _get_session(session_key)
    sheets = [_find_sheet(data, sheet_name)] if sheet_name else data.get("sheets", [])
    tables = []
    for sheet in sheets:
        for table in sheet.get("tables") or []:
            tables.append({"sheet_name": sheet["name"], **table})
    return json.dumps({"count": len(tables), "tables": tables}, default=str, ensure_ascii=False)

@mcp.tool()
def excel_list_defined_names(session_key: str) -> str:
    """List workbook defined names and named ranges from the loaded session."""
    data = _get_session(session_key)
    names = data.get("named_ranges") or []
    return json.dumps({"count": len(names), "defined_names": _strip_private(names)}, default=str, ensure_ascii=False)

@mcp.tool()
def excel_get_sheet_preview(file_path: str, max_rows: int = 20, max_cols: int = 10, sheet_name: str | None = None) -> str:
    """Return compact top-left previews for one sheet or all sheets without creating a session."""
    path = uri_to_path(file_path)
    data = serialize_excel(str(path), sheet_name)
    previews = []
    for sheet in data.get("sheets", []):
        rows = []
        for row in sheet.get("rows", [])[:max_rows]:
            values = []
            for cell in row.get("cells", [])[:max_cols]:
                values.append(None if cell.get("merge") == "slave" else cell.get("v"))
            rows.append(values)
        previews.append({"sheet_name": sheet["name"], "rows": rows, "truncated_rows": len(sheet.get("rows", [])) > max_rows})
    return json.dumps({"source": str(path), "max_rows": max_rows, "max_cols": max_cols, "sheets": previews}, default=str, ensure_ascii=False)
# ── 7. Get cell ───────────────────────────────────────────────────────────────

@mcp.tool()
def excel_get_cell(
    session_key: str,
    sheet_name: str,
    row_index: int,
    col_index: int,
) -> str:
    """
    Get full metadata of a single cell.

    Args:
        session_key: Key returned by excel_load
        sheet_name: Name of the sheet
        row_index: 0-based row index
        col_index: 0-based column index

    Returns:
        JSON object {v, fill, bold, italic, size, fcolor, wrap, halign, valign,
                     numfmt, merge, border}
    """
    data = _get_session(session_key)
    sheet = _find_sheet(data, sheet_name)
    rows = sheet["rows"]
    if not (0 <= row_index < len(rows)):
        raise ValueError(f"row_index {row_index} out of range (0–{len(rows)-1})")
    cells = rows[row_index]["cells"]
    if not (0 <= col_index < len(cells)):
        raise ValueError(f"col_index {col_index} out of range (0–{len(cells)-1})")
    return json.dumps(_strip_private(cells[col_index]), default=str, ensure_ascii=False)


# ── 8. Get column ─────────────────────────────────────────────────────────────

@mcp.tool()
def excel_get_column(
    session_key: str,
    sheet_name: str,
    col_index: int,
    start_row: int = 0,
    end_row: int | None = None,
) -> str:
    """
    Get all cells in a column as JSON.

    Args:
        session_key: Key returned by excel_load
        sheet_name: Name of the sheet
        col_index: 0-based column index
        start_row: 0-based start row (inclusive), default 0
        end_row: 0-based end row (exclusive); omit for all rows

    Returns:
        JSON array of {row_index, cell} objects
    """
    data = _get_session(session_key)
    sheet = _find_sheet(data, sheet_name)
    rows = sheet["rows"][start_row:end_row]
    result = [
        {
            "row_index": start_row + i,
            "cell": _strip_private(row["cells"][col_index]) if col_index < len(row["cells"]) else None,
        }
        for i, row in enumerate(rows)
    ]
    return json.dumps(result, default=str, ensure_ascii=False)


# ── 9. Sheet management ───────────────────────────────────────────────────────

@mcp.tool()
def excel_add_sheet(
    session_key: str,
    sheet_name: str,
    position: int | None = None,
) -> str:
    """
    Add a new empty sheet to the workbook session.

    Args:
        session_key: Key returned by excel_load
        sheet_name: Name for the new sheet (must be unique)
        position: 0-based position to insert at; omit to append at the end

    Returns:
        Confirmation with updated sheet list
    """
    data = _get_session(session_key)
    if any(s["name"] == sheet_name for s in data["sheets"]):
        raise ValueError(f"Sheet '{sheet_name}' already exists.")
    new_sheet = {"name": sheet_name, "cw": {}, "rows": [], "freeze": None, "validations": []}
    if position is None:
        data["sheets"].append(new_sheet)
        actual = len(data["sheets"]) - 1
    else:
        data["sheets"].insert(position, new_sheet)
        actual = position
        _remap_named_range_sheet_ids(data, lambda i: i + 1 if i >= position else i)
    return f"Added sheet '{sheet_name}' at position {actual}. Sheets: {[s['name'] for s in data['sheets']]}"


@mcp.tool()
def excel_delete_sheet(session_key: str, sheet_name: str) -> str:
    """
    Delete a sheet from the workbook session. Cannot delete the only sheet.

    Args:
        session_key: Key returned by excel_load
        sheet_name: Name of the sheet to delete

    Returns:
        Confirmation with updated sheet list
    """
    data = _get_session(session_key)
    idx = next((i for i, s in enumerate(data["sheets"]) if s["name"] == sheet_name), None)
    if idx is None:
        raise ValueError(f"Sheet '{sheet_name}' not found. Available: {[s['name'] for s in data['sheets']]}")
    if len(data["sheets"]) == 1:
        raise ValueError("Cannot delete the only sheet in a workbook.")
    data["sheets"].pop(idx)
    _remap_named_range_sheet_ids(
        data, lambda i: None if i == idx else (i - 1 if i > idx else i))
    return f"Deleted '{sheet_name}'. Remaining: {[s['name'] for s in data['sheets']]}"


@mcp.tool()
def excel_rename_sheet(session_key: str, sheet_name: str, new_name: str) -> str:
    """
    Rename a sheet in the workbook session.

    Args:
        session_key: Key returned by excel_load
        sheet_name: Current name of the sheet
        new_name: New name (must be unique)

    Returns:
        Confirmation
    """
    data = _get_session(session_key)
    if any(s["name"] == new_name for s in data["sheets"]):
        raise ValueError(f"Sheet '{new_name}' already exists.")
    _find_sheet(data, sheet_name)["name"] = new_name
    # Keep defined names and cell formulas pointing at the renamed sheet
    for nr in data.get("named_ranges") or []:
        nr["value"] = _rename_sheet_in_formula(nr.get("value"), sheet_name, new_name)
    _rename_sheet_in_cell_formulas(data, sheet_name, new_name)
    return f"Renamed '{sheet_name}' → '{new_name}'."


@mcp.tool()
def excel_copy_sheet(
    session_key: str,
    source_sheet: str,
    new_name: str,
    position: int | None = None,
) -> str:
    """
    Duplicate a sheet within the same workbook session.

    All rows, styles, column widths, freeze panes, and validations are copied.

    Args:
        session_key: Key returned by excel_load
        source_sheet: Name of the sheet to copy
        new_name: Name for the copy (must be unique)
        position: 0-based insertion position; omit to append

    Returns:
        Confirmation
    """
    data = _get_session(session_key)
    if any(s["name"] == new_name for s in data["sheets"]):
        raise ValueError(f"Sheet '{new_name}' already exists.")
    new_sheet = copy.deepcopy(_find_sheet(data, source_sheet))
    new_sheet["name"] = new_name
    if position is None:
        data["sheets"].append(new_sheet)
        actual = len(data["sheets"]) - 1
    else:
        data["sheets"].insert(position, new_sheet)
        actual = position
        _remap_named_range_sheet_ids(data, lambda i: i + 1 if i >= position else i)
    _dedupe_table_names(data, new_sheet)
    return f"Copied '{source_sheet}' → '{new_name}' at position {actual}."


@mcp.tool()
def excel_copy_sheet_to(
    src_session_key: str,
    src_sheet_name: str,
    dst_session_key: str,
    new_name: str | None = None,
    position: int | None = None,
) -> str:
    """
    Copy a sheet from one loaded workbook session into another.

    Both source and destination files must be loaded with excel_load first.
    After copying, call excel_save on the destination session_key to persist.

    Args:
        src_session_key: session_key of the source file
        src_sheet_name: Name of the sheet to copy from the source
        dst_session_key: session_key of the destination file
        new_name: Name for the sheet in the destination; defaults to src_sheet_name
        position: 0-based insertion position in the destination; omit to append

    Returns:
        Confirmation
    """
    src_data = _get_session(src_session_key)
    dst_data = _get_session(dst_session_key)

    sheet_copy = copy.deepcopy(_find_sheet(src_data, src_sheet_name))
    target_name = new_name or src_sheet_name
    if src_data.get("theme_xml") != dst_data.get("theme_xml"):
        _drop_raw_fills(sheet_copy)
    # Conditional-formatting raw XML references the SOURCE workbook's dxfs
    # style table — in the destination those ids point at the wrong styles.
    sheet_copy.pop("cf_xml", None)

    if any(s["name"] == target_name for s in dst_data["sheets"]):
        raise ValueError(f"Sheet '{target_name}' already exists in destination. Specify a different new_name.")

    sheet_copy["name"] = target_name
    if position is None:
        dst_data["sheets"].append(sheet_copy)
        actual = len(dst_data["sheets"]) - 1
    else:
        dst_data["sheets"].insert(position, sheet_copy)
        actual = position
        _remap_named_range_sheet_ids(dst_data, lambda i: i + 1 if i >= position else i)
    _dedupe_table_names(dst_data, sheet_copy)

    return (
        f"Copied sheet '{src_sheet_name}' from '{src_session_key}' "
        f"→ '{target_name}' in '{dst_session_key}' at position {actual}. "
        f"Call excel_save on the destination to persist."
    )


@mcp.tool()
def excel_move_sheet(
    session_key: str,
    sheet_name: str,
    position: int,
) -> str:
    """
    Move a sheet to a new position within the workbook.

    Args:
        session_key: Key returned by excel_load
        sheet_name: Name of the sheet to move
        position: New 0-based position

    Returns:
        Confirmation with updated sheet order
    """
    data = _get_session(session_key)
    idx = next((i for i, s in enumerate(data["sheets"]) if s["name"] == sheet_name), None)
    if idx is None:
        raise ValueError(f"Sheet '{sheet_name}' not found. Available: {[s['name'] for s in data['sheets']]}")
    n = len(data["sheets"])
    if not (0 <= position < n):
        raise ValueError(f"Position {position} out of range (0–{n-1}).")
    sheet = data["sheets"].pop(idx)
    data["sheets"].insert(position, sheet)
    # Remap localSheetId of defined names to the new sheet order
    order = list(range(n))
    moved = order.pop(idx)
    order.insert(position, moved)
    mapping = {old: new for new, old in enumerate(order)}
    _remap_named_range_sheet_ids(data, lambda i: mapping.get(i, i))
    return f"Moved '{sheet_name}' to position {position}. Order: {[s['name'] for s in data['sheets']]}"


# ── 10. Clone rows ────────────────────────────────────────────────────────────

@mcp.tool()
def excel_clone_rows(
    session_key: str,
    sheet_name: str,
    start_row: int,
    end_row: int | None = None,
) -> str:
    """
    Deep-clone one or more rows and return them as a JSON array WITHOUT inserting.

    Use with excel_insert_rows to insert the cloned block at one or more positions.
    Clone BEFORE any inserts to avoid index drift.

    For a single row: pass start_row only (end_row defaults to start_row).
    For a range:      pass start_row and end_row (both inclusive).

    Args:
        session_key: Key returned by excel_load
        sheet_name: Name of the sheet
        start_row: 0-based index of the first row to clone (inclusive)
        end_row: 0-based index of the last row to clone (inclusive);
                 defaults to start_row for single-row clone

    Returns:
        JSON array of cloned row objects
    """
    data = _get_session(session_key)
    sheet = _find_sheet(data, sheet_name)
    n = len(sheet["rows"])
    if end_row is None:
        end_row = start_row
    if not (0 <= start_row <= end_row < n):
        raise ValueError(f"Row range [{start_row}, {end_row}] out of bounds (0–{n-1})")
    return json.dumps(copy.deepcopy(sheet["rows"][start_row:end_row + 1]), default=str, ensure_ascii=False)


# ── 11. Copy row ──────────────────────────────────────────────────────────────

@mcp.tool()
def excel_copy_row(
    session_key: str,
    sheet_name: str,
    row_index: int,
    after_index: int,
) -> str:
    """
    Clone a row and insert the copy immediately at a new position (one step).

    Contrast with excel_clone_rows, which only returns the row as JSON so you
    can modify it before inserting.

    Args:
        session_key: Key returned by excel_load
        sheet_name: Name of the sheet
        row_index: 0-based index of the row to copy
        after_index: 0-based index to insert AFTER; use -1 to prepend

    Returns:
        Confirmation with new row count
    """
    data = _get_session(session_key)
    sheet = _find_sheet(data, sheet_name)
    n = len(sheet["rows"])
    if not (0 <= row_index < n):
        raise ValueError(f"row_index {row_index} out of range (0–{n-1})")
    cloned = copy.deepcopy(sheet["rows"][row_index])
    pos = after_index + 1
    _apply_row_insert(data, sheet, pos, [cloned])
    return f"Copied row {row_index} → inserted at position {pos}. Sheet '{sheet_name}' now has {len(sheet['rows'])} rows."


# ── 12. Insert rows ───────────────────────────────────────────────────────────

@mcp.tool()
def excel_insert_rows(
    session_key: str,
    sheet_name: str,
    inserts: list[dict],
) -> str:
    """
    Insert rows at one or more positions in a single call.

    Positions are automatically sorted bottom-to-top, so each after_index
    refers to the original row positions (before any insertions).

    For a single insert: pass a list with one item.

    Args:
        session_key: Key returned by excel_load
        sheet_name: Name of the target sheet
        inserts: List of {"after_index": int, "rows_json": list | str} objects.
                 after_index — 0-based original index to insert AFTER; -1 to prepend.
                 rows_json   — a row object, a list of row objects, or a JSON string
                               (as returned by excel_clone_rows).

    Returns:
        Summary: total rows inserted, new total row count
    """
    data = _get_session(session_key)
    sheet = _find_sheet(data, sheet_name)

    parsed: list[tuple[int, list]] = []
    for entry in inserts:
        rows_json = entry["rows_json"]
        if isinstance(rows_json, str):
            rows_json = json.loads(rows_json)
        if isinstance(rows_json, dict):
            rows_json = [rows_json]
        parsed.append((entry["after_index"], rows_json))

    parsed.sort(key=lambda x: x[0], reverse=True)

    total_inserted = 0
    for after_index, new_rows in parsed:
        pos = after_index + 1
        _apply_row_insert(data, sheet, pos, new_rows)
        total_inserted += len(new_rows)

    return (
        f"Inserted {total_inserted} row(s) across {len(parsed)} position(s). "
        f"Sheet '{sheet_name}' now has {len(sheet['rows'])} rows."
    )


# ── 13. Insert column ─────────────────────────────────────────────────────────

@mcp.tool()
def excel_insert_column(
    session_key: str,
    sheet_name: str,
    after_col_index: int,
) -> str:
    """
    Insert a new empty column after the given column index.

    Args:
        session_key: Key returned by excel_load
        sheet_name: Name of the sheet
        after_col_index: 0-based column index to insert AFTER; use -1 to prepend

    Returns:
        Confirmation
    """
    data = _get_session(session_key)
    sheet = _find_sheet(data, sheet_name)
    pos = after_col_index + 1
    regions = _capture_merge_regions(sheet["rows"])
    for row in sheet["rows"]:
        new_cell = copy.deepcopy(_EMPTY_CELL)
        row["cells"] = row["cells"][:pos] + [new_cell] + row["cells"][pos:]
    _finish_col_insert(data, sheet, regions, pos)
    return f"Inserted empty column at position {pos} in sheet '{sheet_name}'."


# ── 14. Edit cells ────────────────────────────────────────────────────────────

@mcp.tool()
def excel_edit_cells(
    session_key: str,
    sheet_name: str,
    edits: list[dict],
) -> str:
    """
    Edit cell values across one or more rows — styles are preserved.

    For a single row: pass a list with one item.

    Args:
        session_key: Key returned by excel_load
        sheet_name: Name of the sheet
        edits: List of {"row_index": int, "edits": {col_key: value}} objects.
               col_key is a 0-based column index as an int or string.
               Pass null as value to clear a cell.

    Returns:
        Summary: rows edited, cells updated
    """
    data = _get_session(session_key)
    sheet = _find_sheet(data, sheet_name)
    rows = sheet["rows"]
    n = len(rows)

    rows_edited = cells_updated = 0

    for entry in edits:
        r = entry["row_index"]
        if r < 0:
            raise ValueError(f"row_index must be >= 0, got {r}")
        # Auto-extend sheet rows (supports editing into an empty or short sheet)
        if r >= len(rows):
            needed_cols = max((int(k) + 1 for k in entry["edits"]), default=1)
            while r >= len(rows):
                rows.append({"h": None, "cells": [copy.deepcopy(_EMPTY_CELL) for _ in range(needed_cols)]})
        row = rows[r]
        updated = 0
        for col_key, value in entry["edits"].items():
            col = int(col_key)
            # Auto-extend cells in row if needed
            while col >= len(row["cells"]):
                row["cells"].append(copy.deepcopy(_EMPTY_CELL))
            if row["cells"][col].get("merge") == "slave":
                raise ValueError(
                    f"Cell [{r},{col}] is a slave cell of a merged range — its value "
                    "would be silently dropped on save. Edit the origin (top-left) "
                    "cell of the merge instead.")
            _store_cell_value(row["cells"][col], value)
            updated += 1
        if updated:
            rows_edited += 1
            cells_updated += updated

    return f"Edited {rows_edited} row(s): {cells_updated} cell(s) updated."


# ── 15. Delete rows ───────────────────────────────────────────────────────────

@mcp.tool()
def excel_delete_rows(
    session_key: str,
    sheet_name: str,
    row_indices: list[int] | None = None,
    start_row: int | None = None,
    end_row: int | None = None,
) -> str:
    """
    Delete one or more rows by index list or by a contiguous range.

    Provide row_indices OR start_row+end_row (or both — they are merged).
    end_row is EXCLUSIVE (Python convention): to delete rows 14–18 inclusive
    pass start_row=14, end_row=19.

    Args:
        session_key: Key returned by excel_load
        sheet_name: Name of the sheet
        row_indices: List of 0-based row indices to delete
        start_row: Start of a contiguous range (0-based, inclusive)
        end_row: End of a contiguous range (0-based, EXCLUSIVE)

    Returns:
        Confirmation with remaining row count
    """
    data = _get_session(session_key)
    sheet = _find_sheet(data, sheet_name)
    n = len(sheet["rows"])

    to_delete: set[int] = set(row_indices or [])
    if start_row is not None and end_row is not None:
        to_delete.update(range(start_row, end_row))
    elif start_row is not None or end_row is not None:
        raise ValueError("Provide both start_row and end_row together.")
    if not to_delete:
        raise ValueError("Provide row_indices or start_row+end_row.")

    invalid = [i for i in to_delete if not (0 <= i < n)]
    if invalid:
        raise ValueError(f"Row indices out of range (0–{n-1}): {sorted(invalid)}")
    _apply_row_delete(data, sheet, to_delete)
    return f"Deleted {len(to_delete)} row(s). Sheet '{sheet_name}' now has {len(sheet['rows'])} rows."


# ── 16. Clear range ───────────────────────────────────────────────────────────

@mcp.tool()
def excel_clear_range(
    session_key: str,
    sheet_name: str,
    r1: int,
    c1: int,
    r2: int,
    c2: int,
    clear_values: bool = True,
    clear_styles: bool = False,
) -> str:
    """
    Clear values and/or styles from a rectangular cell range.

    Slave cells of merged regions are skipped (their content belongs to the
    origin). To also remove merge structure, call excel_merge_cells with
    unmerge=True first.

    All coordinates are 0-based, inclusive on both ends.

    Args:
        session_key: Key returned by excel_load
        sheet_name: Name of the sheet
        r1: 0-based top row (inclusive)
        c1: 0-based left column (inclusive)
        r2: 0-based bottom row (inclusive)
        c2: 0-based right column (inclusive)
        clear_values: If True (default), set cell values to null
        clear_styles: If True, reset fill/font/alignment/border to defaults

    Returns:
        Summary: cells cleared
    """
    _STYLE_DEFAULTS = {
        "fill": None, "bold": False, "italic": False, "size": None,
        "fcolor": None, "wrap": False, "halign": None, "valign": None,
        "numfmt": "General", "border": {},
    }

    data = _get_session(session_key)
    sheet = _find_sheet(data, sheet_name)
    rows = sheet["rows"]

    cells_cleared = 0
    for r in range(r1, r2 + 1):
        if r >= len(rows):
            continue
        row_cells = rows[r]["cells"]
        for c in range(c1, c2 + 1):
            if c >= len(row_cells):
                continue
            cell = row_cells[c]
            if cell.get("merge") == "slave":
                continue
            if clear_values:
                cell["v"] = None
                cell.pop("dt", None)
                cell.pop("qp", None)
            if clear_styles:
                for k, v in _STYLE_DEFAULTS.items():
                    cell[k] = v
                cell.pop("_fill_raw", None)
                cell.pop("_font_raw", None)
                cell.pop("qp", None)
            cells_cleared += 1

    actions = []
    if clear_values:
        actions.append("values")
    if clear_styles:
        actions.append("styles")
    return (
        f"Cleared {' + '.join(actions)} in {cells_cleared} cell(s) "
        f"[{r1},{c1}]–[{r2},{c2}] in sheet '{sheet_name}'."
    )


# ── 17. Copy / Delete column ──────────────────────────────────────────────────

@mcp.tool()
def excel_copy_column(
    session_key: str,
    sheet_name: str,
    col_index: int,
    after_col_index: int,
) -> str:
    """
    Copy a column and insert it after a given column index, preserving all styles.

    Args:
        session_key: Key returned by excel_load
        sheet_name: Name of the sheet
        col_index: 0-based index of the column to copy
        after_col_index: 0-based index to insert AFTER; use -1 to prepend

    Returns:
        Confirmation
    """
    import openpyxl.utils
    data = _get_session(session_key)
    sheet = _find_sheet(data, sheet_name)
    pos = after_col_index + 1
    # Capture source width before shifting
    src_letter = openpyxl.utils.get_column_letter(col_index + 1)
    src_width = sheet["cw"].get(src_letter)
    regions = _capture_merge_regions(sheet["rows"])
    for row in sheet["rows"]:
        src = copy.deepcopy(row["cells"][col_index]) if col_index < len(row["cells"]) else copy.deepcopy(_EMPTY_CELL)
        if src.get("merge"):
            src["merge"] = {}  # copied cells never carry the source's merge state
        row["cells"] = row["cells"][:pos] + [src] + row["cells"][pos:]
    _finish_col_insert(data, sheet, regions, pos)
    if src_width:
        sheet["cw"][openpyxl.utils.get_column_letter(pos + 1)] = src_width
    return f"Copied column {col_index} → inserted at position {pos} in sheet '{sheet_name}'."


@mcp.tool()
def excel_delete_column(session_key: str, sheet_name: str, col_index: int) -> str:
    """
    Delete a column from all rows in a sheet.

    Args:
        session_key: Key returned by excel_load
        sheet_name: Name of the sheet
        col_index: 0-based column index to delete

    Returns:
        Confirmation
    """
    data = _get_session(session_key)
    sheet = _find_sheet(data, sheet_name)
    regions = _capture_merge_regions(sheet["rows"])
    removed = 0
    for row in sheet["rows"]:
        if col_index < len(row["cells"]):
            row["cells"].pop(col_index)
            removed += 1
    _finish_col_delete(data, sheet, regions, col_index)
    return f"Deleted column {col_index} from {removed} row(s) in sheet '{sheet_name}'."


# ── 18. Merge cells ───────────────────────────────────────────────────────────

@mcp.tool()
def excel_merge_cells(
    session_key: str,
    sheet_name: str,
    r1: int,
    c1: int,
    r2: int | None = None,
    c2: int | None = None,
    unmerge: bool = False,
) -> str:
    """
    Merge a rectangular range of cells, or unmerge a merged region.

    Merge:   pass r1, c1, r2, c2 (all required when unmerge=False).
             The top-left cell (r1, c1) becomes the origin and keeps its value.
             All other cells in the range are marked as slave cells.

    Unmerge: pass r1, c1 of the origin cell and unmerge=True.
             r2, c2 are ignored — the full merge region is found automatically.
             All cells in the region revert to independent cells.

    All coordinates are 0-based.

    Args:
        session_key: Key returned by excel_load
        sheet_name: Name of the sheet
        r1: Top row (0-based)
        c1: Left column (0-based)
        r2: Bottom row (0-based, inclusive) — required for merge, ignored for unmerge
        c2: Right column (0-based, inclusive) — required for merge, ignored for unmerge
        unmerge: If True, unmerge the region whose origin is at (r1, c1)

    Returns:
        Confirmation
    """
    data = _get_session(session_key)
    sheet = _find_sheet(data, sheet_name)
    rows = sheet["rows"]
    n_rows = len(rows)

    if unmerge:
        if not (0 <= r1 < n_rows):
            raise ValueError(f"r1={r1} out of range (0–{n_rows-1})")
        cell = rows[r1]["cells"][c1]
        if cell["merge"] == "slave":
            raise ValueError(f"Cell [{r1},{c1}] is a slave cell. Pass the origin (top-left) of the merge.")
        mi = cell["merge"]
        if not isinstance(mi, dict) or (mi.get("rowspan", 1) <= 1 and mi.get("colspan", 1) <= 1):
            raise ValueError(f"Cell [{r1},{c1}] is not a merge origin.")
        er1, ec1 = mi.get("r1", r1), mi.get("c1", c1)
        er2, ec2 = mi.get("r2", r1), mi.get("c2", c1)
        for r in range(er1, er2 + 1):
            for c in range(ec1, ec2 + 1):
                rows[r]["cells"][c]["merge"] = {}
        return f"Unmerged [{er1},{ec1}]–[{er2},{ec2}] in sheet '{sheet_name}'."

    # Merge
    if r2 is None or c2 is None:
        raise ValueError("r2 and c2 are required for merge. Pass unmerge=True to unmerge.")
    if not (0 <= r1 <= r2 < n_rows):
        raise ValueError(f"Row range [{r1}, {r2}] out of bounds (0–{n_rows-1})")
    for r in range(r1, r2 + 1):
        n_cols = len(rows[r]["cells"])
        if not (0 <= c1 <= c2 < n_cols):
            raise ValueError(f"Col range [{c1}, {c2}] out of bounds for row {r} (0–{n_cols-1})")

    # Reject overlap with any existing merged region (would corrupt the file)
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            mi = rows[r]["cells"][c].get("merge")
            if mi == "slave" or (isinstance(mi, dict)
                                 and (mi.get("rowspan", 1) > 1 or mi.get("colspan", 1) > 1)):
                raise ValueError(
                    f"Range overlaps an existing merged region at [{r},{c}]. "
                    "Unmerge it first (excel_merge_cells with unmerge=True).")

    rows[r1]["cells"][c1]["merge"] = {
        "r1": r1, "c1": c1, "r2": r2, "c2": c2,
        "rowspan": r2 - r1 + 1, "colspan": c2 - c1 + 1,
    }
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            if not (r == r1 and c == c1):
                rows[r]["cells"][c]["merge"] = "slave"
    return (
        f"Merged [{r1},{c1}]–[{r2},{c2}] "
        f"({r2-r1+1}×{c2-c1+1}) in sheet '{sheet_name}'."
    )


# ── 19. Style ─────────────────────────────────────────────────────────────────

@mcp.tool()
def excel_set_style(
    session_key: str,
    sheet_name: str,
    r1: int,
    c1: int,
    r2: int | None = None,
    c2: int | None = None,
    style: dict = {},
) -> str:
    """
    Set style properties on a single cell or a rectangular range.

    r2 and c2 default to r1 and c1 (single cell). For a range, pass all four.
    Only keys present in style are updated. Set a key to null to clear it.
    Slave cells of merged regions are skipped.

    Supported style keys:
      fill   — background color as ARGB hex (e.g. "FFDAEEF3"), null = no fill
      fcolor — font color as ARGB hex (e.g. "FF2E75B6"), null = auto
      bold   — bool
      italic — bool
      size   — font size in points (float), null = default
      font   — font family name (e.g. "Times New Roman", "Arial"), null = default
      wrap   — bool (wrap text)
      halign — "left" | "center" | "right" | "general" | null
      valign — "top" | "center" | "bottom" | null
      numfmt — Excel number format string, e.g. "0.00%", "General"

    Args:
        session_key: Key returned by excel_load
        sheet_name: Name of the sheet
        r1: 0-based top row
        c1: 0-based left column
        r2: 0-based bottom row (inclusive); defaults to r1
        c2: 0-based right column (inclusive); defaults to c1
        style: Dict of style overrides

    Returns:
        Summary: cells styled, range
    """
    if r2 is None:
        r2 = r1
    if c2 is None:
        c2 = c1

    data = _get_session(session_key)
    sheet = _find_sheet(data, sheet_name)
    rows = sheet["rows"]
    n_rows = len(rows)
    if not (0 <= r1 <= r2 < n_rows):
        raise ValueError(f"Row range [{r1}, {r2}] out of bounds (0–{n_rows-1})")

    cells_styled = 0
    for r in range(r1, r2 + 1):
        row_cells = rows[r]["cells"]
        for c in range(c1, min(c2 + 1, len(row_cells))):
            if row_cells[c].get("merge") != "slave":
                _apply_style(row_cells[c], style)
                cells_styled += 1

    rng = f"[{r1},{c1}]" if r1 == r2 and c1 == c2 else f"[{r1},{c1}]–[{r2},{c2}]"
    return f"Styled {cells_styled} cell(s) at {rng} ({list(style.keys())}) in sheet '{sheet_name}'."


@mcp.tool()
def excel_set_font_color(
    session_key: str,
    sheet_name: str,
    r1: int,
    c1: int,
    color: str | None,
    r2: int | None = None,
    c2: int | None = None,
) -> str:
    """Set font color on a cell or range. Color is ARGB hex, or null for auto."""
    return excel_set_style(session_key, sheet_name, r1, c1, r2, c2, {"fcolor": color})


@mcp.tool()
def excel_set_strike(
    session_key: str,
    sheet_name: str,
    r1: int,
    c1: int,
    enabled: bool = True,
    r2: int | None = None,
    c2: int | None = None,
) -> str:
    """Enable or disable strikethrough on a cell or range."""
    return excel_set_style(session_key, sheet_name, r1, c1, r2, c2, {"strike": enabled})



def _shape_anchors(drawing_xml: str) -> list[re.Match]:
    return list(re.finditer(
        r"<(?:(?:xdr:)?)(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b.*?</(?:(?:xdr:)?)(?:twoCellAnchor|oneCellAnchor|absoluteAnchor)>",
        drawing_xml,
        re.DOTALL,
    ))


def _rgb_hex(color: str | None) -> str | None:
    if color is None:
        return None
    value = color.strip().lstrip("#").upper()
    if len(value) == 8:
        value = value[2:]
    if not re.fullmatch(r"[0-9A-F]{6}", value):
        raise ValueError("Color must be RGB or ARGB hex, e.g. 'FF0000' or 'FFFF0000'.")
    return value


_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
ET.register_namespace("a", _A_NS)
ET.register_namespace("xdr", _XDR_NS)
_SHAPE_CLEAR = "__DOCFORGE_CLEAR__"
_SHAPE_KEEP = "__DOCFORGE_KEEP__"


def _solid_fill_xml(rgb: str) -> str:
    return f'<a:solidFill><a:srgbClr val="{rgb}"/></a:solidFill>'


def _et_local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _et_tostring(elem: ET.Element) -> str:
    return ET.tostring(elem, encoding="unicode", short_empty_elements=True)


def _parse_sp_pr(sp_pr: str) -> ET.Element:
    wrapper = ET.fromstring(
        f'<docforge:root xmlns:docforge="urn:docforge" xmlns:xdr="{_XDR_NS}" xmlns:a="{_A_NS}">{sp_pr}</docforge:root>'
    )
    return list(wrapper)[0]


def _first_child_index(elem: ET.Element, names: set[str]) -> int | None:
    for idx, child in enumerate(list(elem)):
        if _et_local(child.tag) in names:
            return idx
    return None


def _new_solid_fill(rgb: str) -> ET.Element:
    fill = ET.Element(f"{{{_A_NS}}}solidFill")
    ET.SubElement(fill, f"{{{_A_NS}}}srgbClr", {"val": rgb})
    return fill


def _new_no_fill() -> ET.Element:
    return ET.Element(f"{{{_A_NS}}}noFill")


def _replace_direct_fill(elem: ET.Element, replacement: ET.Element) -> None:
    fill_names = {"noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"}
    for idx, child in enumerate(list(elem)):
        if _et_local(child.tag) in fill_names:
            elem.remove(child)
            elem.insert(idx, replacement)
            return
    geom_idx = _first_child_index(elem, {"prstGeom", "custGeom"})
    insert_at = geom_idx + 1 if geom_idx is not None else 0
    elem.insert(insert_at, replacement)


def _set_shape_fill(sp_pr: str, rgb: str | None) -> str:
    root = _parse_sp_pr(sp_pr)
    _replace_direct_fill(root, _new_no_fill() if rgb is None else _new_solid_fill(rgb))
    return _et_tostring(root)


def _set_shape_line(sp_pr: str, rgb_marker, width_pt: float | None) -> str:
    root = _parse_sp_pr(sp_pr)
    if width_pt is not None and width_pt < 0:
        raise ValueError("outline_width_pt must be >= 0.")
    line = next((c for c in list(root) if _et_local(c.tag) == "ln"), None)
    if line is None:
        line = ET.Element(f"{{{_A_NS}}}ln")
        root.append(line)
    if width_pt is not None:
        line.set("w", str(int(round(width_pt * 12700))))
    if rgb_marker is _SHAPE_CLEAR:
        _replace_direct_fill(line, _new_no_fill())
    elif rgb_marker is not _SHAPE_KEEP:
        _replace_direct_fill(line, _new_solid_fill(rgb_marker))
    return _et_tostring(root)


def _set_shape_text_color(anchor_xml: str, rgb: str) -> str:
    def patch_rpr(match: re.Match) -> str:
        tag = match.group(0)
        if tag.endswith("/>"):
            tag = tag[:-2] + ">" + _solid_fill_xml(rgb) + "</a:rPr>"
        elif re.search(r"<a:solidFill\b.*?</a:solidFill>|<a:noFill\s*/>", tag, re.DOTALL):
            tag = re.sub(r"<a:solidFill\b.*?</a:solidFill>|<a:noFill\s*/>", _solid_fill_xml(rgb), tag, count=1, flags=re.DOTALL)
        else:
            tag = tag.replace("</a:rPr>", _solid_fill_xml(rgb) + "</a:rPr>", 1)
        return tag
    updated, count = re.subn(r"<a:rPr\b.*?</a:rPr>|<a:rPr\b[^/]*/>", patch_rpr, anchor_xml, flags=re.DOTALL)
    return updated if count else anchor_xml

@mcp.tool()
def excel_get_shapes(session_key: str, sheet_name: str | None = None) -> str:
    """List DrawingML shapes/images/charts captured from loaded sheets."""
    data = _get_session(session_key)
    result = {}
    for sheet in data["sheets"]:
        if sheet_name and sheet["name"] != sheet_name:
            continue
        result[sheet["name"]] = sheet.get("shapes") or []
    if sheet_name and sheet_name not in result:
        _find_sheet(data, sheet_name)
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool()
def excel_update_shape_text(session_key: str, sheet_name: str, shape_index: int, text: str) -> str:
    """Update text in a DrawingML shape/textbox by 1-based shape index."""
    from html import escape

    data = _get_session(session_key)
    sheet = _find_sheet(data, sheet_name)
    drawing = sheet.get("drawing_data")
    if not drawing or not drawing.get("drawing_xml"):
        raise ValueError(f"Sheet '{sheet_name}' has no captured DrawingML shapes.")
    if shape_index < 1:
        raise ValueError("shape_index is 1-based and must be >= 1.")

    anchors = list(re.finditer(
        r"<(?:(?:xdr:)?)(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b.*?</(?:(?:xdr:)?)(?:twoCellAnchor|oneCellAnchor|absoluteAnchor)>",
        drawing["drawing_xml"],
        re.DOTALL,
    ))
    if shape_index > len(anchors):
        raise ValueError(f"shape_index {shape_index} out of bounds; found {len(anchors)} shape(s).")

    match = anchors[shape_index - 1]
    anchor_xml = match.group(0)
    text_matches = list(re.finditer(r"<a:t(?:\s[^>]*)?>.*?</a:t>", anchor_xml, re.DOTALL))
    if not text_matches:
        raise ValueError(f"Shape {shape_index} has no editable text runs.")

    escaped = escape(text, quote=False)
    first = True
    def replace_text_run(m: re.Match) -> str:
        nonlocal first
        start_tag = re.match(r"<a:t(?:\s[^>]*)?>", m.group(0)).group(0)
        value = escaped if first else ""
        first = False
        return f"{start_tag}{value}</a:t>"

    new_anchor = re.sub(r"<a:t(?:\s[^>]*)?>.*?</a:t>", replace_text_run, anchor_xml, flags=re.DOTALL)
    drawing["drawing_xml"] = drawing["drawing_xml"][:match.start()] + new_anchor + drawing["drawing_xml"][match.end():]
    shapes = sheet.get("shapes") or []
    if shape_index <= len(shapes):
        shapes[shape_index - 1]["text"] = text
    return f"Updated text for shape {shape_index} on sheet '{sheet_name}'."



@mcp.tool()
def excel_set_shape_style(
    session_key: str,
    sheet_name: str,
    shape_index: int,
    fill_color: str | None = None,
    outline_color: str | None = None,
    outline_width_pt: float | None = None,
    text_color: str | None = None,
    clear_fill: bool = False,
    clear_outline: bool = False,
) -> str:
    """
    Set simple DrawingML shape style by 1-based shape index.

    Colors accept RGB or ARGB hex. Use clear_fill=True or clear_outline=True to remove fill/outline.
    Supports simple DrawingML shapes/textboxes captured by excel_load.
    """
    data = _get_session(session_key)
    sheet = _find_sheet(data, sheet_name)
    drawing = sheet.get("drawing_data")
    if not drawing or not drawing.get("drawing_xml"):
        raise ValueError(f"Sheet '{sheet_name}' has no captured DrawingML shapes.")
    if shape_index < 1:
        raise ValueError("shape_index is 1-based and must be >= 1.")

    anchors = _shape_anchors(drawing["drawing_xml"])
    if shape_index > len(anchors):
        raise ValueError(f"shape_index {shape_index} out of bounds; found {len(anchors)} shape(s).")

    if clear_fill and fill_color is not None:
        raise ValueError("Use either fill_color or clear_fill, not both.")
    if clear_outline and outline_color is not None:
        raise ValueError("Use either outline_color or clear_outline, not both.")
    fill_rgb = _rgb_hex(fill_color) if fill_color is not None else None
    outline_rgb = _rgb_hex(outline_color) if outline_color is not None else None
    text_rgb = _rgb_hex(text_color) if text_color is not None else None

    match = anchors[shape_index - 1]
    anchor_xml = match.group(0)
    sp_match = re.search(r"<(?:(?:xdr:)?sp)\b.*?</(?:(?:xdr:)?sp)>", anchor_xml, re.DOTALL)
    if not sp_match:
        raise ValueError(f"Shape {shape_index} is not a simple editable DrawingML shape.")
    shape_xml = sp_match.group(0)
    sp_pr_match = re.search(r"<(?:(?:xdr:)?spPr)\b.*?</(?:(?:xdr:)?spPr)>", shape_xml, re.DOTALL)
    if not sp_pr_match:
        raise ValueError(f"Shape {shape_index} has no editable shape properties.")

    sp_pr = sp_pr_match.group(0)
    if fill_color is not None or clear_fill:
        sp_pr = _set_shape_fill(sp_pr, None if clear_fill else fill_rgb)
    if outline_color is not None or outline_width_pt is not None or clear_outline:
        outline_marker = _SHAPE_CLEAR if clear_outline else (outline_rgb if outline_color is not None else _SHAPE_KEEP)
        sp_pr = _set_shape_line(sp_pr, outline_marker, outline_width_pt)
    shape_xml = shape_xml[:sp_pr_match.start()] + sp_pr + shape_xml[sp_pr_match.end():]
    anchor_xml = anchor_xml[:sp_match.start()] + shape_xml + anchor_xml[sp_match.end():]
    if text_rgb is not None:
        anchor_xml = _set_shape_text_color(anchor_xml, text_rgb)

    drawing["drawing_xml"] = drawing["drawing_xml"][:match.start()] + anchor_xml + drawing["drawing_xml"][match.end():]
    shapes = sheet.get("shapes") or []
    if shape_index <= len(shapes):
        if fill_color is not None or clear_fill:
            shapes[shape_index - 1]["fill_color"] = None if clear_fill else fill_rgb
        if outline_color is not None or clear_outline:
            shapes[shape_index - 1]["outline_color"] = None if clear_outline else outline_rgb
        if outline_width_pt is not None:
            shapes[shape_index - 1]["outline_width_pt"] = outline_width_pt
        if text_color is not None:
            shapes[shape_index - 1]["text_color"] = text_rgb
    return f"Updated style for shape {shape_index} on sheet '{sheet_name}'."
# ── 20. Borders ───────────────────────────────────────────────────────────────

@mcp.tool()
def excel_set_borders(
    session_key: str,
    sheet_name: str,
    r1: int,
    c1: int,
    r2: int,
    c2: int,
    style: str,
    sides: list[str] | None = None,
    color: str | None = None,
) -> str:
    """
    Set or remove borders on a rectangular cell range.

    All coordinates are 0-based, inclusive. Slave cells are skipped.

    Border styles: "thin" | "medium" | "thick" | "dashed" | "dotted" |
                   "hair" | "mediumDashed" | "dashDot" | "none" (removes)

    Args:
        session_key: Key returned by excel_load
        sheet_name: Name of the sheet
        r1: 0-based top row (inclusive)
        c1: 0-based left column (inclusive)
        r2: 0-based bottom row (inclusive)
        c2: 0-based right column (inclusive)
        style: Border style, or "none" to remove borders
        sides: ["top","bottom","left","right"] subset; omit for all four
        color: Border color as ARGB hex (e.g. "FF000000"); omit for black

    Returns:
        Summary: cells updated
    """
    _VALID_SIDES = {"top", "bottom", "left", "right"}
    if sides is None:
        sides = list(_VALID_SIDES)
    else:
        bad = set(sides) - _VALID_SIDES
        if bad:
            raise ValueError(f"Invalid sides: {sorted(bad)}. Use top/bottom/left/right.")

    side_data: dict | None = None if style == "none" else {"style": style, **({"color": color} if color else {})}

    data = _get_session(session_key)
    sheet = _find_sheet(data, sheet_name)
    rows = sheet["rows"]

    cells_updated = 0
    for r in range(r1, r2 + 1):
        if r >= len(rows):
            continue
        row_cells = rows[r]["cells"]
        for c in range(c1, c2 + 1):
            if c >= len(row_cells):
                continue
            cell = row_cells[c]
            if cell.get("merge") == "slave":
                continue
            bdr = cell.get("border") or {}
            for side in sides:
                if side_data is None:
                    bdr.pop(side, None)
                else:
                    bdr[side] = side_data
            cell["border"] = bdr
            cells_updated += 1

    action = "Removed" if style == "none" else f"Set '{style}'"
    return f"{action} borders ({', '.join(sorted(sides))}) on {cells_updated} cell(s) in [{r1},{c1}]–[{r2},{c2}]."


# ── 21. Dimensions ────────────────────────────────────────────────────────────

@mcp.tool()
def excel_set_dimension(
    session_key: str,
    sheet_name: str,
    axis: str,
    index: int,
    size: float | None,
) -> str:
    """
    Set the height of a row or the width of a column.

    Args:
        session_key: Key returned by excel_load
        sheet_name: Name of the sheet
        axis: "row" to set row height, "col" to set column width
        index: 0-based row or column index
        size: Height in points (rows) or width in character units (cols).
              Pass null to reset to auto.

    Returns:
        Confirmation
    """
    import openpyxl.utils
    data = _get_session(session_key)
    sheet = _find_sheet(data, sheet_name)

    if axis == "row":
        rows = sheet["rows"]
        if not (0 <= index < len(rows)):
            raise ValueError(f"row index {index} out of range (0–{len(rows)-1})")
        rows[index]["h"] = size
        return f"Set row {index} height to {size!r} in sheet '{sheet_name}'."
    elif axis == "col":
        col_letter = openpyxl.utils.get_column_letter(index + 1)
        sheet["cw"][col_letter] = size
        return f"Set column {index} ({col_letter}) width to {size!r} in sheet '{sheet_name}'."
    else:
        raise ValueError(f"axis must be 'row' or 'col', got {axis!r}")


@mcp.tool()
def excel_set_row_height(
    session_key: str,
    sheet_name: str,
    row_heights: dict,
) -> str:
    """
    Set height for one or more rows in a single call.

    Args:
        session_key: Key returned by excel_load
        sheet_name: Name of the sheet
        row_heights: Map of {row_index: height}. Height in points.
                     Pass null to reset a row to auto height.
                     Example: {"0": 30, "1": 20, "5": null}

    Returns:
        Confirmation
    """
    data = _get_session(session_key)
    sheet = _find_sheet(data, sheet_name)
    rows = sheet["rows"]
    updated = 0
    for idx_str, height in row_heights.items():
        idx = int(idx_str)
        if not (0 <= idx < len(rows)):
            raise ValueError(f"row index {idx} out of range (0–{len(rows)-1})")
        rows[idx]["h"] = height
        updated += 1
    return f"Set height for {updated} row(s) in sheet '{sheet_name}'."


@mcp.tool()
def excel_set_column_width(
    session_key: str,
    sheet_name: str,
    col_widths: dict,
) -> str:
    """
    Set width for one or more columns in a single call.

    Args:
        session_key: Key returned by excel_load
        sheet_name: Name of the sheet
        col_widths: Map of {col: width}. Width in character units.
                    Column key accepts letter ("A", "B") or 0-based integer ("0", "1").
                    Pass null to remove an explicit width (resets to default).
                    Example: {"A": 20, "B": 15, "2": 30}

    Returns:
        Confirmation
    """
    import openpyxl.utils
    data = _get_session(session_key)
    sheet = _find_sheet(data, sheet_name)
    updated = 0
    for key, width in col_widths.items():
        try:
            letter = openpyxl.utils.get_column_letter(int(key) + 1)
        except ValueError:
            letter = key.upper()
        if width is None:
            sheet["cw"].pop(letter, None)
        else:
            sheet["cw"][letter] = float(width)
        updated += 1
    return f"Set width for {updated} column(s) in sheet '{sheet_name}'."


@mcp.tool()
def excel_autofit_cols(
    session_key: str,
    sheet_name: str,
    col_indices: list[int] | None = None,
    min_width: float = 8.0,
    max_width: float = 60.0,
) -> str:
    """
    Estimate and set column widths based on content length (heuristic approximation).

    openpyxl cannot measure rendered text, so widths are estimated from string
    length, font size, and bold flag. Results are usually close but may need
    manual adjustment.

    Args:
        session_key: Key returned by excel_load
        sheet_name: Name of the sheet
        col_indices: 0-based column indices to fit; omit for all columns
        min_width: Minimum column width (default 8.0)
        max_width: Maximum column width cap (default 60.0)

    Returns:
        JSON: {columns_fitted, widths: {col_index: width}}
    """
    import openpyxl.utils
    data = _get_session(session_key)
    sheet = _find_sheet(data, sheet_name)
    rows = sheet["rows"]

    n_cols = max((len(r["cells"]) for r in rows), default=0)
    targets = col_indices if col_indices is not None else list(range(n_cols))

    updated: dict[int, float] = {}
    for c in targets:
        max_len = 0.0
        for row in rows:
            if c < len(row["cells"]):
                cell = row["cells"][c]
                if cell.get("merge") != "slave" and cell["v"] is not None:
                    text_len = len(str(cell["v"]))
                    size = cell.get("size") or 11
                    factor = 1.2 if cell.get("bold") else 1.0
                    est = text_len * (size / 11) * factor
                    if est > max_len:
                        max_len = est
        width = round(max(min_width, min(max_width, max_len * 1.1 + 2)), 1)
        sheet["cw"][openpyxl.utils.get_column_letter(c + 1)] = width
        updated[c] = width

    return json.dumps({"sheet": sheet_name, "columns_fitted": len(updated),
                       "widths": {str(k): v for k, v in updated.items()}}, ensure_ascii=False)


# ── 22. Freeze panes ──────────────────────────────────────────────────────────

@mcp.tool()
def excel_freeze_panes(
    session_key: str,
    sheet_name: str,
    row: int,
    col: int,
) -> str:
    """
    Freeze rows above `row` and/or columns to the left of `col`.

    row=1, col=0 → freeze first row only (most common for headers)
    row=0, col=1 → freeze first column only
    row=1, col=1 → freeze both header row and first column
    row=0, col=0 → unfreeze

    Args:
        session_key: Key returned by excel_load
        sheet_name: Name of the sheet
        row: First unfrozen row (0-based); 0 = no row freeze
        col: First unfrozen column (0-based); 0 = no column freeze

    Returns:
        Confirmation with freeze reference cell
    """
    import openpyxl.utils
    data = _get_session(session_key)
    sheet = _find_sheet(data, sheet_name)
    if row == 0 and col == 0:
        sheet["freeze"] = None
        return f"Unfrozen panes in sheet '{sheet_name}'."
    col_letter = openpyxl.utils.get_column_letter(col + 1)
    ref = f"{col_letter}{row + 1}"
    sheet["freeze"] = ref
    frozen = []
    if row > 0:
        frozen.append(f"rows 0–{row-1}")
    if col > 0:
        frozen.append(f"cols 0–{col-1}")
    return f"Freeze → {ref!r} in sheet '{sheet_name}' ({', '.join(frozen)} frozen)."


# ── 23. Data validation ───────────────────────────────────────────────────────

@mcp.tool()
def excel_set_data_validation(
    session_key: str,
    sheet_name: str,
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
    options: list[str],
    allow_blank: bool = True,
) -> str:
    """
    Add a dropdown list validation to a cell range.

    All coordinates are 0-based, inclusive on both ends.

    Args:
        session_key: Key returned by excel_load
        sheet_name: Name of the sheet
        start_row: 0-based top row (inclusive)
        start_col: 0-based left column (inclusive)
        end_row: 0-based bottom row (inclusive)
        end_col: 0-based right column (inclusive)
        options: Allowed values shown in the dropdown, e.g. ["YES", "NO"]
        allow_blank: If True (default), empty cells are valid

    Returns:
        Confirmation with cell range and options
    """
    import openpyxl.utils
    data = _get_session(session_key)
    sheet = _find_sheet(data, sheet_name)

    c1_letter = openpyxl.utils.get_column_letter(start_col + 1)
    c2_letter = openpyxl.utils.get_column_letter(end_col + 1)
    sqref = f"{c1_letter}{start_row + 1}:{c2_letter}{end_row + 1}"
    formula1 = '"' + ",".join(str(o).replace('"', "") for o in options) + '"'

    if "validations" not in sheet:
        sheet["validations"] = []
    sheet.pop("data_validations_xml", None)
    sheet["validations"].append({
        "type": "list", "formula1": formula1,
        "formula2": None, "allow_blank": allow_blank, "sqref": sqref,
    })
    return f"Added dropdown {options} to {sqref} in sheet '{sheet_name}'."


# ── 24. Find rows / Fill column ───────────────────────────────────────────────

@mcp.tool()
def excel_find_rows(
    session_key: str,
    sheet_name: str,
    col_index: int,
    value: str | None = None,
    pattern: str | None = None,
    start_row: int = 0,
    end_row: int | None = None,
) -> str:
    """
    Find all rows where a column cell matches a value or regex pattern.

    value and pattern are mutually exclusive — provide exactly one.

    Args:
        session_key: Key returned by excel_load
        sheet_name: Name of the sheet
        col_index: 0-based column index to search in
        value: Exact match (compared as string)
        pattern: Python regex pattern (re.search)
        start_row: 0-based start row (inclusive), default 0
        end_row: 0-based end row (exclusive); omit for all rows

    Returns:
        JSON array of {row_index, values} for each matching row
    """
    import re
    if value is None and pattern is None:
        raise ValueError("Provide either value or pattern.")
    if value is not None and pattern is not None:
        raise ValueError("value and pattern are mutually exclusive.")

    data = _get_session(session_key)
    sheet = _find_sheet(data, sheet_name)
    rows = sheet["rows"]

    results = []
    for r_idx, row in enumerate(rows[start_row:end_row], start=start_row):
        cells = row["cells"]
        if col_index >= len(cells):
            continue
        cell_val = cells[col_index]["v"]
        cell_str = str(cell_val) if cell_val is not None else ""
        matched = (cell_str == str(value)) if value is not None else bool(re.search(pattern, cell_str))
        if matched:
            results.append({
                "row_index": r_idx,
                "values": [cd["v"] if cd.get("merge") != "slave" else None for cd in cells],
            })
    return json.dumps(results, default=str, ensure_ascii=False)


@mcp.tool()
def excel_fill_column(
    session_key: str,
    sheet_name: str,
    col_index: int,
    start_row: int,
    end_row: int,
    value: str | int | float | None = None,
    sequence_start: int | None = None,
    step: int = 1,
) -> str:
    """
    Fill a column range with a constant value or an auto-incrementing sequence.

    Constant fill: pass value — every cell gets the same value.
    Sequence fill: pass sequence_start — cells get sequence_start,
                   sequence_start+step, sequence_start+2×step, …

    start_row and end_row are both inclusive (0-based).
    Slave cells of merged regions are skipped.

    Args:
        session_key: Key returned by excel_load
        sheet_name: Name of the sheet
        col_index: 0-based column index to fill
        start_row: 0-based start row (inclusive)
        end_row: 0-based end row (inclusive)
        value: Constant fill value (mutually exclusive with sequence_start)
        sequence_start: First integer of the sequence
        step: Increment between sequence values (default 1)

    Returns:
        Summary: cells filled
    """
    if value is not None and sequence_start is not None:
        raise ValueError("value and sequence_start are mutually exclusive.")

    data = _get_session(session_key)
    sheet = _find_sheet(data, sheet_name)
    rows = sheet["rows"]
    n = len(rows)
    if not (0 <= start_row <= end_row < n):
        raise ValueError(f"Row range [{start_row}, {end_row}] out of bounds (0–{n-1})")

    cells_filled = 0
    seq = sequence_start
    for r in range(start_row, end_row + 1):
        row_cells = rows[r]["cells"]
        if col_index >= len(row_cells):
            continue
        cell = row_cells[col_index]
        if cell.get("merge") == "slave":
            continue
        if sequence_start is not None:
            _store_cell_value(cell, seq)
            seq += step
        else:
            _store_cell_value(cell, value)
        cells_filled += 1

    return f"Filled {cells_filled} cell(s) in col {col_index}, rows {start_row}–{end_row} in sheet '{sheet_name}'."


# ── 25. Fill rows (stamp pattern) ─────────────────────────────────────────────

@mcp.tool()
def excel_fill_rows(
    session_key: str,
    sheet_name: str,
    template_row: int,
    after_index: int,
    count: int,
) -> str:
    """
    Clone a template row N times and insert all copies in one call.

    More efficient than clone_rows + insert_rows when inserting many rows with
    the same format (e.g. stamping a formatted template row for a data table).
    All cell values, styles, and borders from the template are preserved in every copy.

    Args:
        session_key: Key returned by excel_load
        sheet_name: Name of the sheet
        template_row: 0-based index of the row to clone as the template
        after_index: Insert the block AFTER this 0-based row index; use -1 to prepend
        count: Number of copies to insert (must be > 0)

    Returns:
        Confirmation with new row count
    """
    data = _get_session(session_key)
    sheet = _find_sheet(data, sheet_name)
    n = len(sheet["rows"])
    if not (0 <= template_row < n):
        raise ValueError(f"template_row {template_row} out of range (0–{n - 1})")
    if count <= 0:
        raise ValueError(f"count must be > 0, got {count}")

    template = sheet["rows"][template_row]
    new_rows = [copy.deepcopy(template) for _ in range(count)]
    pos = after_index + 1
    _apply_row_insert(data, sheet, pos, new_rows)
    return (
        f"Inserted {count} copy/copies of row {template_row} after index {after_index}. "
        f"Sheet '{sheet_name}' now has {len(sheet['rows'])} rows."
    )


if __name__ == "__main__":
    mcp.run(transport="stdio")








