"""md-tools — MCP server for Markdown files."""
import re
import sys
import unicodedata
import csv
import io
import json
import os
import shutil
import tempfile
import threading
from html import escape
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlparse

_SERVER_DIRECTORY = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(_SERVER_DIRECTORY))

from mcp.server.fastmcp import FastMCP
from process_lifecycle import (
    ManagedProcessCancelled,
    run_cancellable_in_thread,
    run_managed_process,
)

MAX_MARKDOWN_BYTES = 10 * 1024 * 1024

mcp = FastMCP(
    "md-tools",
    instructions=(
        "Local-first Markdown structure editor. Start with markdown_outline to inspect "
        "headings, line numbers, slugs, and section spans; use md_get_document_map for "
        "a compact full document map; use md_search or md_read_near when the target "
        "text is unknown; use read_markdown_section or md_read_range for focused reads. "
        "Use preview/max_chars/include_body on large reads to reduce token output. Prefer heading_path "
        "when duplicate titles exist. For section edits use replace_markdown_section, "
        "md_insert_section, md_delete_section, md_append_to_section, md_patch_lines, "
        "md_move_section, md_set_heading_level, md_normalize_headings, or md_rename_heading, "
        "then re-read the affected section. "
        "Use table, diagram, code-block, link/image, TOC, frontmatter, split/merge, and "
        "stats tools for Markdown-native operations. Tools that return Markdown content "
        "return raw Markdown text for direct rendering; whole-file arbitrary reads belong "
        "to text-tools. Diagram validation/rendering is best-effort and reports skipped "
        "when optional external CLIs are unavailable."
    ),
)


@dataclass
class MarkdownHeading:
    title: str
    level: int
    line: int
    path: list[str]
    slug: str


def _resolve_path(path: str) -> Path:
    return Path(path).expanduser().resolve()


def _decode_fuzzy(data: bytes) -> tuple[str, str]:
    if data.startswith(b"\xef\xbb\xbf"):
        return data[3:].decode("utf-8", errors="replace"), "UTF-8"
    if data.startswith(b"\xff\xfe"):
        return data[2:].decode("utf-16-le", errors="replace"), "UTF-16LE"
    if data.startswith(b"\xfe\xff"):
        return data[2:].decode("utf-16-be", errors="replace"), "UTF-16BE"
    try:
        return data.decode("utf-8"), "UTF-8"
    except UnicodeDecodeError:
        return data.decode("cp1252", errors="replace"), "WINDOWS-1252"


def _encode_fuzzy(content: str, encoding: str) -> bytes:
    if encoding == "WINDOWS-1252":
        return content.encode("cp1252")
    if encoding == "UTF-16LE":
        return b"\xff\xfe" + content.encode("utf-16-le")
    if encoding == "UTF-16BE":
        return b"\xfe\xff" + content.encode("utf-16-be")
    return content.encode("utf-8")


def _to_lines(content: str) -> list[str]:
    """Split content into logical lines on ``\\n`` only.

    A single trailing newline does not create an extra empty line. This is the
    one line model used everywhere — heading line numbers, section slicing, and
    byte offsets — so line numbers stay consistent regardless of CR/LF style or
    other characters that ``str.splitlines`` would otherwise treat as breaks.
    """
    if content == "":
        return []
    lines = content.split("\n")
    if content.endswith("\n"):
        lines.pop()
    return lines


def _load_markdown_file(path_value: str) -> tuple[Path, str, str, int, int]:
    path = _resolve_path(path_value)
    if not path.exists() or not path.is_file():
        raise ValueError(f"Markdown file does not exist or is not a file: {path}")
    if path.suffix.lower() not in {".md", ".markdown"}:
        raise ValueError(f"Expected a Markdown file (.md or .markdown): {path}")
    size = path.stat().st_size
    if size > MAX_MARKDOWN_BYTES:
        raise ValueError(f"Markdown file is too large ({size} bytes > {MAX_MARKDOWN_BYTES} bytes)")
    data = path.read_bytes()
    content, encoding = _decode_fuzzy(data)
    total_lines = len(_to_lines(content))
    return path, content, encoding, total_lines, size


def _write_markdown_file(path: Path, content: str, encoding: str) -> dict[str, Any]:
    before = path.read_bytes()
    after = _encode_fuzzy(content, encoding)
    changed = before != after
    if changed:
        path.write_bytes(after)
    return {"changed": changed, "bytes_before": len(before), "bytes_written": len(after)}


def _slugify_heading(title: str) -> str:
    """Return a GitHub-compatible anchor slug for a heading title.

    Lowercases, drops punctuation and symbols, but keeps Unicode letters,
    numbers, and combining marks so non-ASCII headings (e.g. Vietnamese) slug
    the same way GitHub renders them. Matches github-slugger semantics: spaces
    become hyphens, and existing hyphens are neither collapsed nor stripped.
    """
    slug: list[str] = []
    for char in title.strip().lower():
        if char == " ":
            slug.append("-")
        elif char in "-_":
            slug.append(char)
        elif unicodedata.category(char)[0] in {"L", "N", "M"}:
            slug.append(char)
    return "".join(slug)


def _parse_headings(content: str) -> list[MarkdownHeading]:
    heading_re = re.compile(r"^(#{1,6})[ \t]+(.+?)(?:[ \t]+#+[ \t]*)?$")
    fence_re = re.compile(r"^([`~]{3,})")
    headings: list[MarkdownHeading] = []
    stack: list[tuple[int, str]] = []
    active_fence: tuple[str, int] | None = None

    for index, raw_line in enumerate(_to_lines(content), start=1):
        line = raw_line.rstrip("\r")
        leading_spaces = len(line) - len(line.lstrip(" "))
        if leading_spaces >= 4 or line.startswith("\t"):
            continue
        trimmed = line.lstrip()

        fence_match = fence_re.match(trimmed)
        if fence_match:
            marker = fence_match.group(1)[0]
            length = len(fence_match.group(1))
            if active_fence and active_fence[0] == marker and length >= active_fence[1]:
                active_fence = None
            elif active_fence is None:
                active_fence = (marker, length)
            continue

        if active_fence is not None:
            continue

        heading_match = heading_re.match(trimmed)
        if not heading_match:
            continue

        level = len(heading_match.group(1))
        title = heading_match.group(2).strip()
        while stack and stack[-1][0] >= level:
            stack.pop()
        stack.append((level, title))
        path = [item_title for _, item_title in stack]
        headings.append(MarkdownHeading(title=title, level=level, line=index, path=path, slug=_slugify_heading(title)))

    return headings


def _section_end_line(headings: list[MarkdownHeading], heading_index: int, total_lines: int, include_subsections: bool) -> int:
    current = headings[heading_index]
    for next_heading in headings[heading_index + 1 :]:
        if include_subsections:
            if next_heading.level <= current.level:
                return max(next_heading.line - 1, 0)
        else:
            return max(next_heading.line - 1, 0)
    return total_lines


def _heading_matches(candidate: MarkdownHeading, heading: str | None, heading_path: list[str] | None, exact: bool) -> bool:
    if heading_path is not None:
        if len(candidate.path) != len(heading_path):
            return False
        if exact:
            return candidate.path == heading_path
        return all(left.casefold() == right.casefold() for left, right in zip(candidate.path, heading_path))
    if heading is None:
        return False
    return candidate.title == heading if exact else candidate.title.casefold() == heading.casefold()


def _match_heading(headings: list[MarkdownHeading], heading: str | None, heading_path: list[str] | None, exact: bool) -> int:
    matches = [index for index, candidate in enumerate(headings) if _heading_matches(candidate, heading, heading_path, exact)]
    if len(matches) == 1:
        return matches[0]
    if not matches:
        raise ValueError("Requested markdown heading was not found")
    sample = "; ".join(f"{' > '.join(headings[index].path)} (line {headings[index].line})" for index in matches[:5])
    raise ValueError(f"Requested markdown heading is ambiguous. Matching sections: {sample}. Use heading_path to disambiguate.")


def _line_start_offsets(content: str) -> list[int]:
    offsets = [0]
    for index, char in enumerate(content):
        if char == "\n":
            offsets.append(index + 1)
    return offsets


def _line_start_offset(offsets: list[int], line: int, content_len: int) -> int:
    if line <= 0:
        return 0
    return offsets[line - 1] if line - 1 < len(offsets) else content_len


def _slice_lines(content: str, start_line: int, end_line: int, include_line_numbers: bool, max_lines: int | None, max_bytes: int | None) -> str:
    lines = _to_lines(content)
    selected = lines[max(start_line - 1, 0) : min(end_line, len(lines))]
    rendered: list[str] = []
    used_bytes = 0
    line_limit = max_lines if max_lines is not None else len(selected)
    byte_limit = max_bytes if max_bytes is not None else sys.maxsize
    for offset, raw_line in enumerate(selected):
        if len(rendered) >= line_limit:
            break
        line = raw_line.rstrip("\r")
        line_number = start_line + offset
        value = f"{line_number:>6} | {line}" if include_line_numbers else line
        separator = 1 if rendered else 0
        value_bytes = len(value.encode("utf-8"))
        if used_bytes + separator + value_bytes > byte_limit:
            break
        used_bytes += separator + value_bytes
        rendered.append(value)
    return "\n".join(rendered)


def _heading_metadata(heading: MarkdownHeading) -> dict[str, Any]:
    return {
        "title": heading.title,
        "level": heading.level,
        "line": heading.line,
        "slug": heading.slug,
        "path": heading.path,
        "path_text": " > ".join(heading.path),
    }


def _heading_for_line(headings: list[MarkdownHeading], line: int) -> dict[str, Any] | None:
    current = None
    for heading in headings:
        if heading.line > line:
            break
        current = heading
    return _heading_metadata(current) if current else None


def _format_heading_line(level: int, title: str) -> str:
    if not 1 <= level <= 6:
        raise ValueError("heading level must be between 1 and 6")
    return f"{'#' * level} {title.strip()}"


def _replace_line_range(content: str, start_line: int, end_line: int, replacement: str) -> str:
    offsets = _line_start_offsets(content)
    start_byte = _line_start_offset(offsets, start_line, len(content))
    end_byte = _line_start_offset(offsets, end_line + 1, len(content)) if start_line <= end_line else start_byte
    return content[:start_byte] + replacement + content[end_byte:]


def _section_bounds(content: str, heading: str | None, heading_path: list[str] | None, exact: bool, include_subsections: bool = True) -> tuple[list[MarkdownHeading], int, MarkdownHeading, int]:
    if heading is None and heading_path is None:
        raise ValueError("Either heading or heading_path is required")
    total_lines = len(_to_lines(content))
    headings = _parse_headings(content)
    heading_index = _match_heading(headings, heading, heading_path, exact)
    selected = headings[heading_index]
    return headings, heading_index, selected, _section_end_line(headings, heading_index, total_lines, include_subsections)


def _normalize_block(content: str, trailing_newline: bool = True) -> str:
    value = content.strip("\n")
    return value + "\n" if value and trailing_newline else value


def _parse_pipe_row(line: str) -> list[str]:
    stripped = line.strip().strip("|")
    return [cell.strip() for cell in stripped.split("|")]


def _is_table_separator(line: str) -> bool:
    cells = _parse_pipe_row(line)
    return bool(cells) and all(re.fullmatch(r":?-{3,}:?", cell.replace(" ", "")) for cell in cells)


def _find_tables(content: str) -> list[dict[str, Any]]:
    lines = _to_lines(content)
    headings = _parse_headings(content)
    tables = []
    index = 0
    while index < len(lines) - 1:
        if "|" not in lines[index] or not _is_table_separator(lines[index + 1]):
            index += 1
            continue
        start = index + 1
        end_index = index + 1
        while end_index + 1 < len(lines) and "|" in lines[end_index + 1].strip() and lines[end_index + 1].strip():
            end_index += 1
        header = _parse_pipe_row(lines[index])
        rows = [_parse_pipe_row(line) for line in lines[index + 2 : end_index + 1]]
        tables.append({
            "index": len(tables),
            "start_line": start,
            "end_line": end_index + 1,
            "heading": _heading_for_line(headings, start),
            "rows": len(rows),
            "columns": len(header),
            "headers": header,
            "raw_lines": lines[index : end_index + 1],
        })
        index = end_index + 1
    return tables


def _table_alignments(separator: str, columns: int) -> list[str]:
    aligns = []
    for cell in _parse_pipe_row(separator):
        compact = cell.replace(" ", "")
        if compact.startswith(":") and compact.endswith(":"):
            aligns.append("center")
        elif compact.endswith(":"):
            aligns.append("right")
        elif compact.startswith(":"):
            aligns.append("left")
        else:
            aligns.append("left")
    return (aligns + ["left"] * columns)[:columns]


def _render_table(headers: list[str], rows: list[list[str]], alignments: list[str] | None = None) -> str:
    columns = len(headers)
    normalized_rows = [(row + [""] * columns)[:columns] for row in rows]
    widths = [max(len(str(headers[col])), *(len(str(row[col])) for row in normalized_rows), 3) for col in range(columns)]
    alignments = (alignments or ["left"] * columns + ["left"] * columns)[:columns]

    def pad(value: str, col: int) -> str:
        width = widths[col]
        if alignments[col] == "right":
            return value.rjust(width)
        if alignments[col] == "center":
            return value.center(width)
        return value.ljust(width)

    separator_cells = []
    for col, align in enumerate(alignments):
        dashes = "-" * max(widths[col], 3)
        if align == "center":
            separator_cells.append(f":{dashes}:")
        elif align == "right":
            separator_cells.append(f"{dashes}:")
        elif align == "left":
            separator_cells.append(f":{dashes}")
        else:
            separator_cells.append(dashes)
    rendered = ["| " + " | ".join(pad(str(value), col) for col, value in enumerate(headers)) + " |"]
    rendered.append("| " + " | ".join(separator_cells) + " |")
    rendered.extend("| " + " | ".join(pad(str(value), col) for col, value in enumerate(row)) + " |" for row in normalized_rows)
    return "\n".join(rendered) + "\n"


def _select_table(content: str, table_index: int = 0, heading: str | None = None, heading_path: list[str] | None = None, exact: bool = False) -> dict[str, Any]:
    tables = _find_tables(content)
    if heading is not None or heading_path is not None:
        headings = _parse_headings(content)
        selected_heading = headings[_match_heading(headings, heading, heading_path, exact)]
        tables = [table for table in tables if table["heading"] and table["heading"]["line"] == selected_heading.line]
    if table_index < 0 or table_index >= len(tables):
        raise ValueError("Requested markdown table was not found")
    return tables[table_index]


def _find_fenced_blocks(content: str, languages: set[str] | None = None) -> list[dict[str, Any]]:
    lines = _to_lines(content)
    headings = _parse_headings(content)
    blocks = []
    active: dict[str, Any] | None = None
    for number, raw_line in enumerate(lines, start=1):
        line = raw_line.rstrip("\r")
        match = re.match(r"^\s*([`~]{3,})(.*)$", line)
        if not match:
            continue
        marker = match.group(1)[0]
        length = len(match.group(1))
        if active and marker == active["marker"] and length >= active["length"]:
            info = active["info"]
            language = info.split()[0].casefold() if info else ""
            if languages is None or language in languages:
                blocks.append({"index": len(blocks), "start_line": active["start_line"], "end_line": number, "info": info, "language": language, "heading": _heading_for_line(headings, active["start_line"])})
            active = None
        elif active is None:
            active = {"start_line": number, "marker": marker, "length": length, "info": match.group(2).strip()}
    return blocks


def _read_block_source(content: str, block: dict[str, Any]) -> str:
    lines = _to_lines(content)
    return "\n".join(line.rstrip("\r") for line in lines[block["start_line"] : block["end_line"] - 1])

def _limit_text(value: str, max_chars: int | None = None, preview: bool = False) -> str:
    limit = max_chars if max_chars is not None else (1000 if preview else None)
    if limit is None or limit < 0 or len(value) <= limit:
        return value
    return value[:limit] + "\n… truncated …"


def _split_frontmatter(content: str) -> tuple[str | None, str]:
    lines = _to_lines(content)
    if not lines or lines[0].strip() != "---":
        return None, content
    for index in range(1, len(lines)):
        if lines[index].strip() == "---":
            frontmatter = "\n".join(lines[1:index])
            body = "\n".join(lines[index + 1 :])
            if content.endswith("\n") and body:
                body += "\n"
            return frontmatter, body
    return None, content


def _parse_simple_yaml(value: str) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for raw_line in value.split("\n"):
        line = raw_line.strip()
        if not line or line.startswith("#") or ":" not in line:
            continue
        key, raw_value = line.split(":", 1)
        item = raw_value.strip()
        if item in {"true", "false"}:
            parsed: Any = item == "true"
        elif item.startswith("[") and item.endswith("]"):
            parsed = [part.strip().strip('"\'') for part in item[1:-1].split(",") if part.strip()]
        else:
            parsed = item.strip('"\'')
        result[key.strip()] = parsed
    return result


def _render_simple_yaml(data: dict[str, Any]) -> str:
    lines = []
    for key, value in data.items():
        if isinstance(value, bool):
            rendered = "true" if value else "false"
        elif isinstance(value, list):
            rendered = "[" + ", ".join(str(item) for item in value) + "]"
        else:
            rendered = str(value)
        lines.append(f"{key}: {rendered}")
    return "\n".join(lines)


def _anchor_candidates(path: Path, target: str) -> list[str]:
    parsed = urlparse(target)
    if parsed.scheme or parsed.netloc:
        return []
    fragment = unquote(parsed.fragment or "")
    target_path = unquote(parsed.path or "")
    candidates = []
    if target_path:
        candidates.append(str((path.parent / target_path).resolve()))
    if fragment:
        candidates.append("#" + fragment)
    return candidates


def _apply_heading_delta_to_block(block: str, delta: int) -> str:
    updated_lines = []
    heading_re = re.compile(r"^(\s*)(#{1,6})([ \t]+.+)$")
    active_fence: tuple[str, int] | None = None
    for raw_line in _to_lines(block):
        line = raw_line.rstrip("\r")
        fence_match = re.match(r"^\s*([`~]{3,})", line)
        if fence_match:
            marker = fence_match.group(1)[0]
            length = len(fence_match.group(1))
            if active_fence and active_fence[0] == marker and length >= active_fence[1]:
                active_fence = None
            elif active_fence is None:
                active_fence = (marker, length)
            updated_lines.append(line)
            continue
        match = heading_re.match(line) if active_fence is None else None
        if match:
            new_level = len(match.group(2)) + delta
            if not 1 <= new_level <= 6:
                raise ValueError("heading level change would move a heading outside level 1..6")
            updated_lines.append(match.group(1) + "#" * new_level + match.group(3))
        else:
            updated_lines.append(line)
    return "\n".join(updated_lines) + ("\n" if block.endswith("\n") else "")


@mcp.tool()
def markdown_outline(path: str, max_depth: int | None = None, compact: bool = False) -> dict[str, Any]:
    """List Markdown headings with line numbers, hierarchy, and section spans."""
    resolved_path, content, _, total_lines, _ = _load_markdown_file(path)
    headings = _parse_headings(content)
    items = []
    for index, heading in enumerate(headings):
        if max_depth is not None and heading.level > max_depth:
            continue
        section_end = _section_end_line(headings, index, total_lines, include_subsections=True)
        item = {
            "title": heading.title,
            "level": heading.level,
            "line": heading.line,
            "section_end_line": section_end,
            "slug": heading.slug,
        }
        if not compact:
            item["path"] = heading.path
            item["path_text"] = " > ".join(heading.path)
        items.append(item)
    return {"path": str(resolved_path), "total_lines": total_lines, "heading_count": len(headings), "headings": items}


@mcp.tool()
def read_markdown_section(path: str, heading: str | None = None, heading_path: list[str] | None = None, exact: bool = False, include_heading: bool = True, include_subsections: bool = True, include_line_numbers: bool = False, max_lines: int | None = None, max_bytes: int | None = None, max_chars: int | None = None, preview: bool = False) -> str:
    """Return a Markdown section as raw Markdown text by heading name or heading path."""
    if heading is None and heading_path is None:
        raise ValueError("Either heading or heading_path is required")
    _, content, _, total_lines, _ = _load_markdown_file(path)
    headings = _parse_headings(content)
    heading_index = _match_heading(headings, heading, heading_path, exact)
    selected = headings[heading_index]
    section_end = _section_end_line(headings, heading_index, total_lines, include_subsections)
    start_line = selected.line if include_heading else selected.line + 1
    if start_line > section_end:
        return ""
    return _limit_text(_slice_lines(content, start_line, section_end, include_line_numbers, max_lines, max_bytes), max_chars, preview)


@mcp.tool()
def md_read_range(path: str, start_line: int, end_line: int, include_line_numbers: bool = False, max_lines: int | None = None, max_bytes: int | None = None) -> str:
    """Read a precise 1-based inclusive Markdown line range."""
    _, content, _, total_lines, _ = _load_markdown_file(path)
    if start_line < 1 or end_line < start_line:
        raise ValueError("Use 1-based inclusive line numbers with end_line >= start_line")
    return _slice_lines(content, start_line, min(end_line, total_lines), include_line_numbers, max_lines, max_bytes)

@mcp.tool()
def md_read_near(path: str, text: str | None = None, regex: str | None = None, line: int | None = None, before: int = 5, after: int = 5, case_sensitive: bool = False, include_line_numbers: bool = True, max_bytes: int | None = None) -> str:
    """Read Markdown context around a matching text/regex or around a line number."""
    _, content, _, total_lines, _ = _load_markdown_file(path)
    if line is None:
        if (text is None) == (regex is None):
            raise ValueError("Provide exactly one of line, text, or regex")
        flags = 0 if case_sensitive else re.IGNORECASE
        pattern = re.compile(regex if regex is not None else re.escape(text or ""), flags)
        for line_number, raw_line in enumerate(_to_lines(content), start=1):
            if pattern.search(raw_line):
                line = line_number
                break
        if line is None:
            raise ValueError("No matching text found")
    start = max(1, line - max(before, 0))
    end = min(total_lines, line + max(after, 0))
    return _slice_lines(content, start, end, include_line_numbers, None, max_bytes)

@mcp.tool()
def md_get_document_map(path: str) -> dict[str, Any]:
    """Return a compact map of headings, tables, diagrams, code blocks, links, and images."""
    resolved_path, content, _, total_lines, size = _load_markdown_file(path)
    headings = [_heading_metadata(heading) for heading in _parse_headings(content)]
    tables = [{k: v for k, v in table.items() if k != "raw_lines"} for table in _find_tables(content)]
    diagrams = _find_fenced_blocks(content, {"mermaid", "plantuml", "dot", "graphviz"})
    code_blocks = _find_fenced_blocks(content, None)
    inline_link_re = re.compile(r"!?\[([^\]]*)\]\(([^)]+)\)")
    ref_link_re = re.compile(r"^\s*\[([^\]]+)\]:\s*(\S+)")
    links = []
    images = []
    for line_number, raw_line in enumerate(_to_lines(content), start=1):
        line_text = raw_line.rstrip("\r")
        for match in inline_link_re.finditer(line_text):
            item = {"line": line_number, "text": match.group(1), "target": match.group(2)}
            if match.group(0).startswith("!"):
                images.append({"line": line_number, "alt": match.group(1), "src": match.group(2)})
            else:
                links.append({"type": "inline", **item})
        ref_match = ref_link_re.match(line_text)
        if ref_match:
            links.append({"type": "reference", "line": line_number, "text": ref_match.group(1), "target": ref_match.group(2)})
    return {
        "path": str(resolved_path),
        "bytes": size,
        "total_lines": total_lines,
        "heading_count": len(headings),
        "table_count": len(tables),
        "diagram_count": len(diagrams),
        "code_block_count": len(code_blocks),
        "link_count": len(links),
        "image_count": len(images),
        "headings": headings,
        "tables": tables,
        "diagrams": diagrams,
        "code_blocks": code_blocks,
        "links": links,
        "images": images,
    }
@mcp.tool()
def replace_markdown_section(path: str, new_content: str, heading: str | None = None, heading_path: list[str] | None = None, exact: bool = False, include_heading: bool = False, include_subsections: bool = True, ensure_trailing_newline: bool = True) -> dict[str, Any]:
    """Replace a Markdown section by heading name or heading path without exact text matching."""
    if heading is None and heading_path is None:
        raise ValueError("Either heading or heading_path is required")
    resolved_path, content, encoding, total_lines, size = _load_markdown_file(path)
    headings = _parse_headings(content)
    heading_index = _match_heading(headings, heading, heading_path, exact)
    selected = headings[heading_index]
    section_end = _section_end_line(headings, heading_index, total_lines, include_subsections)
    start_line = selected.line if include_heading else selected.line + 1
    end_line = section_end
    offsets = _line_start_offsets(content)
    start_byte = _line_start_offset(offsets, start_line, len(content))
    end_byte = _line_start_offset(offsets, end_line + 1, len(content)) if start_line <= end_line else start_byte
    replacement = new_content
    if ensure_trailing_newline and replacement and not replacement.endswith("\n"):
        replacement += "\n"
    updated = content[:start_byte] + replacement + content[end_byte:]
    final_bytes = _encode_fuzzy(updated, encoding)
    changed = final_bytes != resolved_path.read_bytes()
    if changed:
        resolved_path.write_bytes(final_bytes)
    return {
        "success": True,
        "path": str(resolved_path),
        "changed": changed,
        "bytes_before": size,
        "bytes_written": len(final_bytes),
        "target_encoding": encoding,
        "include_heading": include_heading,
        "include_subsections": include_subsections,
        "replaced_start_line": start_line,
        "replaced_end_line": end_line,
        "selected_heading": _heading_metadata(selected),
        "section": {"section_start_line": selected.line, "section_end_line": section_end},
        "message": "markdown section replaced" if changed else "no content changes detected",
    }


@mcp.tool()
def md_search(path: str, query: str, regex: bool = False, case_sensitive: bool = False, max_results: int = 50, context_lines: int = 0) -> dict[str, Any]:
    """Search Markdown text and return line matches with heading context."""
    resolved_path, content, _, total_lines, _ = _load_markdown_file(path)
    headings = _parse_headings(content)
    flags = 0 if case_sensitive else re.IGNORECASE
    pattern = re.compile(query if regex else re.escape(query), flags)
    matches = []
    for index, raw_line in enumerate(_to_lines(content), start=1):
        line = raw_line.rstrip("\r")
        found = list(pattern.finditer(line))
        if not found:
            continue
        context_start = max(1, index - context_lines)
        context_end = min(total_lines, index + context_lines)
        matches.append({
            "line": index,
            "text": line,
            "spans": [{"start": item.start(), "end": item.end(), "text": item.group(0)} for item in found],
            "heading": _heading_for_line(headings, index),
            "context": _slice_lines(content, context_start, context_end, True, None, None) if context_lines > 0 else None,
        })
        if len(matches) >= max_results:
            break
    return {
        "path": str(resolved_path),
        "query": query,
        "regex": regex,
        "case_sensitive": case_sensitive,
        "total_lines": total_lines,
        "match_count": len(matches),
        "truncated": len(matches) >= max_results,
        "matches": matches,
    }


@mcp.tool()
def md_insert_section(path: str, title: str, content: str = "", parent_heading: str | None = None, parent_heading_path: list[str] | None = None, level: int | None = None, position: str = "end", exact: bool = False, ensure_blank_lines: bool = True) -> dict[str, Any]:
    """Insert a new Markdown section at the document end or under a parent heading."""
    if position not in {"start", "end"}:
        raise ValueError("position must be 'start' or 'end'")
    resolved_path, old_content, encoding, total_lines, size = _load_markdown_file(path)
    headings = _parse_headings(old_content)
    parent = None
    if parent_heading is not None or parent_heading_path is not None:
        parent_index = _match_heading(headings, parent_heading, parent_heading_path, exact)
        parent = headings[parent_index]
        insert_line = parent.line + 1 if position == "start" else _section_end_line(headings, parent_index, total_lines, True) + 1
        section_level = level if level is not None else min(parent.level + 1, 6)
    else:
        section_level = level if level is not None else 1
        insert_line = 1 if position == "start" else total_lines + 1
    heading_line = _format_heading_line(section_level, title)
    body = content.strip("\n")
    new_section = heading_line + "\n" + (body + "\n" if body else "")
    offsets = _line_start_offsets(old_content)
    insert_byte = _line_start_offset(offsets, insert_line, len(old_content))
    prefix = old_content[:insert_byte]
    suffix = old_content[insert_byte:]
    if ensure_blank_lines:
        if prefix and not prefix.endswith("\n\n"):
            prefix = prefix.rstrip("\n") + "\n\n"
        if suffix and not new_section.endswith("\n\n"):
            new_section = new_section.rstrip("\n") + "\n\n"
    updated = prefix + new_section + suffix
    write = _write_markdown_file(resolved_path, updated, encoding)
    return {
        "success": True,
        "path": str(resolved_path),
        **write,
        "target_encoding": encoding,
        "inserted_line": insert_line,
        "inserted_heading": {"title": title, "level": section_level, "slug": _slugify_heading(title)},
        "parent_heading": _heading_metadata(parent) if parent else None,
        "message": "markdown section inserted" if write["changed"] else "no content changes detected",
        "original_bytes_before": size,
    }


@mcp.tool()
def md_delete_section(path: str, heading: str | None = None, heading_path: list[str] | None = None, exact: bool = False, include_subsections: bool = True) -> dict[str, Any]:
    """Delete a Markdown section by heading name or heading path."""
    resolved_path, content, encoding, _, size = _load_markdown_file(path)
    _, _, selected, section_end = _section_bounds(content, heading, heading_path, exact, include_subsections)
    updated = _replace_line_range(content, selected.line, section_end, "")
    write = _write_markdown_file(resolved_path, updated, encoding)
    return {"success": True, "path": str(resolved_path), **write, "bytes_before": size, "deleted_heading": _heading_metadata(selected), "deleted_start_line": selected.line, "deleted_end_line": section_end}


@mcp.tool()
def md_append_to_section(path: str, content: str, heading: str | None = None, heading_path: list[str] | None = None, exact: bool = False, position: str = "end", ensure_blank_lines: bool = True) -> dict[str, Any]:
    """Append or prepend Markdown content inside a section body without overwriting it."""
    if position not in {"start", "end"}:
        raise ValueError("position must be 'start' or 'end'")
    resolved_path, old_content, encoding, total_lines, size = _load_markdown_file(path)
    _, _, selected, section_end = _section_bounds(old_content, heading, heading_path, exact, True)
    insert_line = selected.line + 1 if position == "start" else section_end + 1
    block = _normalize_block(content)
    if ensure_blank_lines:
        block = "\n" + block if position == "start" else block + "\n"
    updated = _replace_line_range(old_content, insert_line, insert_line - 1, block)
    write = _write_markdown_file(resolved_path, updated, encoding)
    return {"success": True, "path": str(resolved_path), **write, "bytes_before": size, "inserted_line": min(insert_line, total_lines + 1), "selected_heading": _heading_metadata(selected)}


@mcp.tool()
def md_rename_heading(path: str, new_title: str, heading: str | None = None, heading_path: list[str] | None = None, exact: bool = False) -> dict[str, Any]:
    """Rename a Markdown heading while preserving its level and section body."""
    resolved_path, content, encoding, _, size = _load_markdown_file(path)
    _, _, selected, _ = _section_bounds(content, heading, heading_path, exact, True)
    updated = _replace_line_range(content, selected.line, selected.line, _format_heading_line(selected.level, new_title) + "\n")
    write = _write_markdown_file(resolved_path, updated, encoding)
    return {"success": True, "path": str(resolved_path), **write, "bytes_before": size, "old_heading": _heading_metadata(selected), "new_heading": {"title": new_title.strip(), "level": selected.level, "slug": _slugify_heading(new_title)}}


@mcp.tool()
def md_replace_text(path: str, find: str, replace: str, regex: bool = False, case_sensitive: bool = True, heading: str | None = None, heading_path: list[str] | None = None, exact: bool = False, max_replacements: int = 0, dry_run: bool = False) -> dict[str, Any]:
    """Find and replace literal or regex text in the whole document or one section."""
    resolved_path, content, encoding, _, size = _load_markdown_file(path)
    flags = 0 if case_sensitive else re.IGNORECASE
    pattern = re.compile(find if regex else re.escape(find), flags)
    start_line, end_line = 1, len(_to_lines(content))
    if heading is not None or heading_path is not None:
        _, _, selected, section_end = _section_bounds(content, heading, heading_path, exact, True)
        start_line, end_line = selected.line + 1, section_end
    segment = _slice_lines(content, start_line, end_line, False, None, None)
    replaced, count = pattern.subn(replace, segment, count=max_replacements if max_replacements > 0 else 0)
    if dry_run:
        return {"success": True, "path": str(resolved_path), "changed": False, "dry_run": True, "replacement_count": count, "start_line": start_line, "end_line": end_line}
    updated = _replace_line_range(content, start_line, end_line, replaced + ("\n" if replaced else ""))
    write = _write_markdown_file(resolved_path, updated, encoding)
    return {"success": True, "path": str(resolved_path), **write, "bytes_before": size, "replacement_count": count}


@mcp.tool()
def md_list_tables(path: str) -> dict[str, Any]:
    """List Markdown pipe tables with line ranges, heading context, shape, and headers."""
    resolved_path, content, _, _, _ = _load_markdown_file(path)
    tables = [{key: value for key, value in table.items() if key != "raw_lines"} for table in _find_tables(content)]
    return {"path": str(resolved_path), "table_count": len(tables), "tables": tables}


@mcp.tool()
def md_read_table(path: str, table_index: int = 0, heading: str | None = None, heading_path: list[str] | None = None, exact: bool = False, include_body: bool = True, max_rows: int | None = None) -> dict[str, Any]:
    """Read a Markdown pipe table as structured headers and rows."""
    resolved_path, content, _, _, _ = _load_markdown_file(path)
    table = _select_table(content, table_index, heading, heading_path, exact)
    raw = table["raw_lines"]
    rows = [_parse_pipe_row(line) for line in raw[2:]]
    if max_rows is not None:
        rows = rows[:max_rows]
    result = {"path": str(resolved_path), "table": {key: value for key, value in table.items() if key != "raw_lines"}, "headers": table["headers"], "alignments": _table_alignments(raw[1], table["columns"])}
    if include_body:
        result["rows"] = rows
    return result


@mcp.tool()
def md_format_table(path: str, table_index: int = 0, heading: str | None = None, heading_path: list[str] | None = None, exact: bool = False) -> dict[str, Any]:
    """Pretty-print one Markdown pipe table with normalized spacing and alignment markers."""
    resolved_path, content, encoding, _, size = _load_markdown_file(path)
    table = _select_table(content, table_index, heading, heading_path, exact)
    raw = table["raw_lines"]
    rendered = _render_table(table["headers"], [_parse_pipe_row(line) for line in raw[2:]], _table_alignments(raw[1], table["columns"]))
    updated = _replace_line_range(content, table["start_line"], table["end_line"], rendered)
    write = _write_markdown_file(resolved_path, updated, encoding)
    return {"success": True, "path": str(resolved_path), **write, "bytes_before": size, "table_index": table["index"]}


@mcp.tool()
def md_edit_table(path: str, op: str, table_index: int = 0, heading: str | None = None, heading_path: list[str] | None = None, exact: bool = False, row: int | None = None, col: int | str | None = None, value: str = "", values: list[str] | None = None, name: str = "", align: str = "left", reverse: bool = False) -> dict[str, Any]:
    """Edit a Markdown table using op: set_cell, add_row, add_col, del_row, del_col, sort, set_align."""
    resolved_path, content, encoding, _, size = _load_markdown_file(path)
    table = _select_table(content, table_index, heading, heading_path, exact)
    raw = table["raw_lines"]
    headers = table["headers"][:]
    rows = [_parse_pipe_row(line) for line in raw[2:]]
    alignments = _table_alignments(raw[1], len(headers))

    def col_index(column: int | str | None) -> int:
        if isinstance(column, int):
            return column
        if isinstance(column, str) and column in headers:
            return headers.index(column)
        raise ValueError("A valid col index or header name is required")

    if op == "set_cell":
        if row is None:
            raise ValueError("row is required")
        rows[row][col_index(col)] = value
    elif op == "add_row":
        rows.append((values or []).copy())
    elif op == "add_col":
        headers.append(name or value or "Column")
        alignments.append(align)
        column_values = values or []
        for index, item in enumerate(rows):
            item.append(column_values[index] if index < len(column_values) else "")
    elif op == "del_row":
        if row is None:
            raise ValueError("row is required")
        del rows[row]
    elif op == "del_col":
        column = col_index(col)
        del headers[column]
        del alignments[column]
        for item in rows:
            del item[column]
    elif op == "sort":
        column = col_index(col)
        rows.sort(key=lambda item: item[column] if column < len(item) else "", reverse=reverse)
    elif op == "set_align":
        alignments[col_index(col)] = align
    else:
        raise ValueError("op must be one of: set_cell, add_row, add_col, del_row, del_col, sort, set_align")
    rendered = _render_table(headers, rows, alignments)
    updated = _replace_line_range(content, table["start_line"], table["end_line"], rendered)
    write = _write_markdown_file(resolved_path, updated, encoding)
    return {"success": True, "path": str(resolved_path), **write, "bytes_before": size, "op": op, "table_index": table["index"]}


@mcp.tool()
def md_insert_table(path: str, headers: list[str], rows: list[list[str]], heading: str | None = None, heading_path: list[str] | None = None, exact: bool = False, position: str = "end", alignments: list[str] | None = None) -> dict[str, Any]:
    """Insert a Markdown table from structured headers and rows."""
    resolved_path, content, encoding, total_lines, size = _load_markdown_file(path)
    if position not in {"start", "end"}:
        raise ValueError("position must be 'start' or 'end'")
    insert_line = 1 if position == "start" else total_lines + 1
    if heading is not None or heading_path is not None:
        _, _, selected, section_end = _section_bounds(content, heading, heading_path, exact, True)
        insert_line = selected.line + 1 if position == "start" else section_end + 1
    block = "\n" + _render_table(headers, rows, alignments) + "\n"
    updated = _replace_line_range(content, insert_line, insert_line - 1, block)
    write = _write_markdown_file(resolved_path, updated, encoding)
    return {"success": True, "path": str(resolved_path), **write, "bytes_before": size, "inserted_line": insert_line}


@mcp.tool()
def md_table_export(path: str, table_index: int = 0, format: str = "csv", heading: str | None = None, heading_path: list[str] | None = None, exact: bool = False) -> str:
    """Export one Markdown table to CSV or JSON text."""
    _, content, _, _, _ = _load_markdown_file(path)
    table = _select_table(content, table_index, heading, heading_path, exact)
    rows = [_parse_pipe_row(line) for line in table["raw_lines"][2:]]
    if format == "json":
        import json
        return json.dumps([dict(zip(table["headers"], row)) for row in rows], ensure_ascii=False, indent=2)
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(table["headers"])
    writer.writerows(rows)
    return output.getvalue()


@mcp.tool()
def md_list_diagrams(path: str) -> dict[str, Any]:
    """List fenced diagram blocks such as mermaid, plantuml, and dot."""
    resolved_path, content, _, _, _ = _load_markdown_file(path)
    blocks = _find_fenced_blocks(content, {"mermaid", "plantuml", "puml", "dot", "graphviz"})
    return {"path": str(resolved_path), "diagram_count": len(blocks), "diagrams": blocks}


@mcp.tool()
def md_read_diagram(path: str, diagram_index: int = 0, include_body: bool = True, max_chars: int | None = None, preview: bool = False) -> dict[str, Any]:
    """Read the source of a fenced diagram block."""
    resolved_path, content, _, _, _ = _load_markdown_file(path)
    block = _find_fenced_blocks(content, {"mermaid", "plantuml", "puml", "dot", "graphviz"})[diagram_index]
    result = {"path": str(resolved_path), "diagram": block}
    if include_body:
        result["source"] = _limit_text(_read_block_source(content, block), max_chars, preview)
    return result


@mcp.tool()
def md_insert_diagram(path: str, source: str, diagram_type: str = "mermaid", heading: str | None = None, heading_path: list[str] | None = None, exact: bool = False, position: str = "end") -> dict[str, Any]:
    """Insert a fenced diagram block."""
    return md_insert_code_block(path, source, info=diagram_type, heading=heading, heading_path=heading_path, exact=exact, position=position)


@mcp.tool()
def md_replace_diagram(path: str, source: str, diagram_index: int = 0) -> dict[str, Any]:
    """Replace the source inside one fenced diagram block."""
    return md_replace_code_block(path, source, block_index=diagram_index, language="mermaid|plantuml|puml|dot|graphviz", regex_language=True)


@mcp.tool()
def md_extract_code_blocks(path: str, language: str | None = None, regex_language: bool = False, include_body: bool = True, max_chars: int | None = None, preview: bool = False) -> dict[str, Any]:
    """List fenced code blocks and return their source."""
    resolved_path, content, _, _, _ = _load_markdown_file(path)
    blocks = _find_fenced_blocks(content)
    if language:
        if regex_language:
            pattern = re.compile(language, re.IGNORECASE)
            blocks = [block for block in blocks if pattern.search(block["language"])]
        else:
            blocks = [block for block in blocks if block["language"] == language.casefold()]
    result_blocks = []
    for block in blocks:
        item = dict(block)
        if include_body:
            item["source"] = _limit_text(_read_block_source(content, block), max_chars, preview)
        result_blocks.append(item)
    return {"path": str(resolved_path), "block_count": len(blocks), "blocks": result_blocks}


@mcp.tool()
def md_replace_code_block(path: str, source: str, block_index: int = 0, language: str | None = None, regex_language: bool = False) -> dict[str, Any]:
    """Replace the source inside a fenced code block while keeping its fence and info string."""
    resolved_path, content, encoding, _, size = _load_markdown_file(path)
    blocks = md_extract_code_blocks(path, language, regex_language)["blocks"] if language else _find_fenced_blocks(content)
    if block_index < 0 or block_index >= len(blocks):
        raise ValueError("Requested code block was not found")
    block = blocks[block_index]
    updated = _replace_line_range(content, block["start_line"] + 1, block["end_line"] - 1, _normalize_block(source))
    write = _write_markdown_file(resolved_path, updated, encoding)
    return {"success": True, "path": str(resolved_path), **write, "bytes_before": size, "block": block}


@mcp.tool()
def md_insert_code_block(path: str, source: str, info: str = "", heading: str | None = None, heading_path: list[str] | None = None, exact: bool = False, position: str = "end") -> dict[str, Any]:
    """Insert a fenced code block with an optional info string."""
    resolved_path, content, encoding, total_lines, size = _load_markdown_file(path)
    insert_line = 1 if position == "start" else total_lines + 1
    if heading is not None or heading_path is not None:
        _, _, selected, section_end = _section_bounds(content, heading, heading_path, exact, True)
        insert_line = selected.line + 1 if position == "start" else section_end + 1
    block = f"\n```{info.strip()}\n{source.strip(chr(10))}\n```\n"
    updated = _replace_line_range(content, insert_line, insert_line - 1, block)
    write = _write_markdown_file(resolved_path, updated, encoding)
    return {"success": True, "path": str(resolved_path), **write, "bytes_before": size, "inserted_line": insert_line, "info": info}


@mcp.tool()
def md_list_links(path: str) -> dict[str, Any]:
    """List inline and reference Markdown links with line numbers."""
    resolved_path, content, _, _, _ = _load_markdown_file(path)
    links = []
    inline_re = re.compile(r"!?\[([^\]]*)\]\(([^)]+)\)")
    ref_re = re.compile(r"^\s*\[([^\]]+)\]:\s*(\S+)")
    for line_number, raw_line in enumerate(_to_lines(content), start=1):
        line = raw_line.rstrip("\r")
        for match in inline_re.finditer(line):
            if match.group(0).startswith("!"):
                continue
            links.append({"type": "inline", "line": line_number, "text": match.group(1), "target": match.group(2)})
        match = ref_re.match(line)
        if match:
            links.append({"type": "reference", "line": line_number, "text": match.group(1), "target": match.group(2)})
    return {"path": str(resolved_path), "link_count": len(links), "links": links}


@mcp.tool()
def md_list_images(path: str) -> dict[str, Any]:
    """List Markdown images with alt text, source, and line numbers."""
    resolved_path, content, _, _, _ = _load_markdown_file(path)
    image_re = re.compile(r"!\[([^\]]*)\]\(([^)]+)\)")
    images = []
    for line_number, raw_line in enumerate(_to_lines(content), start=1):
        for match in image_re.finditer(raw_line.rstrip("\r")):
            images.append({"line": line_number, "alt": match.group(1), "src": match.group(2)})
    return {"path": str(resolved_path), "image_count": len(images), "images": images}


@mcp.tool()
def md_get_anchor(title: str) -> dict[str, str]:
    """Return the GitHub-compatible anchor slug for a heading title."""
    return {"title": title, "anchor": _slugify_heading(title), "href": "#" + _slugify_heading(title)}


@mcp.tool()
def md_update_toc(path: str, min_level: int = 1, max_level: int = 3, marker_start: str = "<!-- md-tools-toc:start -->", marker_end: str = "<!-- md-tools-toc:end -->") -> dict[str, Any]:
    """Generate or refresh a Markdown table of contents between marker comments."""
    resolved_path, content, encoding, _, size = _load_markdown_file(path)
    headings = [heading for heading in _parse_headings(content) if min_level <= heading.level <= max_level and heading.title not in {marker_start, marker_end}]
    toc = [marker_start]
    for heading in headings:
        indent = "  " * max(heading.level - min_level, 0)
        toc.append(f"{indent}- [{heading.title}](#{heading.slug})")
    toc.append(marker_end)
    toc_block = "\n".join(toc) + "\n"
    start = content.find(marker_start)
    end = content.find(marker_end)
    updated = content[:start] + toc_block + content[end + len(marker_end):].lstrip("\n") if start != -1 and end != -1 and end > start else toc_block + "\n" + content
    write = _write_markdown_file(resolved_path, updated, encoding)
    return {"success": True, "path": str(resolved_path), **write, "bytes_before": size, "heading_count": len(headings)}


@mcp.tool()
def md_stats(path: str) -> dict[str, Any]:
    """Return simple Markdown document statistics."""
    resolved_path, content, _, total_lines, size = _load_markdown_file(path)
    headings = _parse_headings(content)
    words = re.findall(r"\b\w+\b", content, re.UNICODE)
    return {"path": str(resolved_path), "bytes": size, "lines": total_lines, "words": len(words), "reading_time_minutes": max(1, round(len(words) / 200)), "heading_count": len(headings), "sections": len(headings), "tables": len(_find_tables(content)), "code_blocks": len(_find_fenced_blocks(content))}


@mcp.tool()
def md_patch_lines(path: str, start_line: int, end_line: int, replacement: str, ensure_trailing_newline: bool = True) -> dict[str, Any]:
    """Replace a precise 1-based inclusive line range with Markdown text."""
    resolved_path, content, encoding, _, size = _load_markdown_file(path)
    if start_line < 1 or end_line < start_line - 1:
        raise ValueError("Use 1-based lines; end_line may be start_line - 1 only for insertion")
    new_text = replacement
    if ensure_trailing_newline and new_text and not new_text.endswith("\n"):
        new_text += "\n"
    updated = _replace_line_range(content, start_line, end_line, new_text)
    write = _write_markdown_file(resolved_path, updated, encoding)
    return {"success": True, "path": str(resolved_path), **write, "bytes_before": size, "start_line": start_line, "end_line": end_line}

@mcp.tool()
def md_normalize_headings(path: str, max_increment: int = 1, dry_run: bool = False) -> dict[str, Any]:
    """Normalize heading hierarchy so levels do not jump by more than max_increment."""
    resolved_path, content, encoding, _, size = _load_markdown_file(path)
    lines = _to_lines(content)
    in_fence: tuple[str, int] | None = None
    previous_level = 0
    changes = []
    heading_re = re.compile(r"^(#{1,6})([ \t]+.+)$")
    fence_re = re.compile(r"^\s*([`~]{3,})")
    for index, raw_line in enumerate(lines):
        line = raw_line.rstrip("\r")
        fence = fence_re.match(line)
        if fence:
            marker = fence.group(1)[0]
            length = len(fence.group(1))
            if in_fence and marker == in_fence[0] and length >= in_fence[1]:
                in_fence = None
            elif in_fence is None:
                in_fence = (marker, length)
            continue
        if in_fence:
            continue
        match = heading_re.match(line)
        if not match:
            continue
        level = len(match.group(1))
        allowed = 1 if previous_level == 0 else min(6, previous_level + max(1, max_increment))
        new_level = min(level, allowed)
        if new_level != level:
            lines[index] = "#" * new_level + match.group(2)
            changes.append({"line": index + 1, "old_level": level, "new_level": new_level})
        previous_level = new_level
    if dry_run:
        return {"success": True, "path": str(resolved_path), "changed": False, "dry_run": True, "change_count": len(changes), "changes": changes}
    updated = "\n".join(lines) + ("\n" if content.endswith("\n") else "")
    write = _write_markdown_file(resolved_path, updated, encoding)
    return {"success": True, "path": str(resolved_path), **write, "bytes_before": size, "change_count": len(changes), "changes": changes}

@mcp.tool()
def md_check_internal_links(path: str) -> dict[str, Any]:
    """Validate local file links and same-file heading anchors without remote checks."""
    result = md_validate_links(path, check_remote=False)
    result["remote_checked"] = False
    return result

@mcp.tool()
def md_set_heading_level(path: str, level: int, heading: str | None = None, heading_path: list[str] | None = None, exact: bool = False) -> dict[str, Any]:
    """Set a heading level and shift every heading in its subtree by the same delta."""
    resolved_path, content, encoding, _, size = _load_markdown_file(path)
    _, _, selected, section_end = _section_bounds(content, heading, heading_path, exact, True)
    if not 1 <= level <= 6:
        raise ValueError("level must be between 1 and 6")
    delta = level - selected.level
    block = _slice_lines(content, selected.line, section_end, False, None, None)
    updated_block = _apply_heading_delta_to_block(block, delta)
    updated = _replace_line_range(content, selected.line, section_end, updated_block)
    write = _write_markdown_file(resolved_path, updated, encoding)
    return {"success": True, "path": str(resolved_path), **write, "bytes_before": size, "old_level": selected.level, "new_level": level, "selected_heading": _heading_metadata(selected)}


@mcp.tool()
def md_move_section(path: str, heading: str | None = None, heading_path: list[str] | None = None, target_heading: str | None = None, target_heading_path: list[str] | None = None, exact: bool = False, position: str = "after", under_parent: bool = False) -> dict[str, Any]:
    """Move a section before/after another heading, or under another heading as a child."""
    if position not in {"before", "after", "start", "end"}:
        raise ValueError("position must be before, after, start, or end")
    resolved_path, content, encoding, total_lines, size = _load_markdown_file(path)
    headings, source_index, selected, section_end = _section_bounds(content, heading, heading_path, exact, True)
    if target_heading is None and target_heading_path is None:
        raise ValueError("target_heading or target_heading_path is required")
    target_index = _match_heading(headings, target_heading, target_heading_path, exact)
    target = headings[target_index]
    if selected.line <= target.line <= section_end:
        raise ValueError("Cannot move a section relative to itself or its subtree")
    block = _slice_lines(content, selected.line, section_end, False, None, None)
    if under_parent:
        delta = min(target.level + 1, 6) - selected.level
        block = _apply_heading_delta_to_block(block, delta)
    without = _replace_line_range(content, selected.line, section_end, "")
    adjusted_headings = _parse_headings(without)
    adjusted_target = adjusted_headings[_match_heading(adjusted_headings, target.title, target.path, True)]
    adjusted_index = adjusted_headings.index(adjusted_target)
    adjusted_total = len(_to_lines(without))
    if under_parent:
        insert_line = adjusted_target.line + 1 if position == "start" else _section_end_line(adjusted_headings, adjusted_index, adjusted_total, True) + 1
    elif position == "before":
        insert_line = adjusted_target.line
    else:
        insert_line = _section_end_line(adjusted_headings, adjusted_index, adjusted_total, True) + 1
    updated = _replace_line_range(without, insert_line, insert_line - 1, "\n" + block.strip("\n") + "\n")
    write = _write_markdown_file(resolved_path, updated, encoding)
    return {"success": True, "path": str(resolved_path), **write, "bytes_before": size, "moved_heading": _heading_metadata(selected), "target_heading": _heading_metadata(target), "inserted_line": insert_line, "original_total_lines": total_lines}


@mcp.tool()
def md_rewrite_links(path: str, replacements: dict[str, str], regex: bool = False) -> dict[str, Any]:
    """Bulk rewrite Markdown link and image targets."""
    resolved_path, content, encoding, _, size = _load_markdown_file(path)
    count = 0

    def rewrite_target(target: str) -> str:
        nonlocal count
        for old, new in replacements.items():
            if (regex and re.search(old, target)) or (not regex and target == old):
                count += 1
                return re.sub(old, new, target) if regex else new
        return target

    def repl(match: re.Match[str]) -> str:
        bang, text, target = match.group(1), match.group(2), match.group(3)
        return f"{bang}[{text}]({rewrite_target(target)})"

    updated = re.sub(r"(!?)\[([^\]]*)\]\(([^)]+)\)", repl, content)
    write = _write_markdown_file(resolved_path, updated, encoding)
    return {"success": True, "path": str(resolved_path), **write, "bytes_before": size, "replacement_count": count}


@mcp.tool()
def md_validate_links(path: str, check_remote: bool = False) -> dict[str, Any]:
    """Validate local file links and same-file heading anchors; remote links are counted as skipped."""
    resolved_path, content, _, _, _ = _load_markdown_file(path)
    anchors = {heading.slug for heading in _parse_headings(content)}
    problems = []
    skipped_remote = 0
    for link in md_list_links(path)["links"] + [{"line": item["line"], "text": item["alt"], "target": item["src"], "type": "image"} for item in md_list_images(path)["images"]]:
        target = link["target"].strip()
        parsed = urlparse(target)
        if parsed.scheme in {"http", "https"}:
            skipped_remote += 1
            continue
        if parsed.scheme and parsed.scheme not in {"", "file"}:
            continue
        target_path = unquote(parsed.path or "")
        fragment = unquote(parsed.fragment or "")
        file_path = (resolved_path.parent / target_path).resolve() if target_path else resolved_path
        if target_path and not file_path.exists():
            problems.append({**link, "problem": "missing_file", "resolved_path": str(file_path)})
            continue
        if fragment and file_path == resolved_path and fragment not in anchors:
            problems.append({**link, "problem": "missing_anchor", "anchor": fragment})
    return {"path": str(resolved_path), "ok": not problems, "problem_count": len(problems), "problems": problems, "remote_checked": check_remote, "remote_skipped": skipped_remote if not check_remote else 0}


@mcp.tool()
def md_frontmatter(path: str, op: str = "read", data: dict[str, Any] | None = None, key: str | None = None, value: Any = None) -> dict[str, Any]:
    """Read or write simple YAML frontmatter using op: read, replace, set, delete."""
    resolved_path, content, encoding, _, size = _load_markdown_file(path)
    raw, body = _split_frontmatter(content)
    parsed = _parse_simple_yaml(raw or "")
    if op == "read":
        return {"path": str(resolved_path), "has_frontmatter": raw is not None, "frontmatter": parsed, "raw": raw or ""}
    if op == "replace":
        parsed = data or {}
    elif op == "set":
        if key is None:
            raise ValueError("key is required for set")
        parsed[key] = value
    elif op == "delete":
        if key is None:
            raise ValueError("key is required for delete")
        parsed.pop(key, None)
    else:
        raise ValueError("op must be read, replace, set, or delete")
    frontmatter = "---\n" + _render_simple_yaml(parsed) + "\n---\n"
    updated = frontmatter + body.lstrip("\n")
    write = _write_markdown_file(resolved_path, updated, encoding)
    return {"success": True, "path": str(resolved_path), **write, "bytes_before": size, "frontmatter": parsed}


@mcp.tool()
def md_tangle(path: str, output_dir: str, language: str | None = None, overwrite: bool = False) -> dict[str, Any]:
    """Extract fenced code blocks to files for lightweight literate programming."""
    resolved_path, content, _, _, _ = _load_markdown_file(path)
    blocks = md_extract_code_blocks(path, language)["blocks"] if language else md_extract_code_blocks(path)["blocks"]
    target_dir = Path(output_dir).expanduser().resolve()
    target_dir.mkdir(parents=True, exist_ok=True)
    written = []
    ext_map = {"python": ".py", "py": ".py", "javascript": ".js", "js": ".js", "typescript": ".ts", "ts": ".ts", "bash": ".sh", "sh": ".sh", "json": ".json"}
    for index, block in enumerate(blocks, start=1):
        info_parts = block["info"].split()
        target_name = None
        for part in info_parts[1:]:
            if part.startswith("file="):
                target_name = part.split("=", 1)[1].strip('"\'')
        if target_name is None:
            suffix = ext_map.get(block["language"], ".txt")
            target_name = f"block_{index}{suffix}"
        target_path = (target_dir / target_name).resolve()
        if not str(target_path).startswith(str(target_dir)):
            raise ValueError("code block output path escapes output_dir")
        if target_path.exists() and not overwrite:
            raise ValueError(f"Output file already exists: {target_path}")
        target_path.parent.mkdir(parents=True, exist_ok=True)
        target_path.write_text(block["source"] + "\n", encoding="utf-8")
        written.append({"path": str(target_path), "language": block["language"], "source_line": block["start_line"]})
    return {"path": str(resolved_path), "output_dir": str(target_dir), "written_count": len(written), "files": written}


@mcp.tool()
def md_split(path: str, output_dir: str, level: int = 1, overwrite: bool = False) -> dict[str, Any]:
    """Split a Markdown document into files at headings of a chosen level."""
    resolved_path, content, _, total_lines, _ = _load_markdown_file(path)
    headings = _parse_headings(content)
    targets = [(idx, heading) for idx, heading in enumerate(headings) if heading.level == level]
    target_dir = Path(output_dir).expanduser().resolve()
    target_dir.mkdir(parents=True, exist_ok=True)
    files = []
    for index, (heading_index, heading) in enumerate(targets, start=1):
        end_line = _section_end_line(headings, heading_index, total_lines, True)
        slug = heading.slug or f"section-{index}"
        target_path = target_dir / f"{index:02d}-{slug}.md"
        if target_path.exists() and not overwrite:
            raise ValueError(f"Output file already exists: {target_path}")
        target_path.write_text(_slice_lines(content, heading.line, end_line, False, None, None) + "\n", encoding="utf-8")
        files.append({"path": str(target_path), "heading": _heading_metadata(heading)})
    return {"path": str(resolved_path), "output_dir": str(target_dir), "file_count": len(files), "files": files}


@mcp.tool()
def md_merge(paths: list[str], output_path: str, heading_offset: int = 0, separator: str = "\n") -> dict[str, Any]:
    """Merge multiple Markdown files, optionally offsetting heading levels."""
    parts = []
    for item in paths:
        _, content, _, _, _ = _load_markdown_file(item)
        parts.append(_apply_heading_delta_to_block(content, heading_offset) if heading_offset else content)
    output = Path(output_path).expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    merged = separator.join(part.strip("\n") for part in parts) + "\n"
    output.write_text(merged, encoding="utf-8")
    return {"success": True, "path": str(output), "merged_count": len(paths), "bytes_written": output.stat().st_size}


@mcp.tool()
def md_to_html(path: str, body_only: bool = False) -> str:
    """Render a small Markdown subset to HTML without external dependencies."""
    _, content, _, _, _ = _load_markdown_file(path)
    html_lines = []
    in_code = False
    code_lines = []
    for raw_line in _to_lines(content):
        line = raw_line.rstrip("\r")
        fence = re.match(r"^```(.*)$", line)
        if fence:
            if in_code:
                html_lines.append("<pre><code>" + escape("\n".join(code_lines)) + "</code></pre>")
                code_lines = []
                in_code = False
            else:
                in_code = True
            continue
        if in_code:
            code_lines.append(line)
            continue
        heading_match = re.match(r"^(#{1,6})\s+(.+)$", line)
        if heading_match:
            level = len(heading_match.group(1))
            html_lines.append(f"<h{level} id=\"{_slugify_heading(heading_match.group(2))}\">{escape(heading_match.group(2))}</h{level}>")
        elif line.strip():
            html_lines.append(f"<p>{escape(line)}</p>")
    body = "\n".join(html_lines)
    return body if body_only else "<!doctype html>\n<html><body>\n" + body + "\n</body></html>\n"


@mcp.tool()
def md_validate_diagram(path: str, diagram_index: int = 0) -> dict[str, Any]:
    """Report whether diagram validation can run; returns skipped when the required external CLI is unavailable."""
    diagram = md_read_diagram(path, diagram_index)
    lang = diagram["diagram"]["language"]
    cli = "mmdc" if lang == "mermaid" else "plantuml" if lang in {"plantuml", "puml"} else "dot" if lang in {"dot", "graphviz"} else None
    if not cli or not shutil.which(cli):
        return {"path": diagram["path"], "ok": None, "skipped": True, "reason": f"{cli or lang} CLI is not installed", "diagram": diagram["diagram"]}
    return {"path": diagram["path"], "ok": None, "skipped": True, "reason": "CLI validation is available but not run inline by this server", "diagram": diagram["diagram"]}


def _md_render_diagram_impl(
    path: str,
    output_path: str,
    diagram_index: int,
    replace_with_image: bool,
    timeout_seconds: float,
    cancel_event: threading.Event | None,
) -> dict[str, Any]:
    diagram = md_read_diagram(path, diagram_index)
    language = diagram["diagram"]["language"]
    cli = shutil.which("mmdc")
    if language != "mermaid" or not cli:
        return {
            "path": diagram["path"],
            "ok": None,
            "skipped": True,
            "reason": "mmdc CLI is required for Mermaid rendering",
            "output_path": output_path,
        }

    output = Path(output_path).expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    command_label = (cli, "-i", "<source>", "-o", str(output))
    if cancel_event is not None and cancel_event.is_set():
        raise ManagedProcessCancelled(command_label, process_tree_stopped=True)

    with tempfile.TemporaryDirectory(
        prefix=".docloupe-mermaid-",
        dir=output.parent,
    ) as workspace_name:
        workspace = Path(workspace_name)
        source_path = workspace / "diagram.mmd"
        staged_output = workspace / output.name
        source_path.write_text(diagram["source"], encoding="utf-8")
        command = [cli, "-i", str(source_path), "-o", str(staged_output)]
        result = run_managed_process(
            command,
            timeout_seconds=timeout_seconds,
            cancel_event=cancel_event,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if cancel_event is not None and cancel_event.is_set():
            raise ManagedProcessCancelled(
                command,
                process_tree_stopped=result.process_tree_stopped,
                stdout=result.stdout,
                stderr=result.stderr,
            )
        if result.returncode != 0:
            return {
                "path": diagram["path"],
                "ok": False,
                "stderr": result.stderr,
                "exit_code": result.returncode,
                "process_tree_stopped": result.process_tree_stopped,
                "output_path": str(output),
            }
        if not staged_output.is_file():
            return {
                "path": diagram["path"],
                "ok": False,
                "stderr": "mmdc exited successfully but produced no output file",
                "exit_code": result.returncode,
                "process_tree_stopped": result.process_tree_stopped,
                "output_path": str(output),
            }
        if cancel_event is not None and cancel_event.is_set():
            raise ManagedProcessCancelled(command, process_tree_stopped=True)
        os.replace(staged_output, output)

    if replace_with_image:
        resolved_path, content, encoding, _, _ = _load_markdown_file(path)
        block = _find_fenced_blocks(
            content,
            {"mermaid", "plantuml", "puml", "dot", "graphviz"},
        )[diagram_index]
        image_markdown = f"![diagram]({output.name})\n"
        updated = _replace_line_range(
            content,
            block["start_line"],
            block["end_line"],
            image_markdown,
        )
        _write_markdown_file(resolved_path, updated, encoding)
    return {
        "path": diagram["path"],
        "ok": True,
        "process_tree_stopped": True,
        "output_path": str(output),
    }


def md_render_diagram(
    path: str,
    output_path: str,
    diagram_index: int = 0,
    replace_with_image: bool = False,
    timeout_seconds: float = 60.0,
) -> dict[str, Any]:
    """Render Mermaid through a bounded mmdc/Chromium process tree."""
    return _md_render_diagram_impl(
        path,
        output_path,
        diagram_index,
        replace_with_image,
        timeout_seconds,
        None,
    )


@mcp.tool(name="md_render_diagram")
async def _md_render_diagram_tool(
    path: str,
    output_path: str,
    diagram_index: int = 0,
    replace_with_image: bool = False,
    timeout_seconds: float = 60.0,
) -> dict[str, Any]:
    """Render Mermaid; cancellation stops mmdc and browser descendants."""
    return await run_cancellable_in_thread(
        lambda cancel_event: _md_render_diagram_impl(
            path,
            output_path,
            diagram_index,
            replace_with_image,
            timeout_seconds,
            cancel_event,
        )
    )


async def _run_frozen_stdio_server() -> None:
    import anyio
    from io import TextIOWrapper
    from mcp.server.stdio import stdio_server

    stdin_text = TextIOWrapper(sys.stdin.buffer, encoding="utf-8", errors="replace")
    stdout_text = TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    try:
        async with stdio_server(
            stdin=anyio.wrap_file(stdin_text),
            stdout=anyio.wrap_file(stdout_text),
        ) as (read_stream, write_stream):
            await mcp._mcp_server.run(
                read_stream,
                write_stream,
                mcp._mcp_server.create_initialization_options(),
            )
    finally:
        for stream in (stdin_text, stdout_text):
            try:
                stream.detach()
            except (OSError, ValueError):
                pass


if __name__ == "__main__":
    if getattr(sys, "frozen", False):
        import anyio
        anyio.run(_run_frozen_stdio_server)
    else:
        mcp.run(transport="stdio")



