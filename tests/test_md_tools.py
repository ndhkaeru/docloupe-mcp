"""Tests for md-tools (servers/md/main.py).

Focus on the two correctness fixes and the trimmed tool surface:

S.  Unicode-correct, GitHub-compatible heading slugs (Vietnamese etc.).
L.  A single '\\n'-based line model so heading line numbers stay consistent
    across LF / CRLF / lone-CR / no-trailing-newline files.
F.  Fenced code blocks are not mistaken for headings.
T.  The five kept tools work; the four removed tools are gone.
"""
import importlib.util
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

ROOT = Path(__file__).resolve().parents[1]
_spec = importlib.util.spec_from_file_location("md_tools_main", ROOT / "servers" / "md" / "main.py")
M = importlib.util.module_from_spec(_spec)
sys.modules["md_tools_main"] = M
_spec.loader.exec_module(M)


def _write(tmp_path: Path, content: str, name: str = "doc.md") -> str:
    p = tmp_path / name
    p.write_text(content, encoding="utf-8", newline="")
    return str(p)


# --- S. Unicode slugs -------------------------------------------------------

@pytest.mark.parametrize("title, expected", [
    ("Cài đặt", "cài-đặt"),
    ("Tổng quan", "tổng-quan"),
    ("Hướng dẫn sử dụng", "hướng-dẫn-sử-dụng"),
    ("1.2 Mục tiêu", "12-mục-tiêu"),
    ("Hello, World!", "hello-world"),
    ("C++ Guide", "c-guide"),
    ("foo_bar", "foo_bar"),       # underscore kept
    ("a  b", "a--b"),             # hyphens not collapsed
    ("  Trim me  ", "trim-me"),
])
def test_slugify_matches_github(title, expected):
    assert M._slugify_heading(title) == expected


# --- L. Consistent line model ----------------------------------------------

def _fixture_lines():
    return [
        "# Tài liệu",                  # 1
        "",                            # 2
        "Giới thiệu.",                 # 3
        "",                            # 4
        "## Cài đặt",                  # 5
        "",                            # 6
        "Bước cài đặt.",               # 7
        "",                            # 8
        "## Tổng quan",                # 9
        "",                            # 10
        "```python",                   # 11
        "# day khong phai heading",    # 12 (inside fence)
        "print('x')",                  # 13
        "```",                         # 14
        "",                            # 15
        "Nội dung tổng quan.",         # 16
    ]


@pytest.mark.parametrize("newline", ["\n", "\r\n"])
def test_outline_line_numbers_stable_across_newline_styles(tmp_path, newline):
    content = newline.join(_fixture_lines()) + newline
    path = _write(tmp_path, content)
    out = M.markdown_outline(path)
    assert out["total_lines"] == 16
    assert out["heading_count"] == 3  # fenced "# ..." is not a heading
    titles = [(h["title"], h["level"], h["line"], h["slug"]) for h in out["headings"]]
    assert titles == [
        ("Tài liệu", 1, 1, "tài-liệu"),
        ("Cài đặt", 2, 5, "cài-đặt"),
        ("Tổng quan", 2, 9, "tổng-quan"),
    ]


def test_lone_cr_does_not_desync_line_numbers(tmp_path):
    # A bare CR inside a paragraph must NOT be treated as a line break,
    # otherwise the heading below it would be reported on the wrong line.
    path = _write(tmp_path, "# H1\nfoo\rbar\n## H2\n")
    out = M.markdown_outline(path)
    assert out["total_lines"] == 3
    assert [(h["title"], h["line"]) for h in out["headings"]] == [("H1", 1), ("H2", 3)]


def test_read_section_strips_cr_and_no_trailing_newline(tmp_path):
    content = "\r\n".join(_fixture_lines())  # CRLF, and no trailing newline
    path = _write(tmp_path, content)
    section = M.read_markdown_section(path, heading="Cài đặt")
    assert "\r" not in section
    assert section.startswith("## Cài đặt")
    assert "Bước cài đặt." in section
    assert "Tổng quan" not in section  # stops before the next same-level heading


# --- F. read by path + edits ------------------------------------------------

def test_read_section_by_heading_path(tmp_path):
    content = "# A\n\n## Setup\n\nfrom A\n\n# B\n\n## Setup\n\nfrom B\n"
    path = _write(tmp_path, content)
    section = M.read_markdown_section(path, heading_path=["B", "Setup"], include_heading=False)
    assert "from B" in section
    assert "from A" not in section

def test_targeted_read_and_document_map_tools(tmp_path):
    path = _write(tmp_path, "# A\n\nintro\n\n## Data\n\n| K | V |\n|---|---|\n| a | b |\n\n[link](target.md)\n")
    assert "2 |" in M.md_read_range(path, 2, 3, include_line_numbers=True)
    near = M.md_read_near(path, text="Data", before=1, after=2)
    assert "## Data" in near and "| K | V |" in near
    doc_map = M.md_get_document_map(path)
    assert doc_map["heading_count"] == 2
    assert doc_map["table_count"] == 1
    assert doc_map["link_count"] == 1


def test_replace_section_body_only(tmp_path):
    path = _write(tmp_path, "# A\n\nintro\n\n## B\n\nold body\n")
    res = M.replace_markdown_section(path, "new body", heading="B")
    assert res["changed"] is True
    text = Path(path).read_text(encoding="utf-8")
    assert "## B" in text and "new body" in text and "old body" not in text
    assert "# A" in text  # untouched


def test_insert_section_under_parent(tmp_path):
    path = _write(tmp_path, "# A\n\nintro\n")
    res = M.md_insert_section(path, title="Details", content="hello", parent_heading="A")
    assert res["changed"] is True
    assert res["inserted_heading"]["level"] == 2
    text = Path(path).read_text(encoding="utf-8")
    assert "## Details" in text and "hello" in text


def test_search_returns_heading_context_without_cr(tmp_path):
    content = "\r\n".join(["## Cài đặt", "", "Bước cài đặt ngay."])
    path = _write(tmp_path, content)
    res = M.md_search(path, "cài đặt")
    assert res["match_count"] >= 1
    hit = res["matches"][-1]
    assert "\r" not in hit["text"]
    assert hit["heading"]["title"] == "Cài đặt"

def test_replace_text_dry_run_does_not_write(tmp_path):
    path = _write(tmp_path, "# A\n\nalpha alpha\n")
    res = M.md_replace_text(path, "alpha", "beta", dry_run=True)
    assert res["dry_run"] is True
    assert res["replacement_count"] == 2
    assert "alpha alpha" in Path(path).read_text(encoding="utf-8")

def test_patch_lines_normalize_headings_and_internal_links(tmp_path):
    (tmp_path / "ok.md").write_text("# OK\n", encoding="utf-8")
    path = _write(tmp_path, "# A\n\n#### Too Deep\n\n[ok](ok.md)\n[bad](missing.md)\n")
    patch = M.md_patch_lines(path, 3, 3, "## Fixed")
    assert patch["changed"] is True
    text = Path(path).read_text(encoding="utf-8")
    assert "## Fixed" in text

    path2 = _write(tmp_path, "# A\n\n#### Too Deep\n", name="levels.md")
    dry = M.md_normalize_headings(path2, dry_run=True)
    assert dry["change_count"] == 1
    assert "#### Too Deep" in Path(path2).read_text(encoding="utf-8")
    normalized = M.md_normalize_headings(path2)
    assert normalized["changed"] is True
    assert "## Too Deep" in Path(path2).read_text(encoding="utf-8")

    links = M.md_check_internal_links(path)
    assert links["problem_count"] == 1
    assert links["remote_checked"] is False


# --- T. Trimmed tool surface ------------------------------------------------

def test_kept_tools_present():
    for name in ["markdown_outline", "read_markdown_section", "replace_markdown_section", "md_search", "md_insert_section",
                 "md_read_range", "md_read_near", "md_get_document_map", "md_patch_lines", "md_normalize_headings", "md_check_internal_links", "md_move_section", "md_set_heading_level", "md_validate_links", "md_frontmatter", "md_split", "md_merge"]:
        assert callable(getattr(M, name, None)), f"missing tool {name}"


def test_removed_tools_absent():
    for name in ["md_read_file", "_render_frontmatter"]:
        assert not hasattr(M, name), f"{name} should have been removed"


def test_section_helpers_delete_append_and_rename(tmp_path):
    path = _write(tmp_path, "# A\n\nintro\n\n## B\n\nold\n")
    assert M.md_append_to_section(path, "more", heading="B")["changed"] is True
    assert "more" in Path(path).read_text(encoding="utf-8")
    assert M.md_rename_heading(path, "Bee", heading="B")["new_heading"]["title"] == "Bee"
    assert M.md_delete_section(path, heading="Bee")["changed"] is True
    assert "## Bee" not in Path(path).read_text(encoding="utf-8")


def test_table_tools_read_format_and_edit(tmp_path):
    path = _write(tmp_path, "# Data\n\n| B | A |\n|---|---:|\n| 2 | x |\n| 1 | y |\n")
    listed = M.md_list_tables(path)
    assert listed["table_count"] == 1
    assert listed["tables"][0]["headers"] == ["B", "A"]
    table = M.md_read_table(path)
    assert table["rows"] == [["2", "x"], ["1", "y"]]
    assert M.md_edit_table(path, op="sort", col="B")["changed"] is True
    assert "| 1" in Path(path).read_text(encoding="utf-8")
    assert M.md_format_table(path)["success"] is True


def test_diagram_code_link_toc_and_stats_tools(tmp_path):
    path = _write(tmp_path, "# A\n\n[OpenAI](https://openai.com)\n\n![Alt](img.png)\n\n```mermaid\ngraph TD; A-->B\n```\n")
    assert M.md_list_links(path)["link_count"] == 1
    assert M.md_list_images(path)["images"][0]["src"] == "img.png"
    assert M.md_list_diagrams(path)["diagram_count"] == 1
    assert "A-->B" in M.md_read_diagram(path)["source"]
    assert M.md_replace_diagram(path, "graph TD; B-->C")["changed"] is True
    assert M.md_extract_code_blocks(path, language="mermaid")["block_count"] == 1
    assert M.md_get_anchor("Cài đặt")["href"] == "#cài-đặt"
    assert M.md_update_toc(path)["heading_count"] >= 1
    assert M.md_stats(path)["code_blocks"] == 1

def test_large_read_tools_support_preview_and_metadata_only(tmp_path):
    path = _write(tmp_path, "# A\n\n| H |\n|---|\n| 1 |\n| 2 |\n\n```python\n" + "x = 1\n" * 200 + "```\n")
    section = M.read_markdown_section(path, heading="A", preview=True)
    assert "truncated" in section
    table = M.md_read_table(path, include_body=False)
    assert "rows" not in table
    assert table["headers"] == ["H"]
    blocks = M.md_extract_code_blocks(path, include_body=False)
    assert "source" not in blocks["blocks"][0]
    preview = M.md_extract_code_blocks(path, max_chars=10)
    assert "truncated" in preview["blocks"][0]["source"]


def test_move_and_set_heading_level(tmp_path):
    path = _write(tmp_path, "# A\n\n## B\n\n### C\n\nbody\n\n# D\n\nend\n")
    assert M.md_set_heading_level(path, level=3, heading="B")["changed"] is True
    text = Path(path).read_text(encoding="utf-8")
    assert "### B" in text and "#### C" in text
    assert M.md_move_section(path, heading="B", target_heading="D", position="after")["changed"] is True
    text = Path(path).read_text(encoding="utf-8")
    assert text.index("# D") < text.index("### B")


def test_frontmatter_link_validate_rewrite_and_html(tmp_path):
    (tmp_path / "target.md").write_text("# Target\n", encoding="utf-8")
    path = _write(tmp_path, "---\ntitle: Old\n---\n# A\n\n[bad](missing.md)\n[ok](target.md)\n[anchor](#missing)\n")
    assert M.md_frontmatter(path)["frontmatter"]["title"] == "Old"
    assert M.md_frontmatter(path, op="set", key="title", value="New")["frontmatter"]["title"] == "New"
    validation = M.md_validate_links(path)
    assert validation["problem_count"] == 2
    assert M.md_rewrite_links(path, {"missing.md": "target.md"})["replacement_count"] == 1
    assert "<h1" in M.md_to_html(path, body_only=True)


def test_tangle_split_merge_and_diagram_skip(tmp_path):
    path = _write(tmp_path, "# A\n\n```python file=src/app.py\nprint('hi')\n```\n\n# B\n\ntext\n")
    tangled = M.md_tangle(path, str(tmp_path / "out"))
    assert tangled["written_count"] == 1
    assert (tmp_path / "out" / "src" / "app.py").exists()
    split = M.md_split(path, str(tmp_path / "parts"), level=1)
    assert split["file_count"] == 2
    merged_path = tmp_path / "merged.md"
    assert M.md_merge([item["path"] for item in split["files"]], str(merged_path), heading_offset=1)["merged_count"] == 2
    diagram_path = _write(tmp_path, "# D\n\n```mermaid\ngraph TD; A-->B\n```\n", name="diagram.md")
    assert "skipped" in M.md_validate_diagram(diagram_path)


def test_mermaid_render_uses_staging_and_removes_workspace(tmp_path, monkeypatch):
    markdown_path = _write(
        tmp_path,
        "# Diagram\n\n```mermaid\ngraph TD; A-->B\n```\n",
        name="diagram.md",
    )
    output_path = tmp_path / "diagram.svg"
    workspaces = []

    monkeypatch.setattr(M.shutil, "which", lambda _name: "fake-mmdc")

    def fake_run(command, **kwargs):
        assert kwargs["timeout_seconds"] == 5.0
        source_path = Path(command[command.index("-i") + 1])
        staged_output = Path(command[command.index("-o") + 1])
        assert source_path.read_text(encoding="utf-8") == "graph TD; A-->B"
        workspaces.append(source_path.parent)
        staged_output.write_text("<svg>ok</svg>", encoding="utf-8")
        return SimpleNamespace(
            returncode=0,
            stdout="",
            stderr="",
            process_tree_stopped=True,
        )

    monkeypatch.setattr(M, "run_managed_process", fake_run)
    result = M.md_render_diagram(
        markdown_path,
        str(output_path),
        timeout_seconds=5.0,
    )

    assert result["ok"] is True
    assert result["process_tree_stopped"] is True
    assert output_path.read_text(encoding="utf-8") == "<svg>ok</svg>"
    assert workspaces and all(not workspace.exists() for workspace in workspaces)
    assert not (tmp_path / "diagram.mmd").exists()
    assert not list(tmp_path.glob(".docloupe-mermaid-*"))


@pytest.mark.parametrize("source", [
    b"# A\n\nbody", b"# A\r\n\r\nbody",
    b"\xef\xbb\xbf# A\r\n\r\nbody",
])
def test_append_preserves_representation_and_separates_eof(tmp_path, source):
    path = tmp_path / "doc.md"
    path.write_bytes(source)
    M.md_append_to_section(str(path), "added", heading="A")
    result = path.read_bytes()
    assert result.startswith(b"\xef\xbb\xbf") == source.startswith(b"\xef\xbb\xbf")
    assert not result.endswith(b"\n")
    assert b"bodyadded" not in result
    assert b"body\r\nadded" in result if b"\r\n" in source else b"body\nadded" in result
    if b"\r\n" in source:
        assert b"\n" not in result.replace(b"\r\n", b"")


def test_tangle_rejects_parent_and_symlink_escapes(tmp_path):
    destination = tmp_path / "out"
    outside = tmp_path / "outside"
    outside.mkdir()
    path = _write(tmp_path, "```python file=../outside/pwn.txt\n1\n```\n")
    with pytest.raises(ValueError, match="escapes"):
        M.md_tangle(path, str(destination))
    assert not (outside / "pwn.txt").exists()
    try:
        (destination / "linked").symlink_to(outside, target_is_directory=True)
    except (OSError, NotImplementedError):
        pytest.skip("Directory symlinks unavailable")
    Path(path).write_text("```python file=linked/pwn.txt\n1\n```\n", encoding="utf-8")
    with pytest.raises(ValueError, match="escapes"):
        M.md_tangle(path, str(destination))
    assert not (outside / "pwn.txt").exists()


def test_table_pipes_and_single_cell_edit_leave_other_rows_unchanged(tmp_path):
    original = "| First  | Second |\n| :---- | ---: |\n| a\\|b | `x|y` |\n| keep   | as-is |\n"
    path = _write(tmp_path, original)
    assert M.md_read_table(path)["rows"] == [[r"a\|b", "`x|y`"], ["keep", "as-is"]]
    M.md_edit_table(path, op="set_cell", row=0, col=0, value="new|value")
    assert Path(path).read_text(encoding="utf-8") == original.replace(r"a\|b", r"new\|value")
    M.md_format_table(path)
    assert M.md_read_table(path)["rows"][0] == [r"new\|value", "`x|y`"]


def test_complex_frontmatter_is_rejected_without_mutation(tmp_path):
    path = _write(tmp_path, "---\n# comment\nitems:\n  - value\n---\n# Body\n")
    before = Path(path).read_bytes()
    with pytest.raises(ValueError, match="Complex YAML"):
        M.md_frontmatter(path, op="set", key="title", value="new")
    assert Path(path).read_bytes() == before


def test_links_skip_code_and_keep_parenthesized_destinations(tmp_path):
    path = _write(tmp_path, "`[example](missing.md)`\n```md\n[example](missing.md)\n```\n[valid](https://example.com/a(b))\n![img](image(a).png)\n")
    assert [item["target"] for item in M.md_list_links(path)["links"]] == ["https://example.com/a(b)"]
    assert [item["src"] for item in M.md_list_images(path)["images"]] == ["image(a).png"]
    validation = M.md_validate_links(path, check_remote=True)
    assert validation["remote_checked"] is False and validation["remote_skipped"] == 1
    M.md_rewrite_links(path, {"https://example.com/a(b)": "new.md"})
    assert "[valid](new.md)" in Path(path).read_text(encoding="utf-8")
    assert "`[example](missing.md)`" in Path(path).read_text(encoding="utf-8")


def test_mixed_newlines_reject_mutation_without_writing(tmp_path):
    path = tmp_path / "mixed.md"
    original = b"# A\r\n\nbody\r\n"
    path.write_bytes(original)
    with pytest.raises(ValueError, match="Mixed line endings"):
        M.md_append_to_section(str(path), "more", heading="A")
    assert path.read_bytes() == original


def test_diagram_validation_runs_installed_cli(tmp_path, monkeypatch):
    path = _write(tmp_path, "```mermaid\ngraph TD; A-->B\n```\n")
    monkeypatch.setattr(M.shutil, "which", lambda _name: "fake-mmdc")

    def fake_run(command, **_kwargs):
        Path(command[command.index("-o") + 1]).write_text("<svg/>", encoding="utf-8")
        return SimpleNamespace(returncode=0, stderr="")

    monkeypatch.setattr(M, "run_managed_process", fake_run)
    result = M.md_validate_diagram(path)
    assert result["ok"] is True and result["skipped"] is False


def test_render_image_link_is_relative_to_document(tmp_path, monkeypatch):
    path = _write(tmp_path, "```mermaid\ngraph TD; A-->B\n```\n")
    output = tmp_path / "images" / "my chart.svg"
    monkeypatch.setattr(M.shutil, "which", lambda _name: "fake-mmdc")

    def fake_run(command, **_kwargs):
        Path(command[command.index("-o") + 1]).write_text("<svg/>", encoding="utf-8")
        return SimpleNamespace(returncode=0, stderr="", process_tree_stopped=True)

    monkeypatch.setattr(M, "run_managed_process", fake_run)
    assert M.md_render_diagram(path, str(output), replace_with_image=True)["ok"] is True
    assert "![diagram](images/my%20chart.svg)" in Path(path).read_text(encoding="utf-8")


def test_mermaid_render_failure_keeps_existing_output(tmp_path, monkeypatch):
    markdown_path = _write(
        tmp_path,
        "# Diagram\n\n```mermaid\ngraph TD; A-->B\n```\n",
        name="diagram.md",
    )
    output_path = tmp_path / "diagram.svg"
    output_path.write_text("existing", encoding="utf-8")
    workspaces = []

    monkeypatch.setattr(M.shutil, "which", lambda _name: "fake-mmdc")

    def fake_run(command, **_kwargs):
        staged_output = Path(command[command.index("-o") + 1])
        workspaces.append(staged_output.parent)
        staged_output.write_text("partial", encoding="utf-8")
        return SimpleNamespace(
            returncode=12,
            stdout="",
            stderr="render failed",
            process_tree_stopped=True,
        )

    monkeypatch.setattr(M, "run_managed_process", fake_run)
    result = M.md_render_diagram(markdown_path, str(output_path))

    assert result["ok"] is False
    assert result["exit_code"] == 12
    assert output_path.read_text(encoding="utf-8") == "existing"
    assert workspaces and all(not workspace.exists() for workspace in workspaces)
    assert not list(tmp_path.glob(".docloupe-mermaid-*"))
