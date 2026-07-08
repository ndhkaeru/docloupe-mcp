---
name: release
description: Prepare a DocForge MCP release: validate Python tests, build native server executables for Windows, Linux, and macOS, update MCP metadata, publish GitHub Release artifacts, and smoke test real document workflows.
---

# Preparing a DocForge MCP Release

Use this checklist when cutting a `docforge-mcp` release. DocForge ships focused stdio MCP servers as native executables; do not publish Docker or npm artifacts unless that distribution channel is explicitly added.

## 1. Pick the version

- Review recent releases: `gh release list --repo ndhkaeru/docforge-mcp --limit 5`.
- Review user-visible changes: `git log <previous-tag>..HEAD --oneline`.
- Choose a semver tag such as `v0.1.1`.

## 2. Bump versioned metadata

Update these fields together:

- `server.json` top-level `version`
- `server.json` package entry `version`
- README release examples or release notes if they mention a fixed version

Keep `server.json#name` as `io.github.ndhkaeru/docforge-mcp` unless the repository identity changes.

## 3. Validate before release

From `docforge-mcp/`:

```powershell
python -m venv .venv-build
.\.venv-build\Scripts\python.exe -m pip install -U pip
.\.venv-build\Scripts\python.exe -m pip install pytest pyinstaller "mcp[cli]" openpyxl pillow pdfminer.six pdfplumber mammoth python-pptx markdownify beautifulsoup4 lxml defusedxml charset-normalizer
.\.venv-build\Scripts\python.exe -m pytest -q
.\build.ps1 -Only excel
```

For Excel-related releases, run at least one real-file smoke test:

- load a real `.xlsx`
- edit cells/styles and a DrawingML shape if shape code changed
- `excel_save_as_copy`
- open the output with Excel COM when available

## 4. Build all artifacts

```powershell
.\build.ps1
# GitHub Actions prepares platform release artifacts from dist/ during release builds.
```

Expected release artifacts per platform (`windows-x64`, `linux-x64`, `macos-x64`, `macos-arm64`):

- `docforge-mcp-md-tools-<platform>[.exe]`
- `docforge-mcp-excel-tools-<platform>[.exe]`
- `docforge-mcp-pdf-tools-<platform>[.exe]`
- `docforge-mcp-docx-tools-<platform>[.exe]`
- `docforge-mcp-pptx-tools-<platform>[.exe]`
- `docforge-mcp-csv-tools-<platform>[.exe]`
- `docforge-mcp-html-tools-<platform>[.exe]`
- `docforge-mcp-text-tools-<platform>[.exe]`
- `docforge-mcp-json-tools-<platform>[.exe]`
- `docforge-mcp-vX.Y.Z-<platform>-sha256sums.txt`

## 5. Push and publish

- Push `main` and confirm `Push` workflow succeeds.
- Push the semver tag: `git tag vX.Y.Z && git push origin vX.Y.Z`.
- Confirm `Release` workflow succeeds for every platform and attaches every executable plus checksums.
- Download at least `excel-tools` and `md-tools` for one target platform and run a smoke test.

## 6. Release notes

Use short human-facing notes:

- `## What's New` for new tools or formats
- `## Improvements` for behavior, preservation, performance, or packaging
- `## Fixes` for corruption, save, validation, or security fixes
- `## Validation` for real-file/Excel COM evidence when relevant

Mention best-effort limitations clearly for advanced Excel parts such as external links, pivot tables, slicers, VML, and macros.

