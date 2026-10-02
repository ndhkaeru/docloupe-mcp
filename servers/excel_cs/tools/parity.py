"""Bounded Python/C# cell-edit parity scenario, checked against OOXML directly."""

import asyncio
import posixpath
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from urllib.parse import unquote

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


ROOT = Path(__file__).resolve().parents[3]
EXCEL = ROOT / "servers" / "excel"
SERVER = ROOT / "servers" / "excel_cs"
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
OFFICE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE = "http://schemas.openxmlformats.org/package/2006/relationships"
VALUES = {"B1": 27, "C1": 9, "D3": "new text", "E5": True,
          "F6": 0.1, "G7": 4, "H7": 4, "I8": 12, "J8": "batch"}
VARIANTS = ("default", "prefixed-x", "bom-crlf-standalone", "opc-percent-case",
            "new-shared-strings", "nested-workbook")


def dotnet(*arguments):
    subprocess.run(["dotnet", *arguments], cwd=ROOT, check=True, stdout=subprocess.DEVNULL)


def read_cells(path):
    """Resolve OPC relationships and read cells without either implementation's locator."""
    with zipfile.ZipFile(path) as archive:
        entries = {name.lower(): name for name in archive.namelist()}

        def part(name):
            return archive.read(entries[name.lower()])

        def relationships(name):
            base = "" if name == "_rels/.rels" else posixpath.dirname(posixpath.dirname(name))
            return {node.attrib["Id"]: (node.attrib["Type"].split("/")[-1],
                    posixpath.normpath(posixpath.join(base, unquote(node.attrib["Target"]).lstrip("/")))
                    if not node.attrib["Target"].startswith("/") else unquote(node.attrib["Target"]).lstrip("/"))
                    for node in ET.fromstring(part(name)).findall(f"{{{PACKAGE}}}Relationship")
                    if node.attrib.get("TargetMode") != "External"}

        workbook = next(target for kind, target in relationships("_rels/.rels").values()
                        if kind == "officeDocument")
        book = ET.fromstring(part(workbook))
        rels = relationships(posixpath.join(posixpath.dirname(workbook), "_rels",
                                          posixpath.basename(workbook) + ".rels"))
        sheet = next(node for node in book.iter(f"{{{MAIN}}}sheet") if node.attrib["name"] == "Sheet1")
        worksheet = ET.fromstring(part(rels[sheet.attrib[f"{{{OFFICE}}}id"]][1]))
        strings = []
        for kind, target in rels.values():
            if kind == "sharedStrings":
                for item in ET.fromstring(part(target)).findall(f"{{{MAIN}}}si"):
                    strings.append("".join(node.text or "" for node in
                                           item.findall(f"{{{MAIN}}}t") + item.findall(f"{{{MAIN}}}r/{{{MAIN}}}t")))
        result = {}
        for cell in worksheet.iter(f"{{{MAIN}}}c"):
            formula = cell.find(f"{{{MAIN}}}f")
            value = cell.find(f"{{{MAIN}}}v")
            kind = cell.attrib.get("t", "n")
            if kind == "s":
                scalar = strings[int(value.text)] if value is not None and value.text else None
            elif kind == "inlineStr":
                scalar = "".join(node.text or "" for node in cell.findall(f"{{{MAIN}}}is//{{{MAIN}}}t"))
            elif kind == "b":
                scalar = value is not None and value.text == "1"
            elif kind == "n" and value is not None and value.text is not None:
                scalar = float(value.text)
            else:
                scalar = value.text if value is not None else None
            result[cell.attrib["r"]] = (scalar, formula.text if formula is not None else None)
        return result


def check_output(source, output):
    original = read_cells(source)
    actual = read_cells(output)
    errors = []
    for address, expected in VALUES.items():
        value, formula = actual.get(address, (None, None))
        if value != expected or type(value) is not type(float(expected) if type(expected) in (int, float) else expected):
            errors.append(f"{address}: expected {expected!r}, got {value!r}")
        if formula is not None:
            errors.append(f"{address}: stale formula {formula!r}")
    for address in original.keys() - VALUES.keys():
        if original[address] != actual.get(address):
            errors.append(f"{address}: unrelated cell changed: {original[address]!r} -> {actual.get(address)!r}")
    return errors


async def save_new(source, output, dll):
    parameters = StdioServerParameters(command="dotnet", args=[str(dll)], cwd=str(ROOT))
    async with stdio_client(parameters) as (reader, writer):
        async with ClientSession(reader, writer) as client:
            await client.initialize()

            async def call(name, arguments):
                response = await client.call_tool(name, arguments)
                data = response.structuredContent
                if response.isError or data is None or not data.get("ok"):
                    raise RuntimeError(f"{name}: {data}")
                return data["data"]

            opened = await call("excel_open", {"path": str(source)})
            session = opened["session"]
            await call("excel_apply", {"session": session, "base_revision": 0, "sheet": "Sheet1", "ops": [
                *({"op": "set_value", "target": address, "value": VALUES[address]}
                  for address in ("B1", "C1", "D3", "E5", "F6")),
                {"op": "fill", "target": "G7:H7", "value": 4},
                {"op": "set_values", "target": "I8:J8", "values": [[12, "batch"]]},
            ]})
            await call("excel_save", {"session": session, "mode": "copy", "path": str(output)})
            await call("excel_close", {"session": session, "discard_unsaved": True})


def save_legacy(source, output):
    sys.path.insert(0, str(EXCEL))
    import main

    main.excel_load(str(source))
    session = str(source.resolve())
    main.excel_edit_cells(session, "Sheet1", [{"cell": address, "value": value}
                                              for address, value in VALUES.items()])
    main.excel_save_as_copy(session, str(output), report_format="json", verify_preservation=False)
    main.excel_close(session)


async def main():
    dotnet("build", str(SERVER / "src/DocLoupe.Excel.Server/DocLoupe.Excel.Server.csproj"), "-c", "Release", "--nologo")
    dll = SERVER / "src/DocLoupe.Excel.Server/bin/Release/net10.0/DocLoupe.Excel.Server.dll"
    with tempfile.TemporaryDirectory(prefix="docloupe-parity-") as temporary:
        directory = Path(temporary)
        dotnet("run", "--project", str(SERVER / "fixtures/DocLoupe.Excel.Fixtures.csproj"),
               "-c", "Release", "--", str(directory))
        failures = []
        for variant in VARIANTS:
            source = directory / f"{variant}.xlsx"
            current = directory / f"{variant}-cs.xlsx"
            legacy = directory / f"{variant}-python.xlsx"
            await save_new(source, current, dll)
            current_errors = check_output(source, current)
            if current_errors:
                failures.append(f"{variant} C#: {current_errors}")
            try:
                save_legacy(source, legacy)
                legacy_errors = check_output(source, legacy)
            except (KeyError, ValueError) as error:
                legacy_errors = [f"{type(error).__name__}: {error}"]
            if variant == "opc-percent-case":
                expected_errors = ["ValueError: Sheet 'Sheet1' not found. Available: []"]
            elif variant == "new-shared-strings":
                expected_errors = []
            elif variant in ("prefixed-x", "nested-workbook"):
                expected_errors = ["A1: unrelated cell changed: ('hello', None) -> None"]
            else:
                expected_errors = ["A1: unrelated cell changed: ('hello', None) -> ('hellohe', None)"]
            if legacy_errors != expected_errors:
                failures.append(f"{variant} Python: expected {expected_errors}, got {legacy_errors}")
            print(f"{variant}: C# {'PASS' if not current_errors else 'FAIL'}; "
                  f"Python {'PASS' if not legacy_errors else 'known divergence: ' + str(legacy_errors)}")
        if failures:
            raise AssertionError("\n".join(failures))
    print("Parity scenario PASS: 6 C# outputs, 1 comparable Python output, 5 recorded divergences")


if __name__ == "__main__":
    asyncio.run(main())
