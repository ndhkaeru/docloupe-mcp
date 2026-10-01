"""Check the packaged Excel MCP server and its embedded release metadata."""

import json
import os
import shutil
from datetime import timedelta
from pathlib import Path

import anyio
import openpyxl
import pytest
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


def test_packaged_excel_verification_over_stdio(tmp_path):
    binary = os.environ.get("DOCLOUPE_EXCEL_BINARY")
    if not binary:
        pytest.skip("set DOCLOUPE_EXCEL_BINARY to test the packaged server")
    assert Path(binary).is_file()

    before = tmp_path / "before.xlsx"
    after = tmp_path / "after.xlsx"
    workbook = openpyxl.Workbook()
    workbook.active["A1"] = "verified"
    workbook.save(before)
    shutil.copyfile(before, after)

    async def check():
        async with stdio_client(StdioServerParameters(command=binary)) as streams:
            async with ClientSession(*streams) as session:
                await session.initialize()
                assert "excel_verify_preservation" in {
                    tool.name for tool in (await session.list_tools()).tools
                }
                result = await session.call_tool("excel_verify_preservation", {
                    "after_path": str(after), "before_path": str(before),
                }, read_timeout_seconds=timedelta(seconds=60))
                assert not result.isError
                report = json.loads(result.content[0].text)
                assert report["preservation_ok"] is True
                assert report["runtime"]["server_version"] == json.loads(
                    (Path(__file__).resolve().parents[1] / "server.json").read_text(encoding="utf-8")
                )["version"]

    async def bounded_check():
        with anyio.fail_after(90):
            await check()

    anyio.run(bounded_check)
