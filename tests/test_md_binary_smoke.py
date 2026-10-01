"""End-to-end checks for a packaged md-tools binary with its renderer."""

import os
from pathlib import Path

import anyio
import pytest
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


def test_packaged_md_renderer_over_stdio(tmp_path):
    binary = os.environ.get("DOCLOUPE_MD_BINARY")
    if not binary:
        pytest.skip("set DOCLOUPE_MD_BINARY to test the packaged server")
    assert Path(binary).is_file()

    async def check():
        async with stdio_client(StdioServerParameters(command=binary)) as streams:
            async with ClientSession(*streams) as session:
                await session.initialize()
                tools = {tool.name for tool in (await session.list_tools()).tools}
                assert {"md_runtime_capabilities", "md_validate_diagram", "md_render_diagram"} <= tools
                capabilities = await session.call_tool("md_runtime_capabilities")
                assert not capabilities.isError
                for language in ("mermaid", "dot"):
                    assert capabilities.structuredContent["diagram"][language]["backend"] == "rust"

                for language, source, invalid in (
                    ("mermaid", "flowchart TD\n A --> B", "flowchart TD\n A -->"),
                    ("dot", "digraph G { A -> B; }", "digraph G {"),
                ):
                    document = tmp_path / f"{language}.md"
                    output = tmp_path / f"{language}.svg"
                    document.write_text(f"```{language}\n{source}\n```\n", encoding="utf-8")
                    validation = await session.call_tool("md_validate_diagram", {"path": str(document)})
                    assert not validation.isError and validation.structuredContent["ok"] is True
                    render = await session.call_tool("md_render_diagram", {
                        "path": str(document), "output_path": str(output),
                    })
                    assert not render.isError and render.structuredContent["ok"] is True
                    assert "<svg" in output.read_text(encoding="utf-8")
                    original = output.read_bytes()
                    document.write_text(f"```{language}\n{invalid}\n```\n", encoding="utf-8")
                    validation = await session.call_tool("md_validate_diagram", {"path": str(document)})
                    assert not validation.isError and validation.structuredContent["ok"] is False
                    render = await session.call_tool("md_render_diagram", {
                        "path": str(document), "output_path": str(output),
                    })
                    assert not render.isError and render.structuredContent["ok"] is False
                    assert output.read_bytes() == original

    async def bounded_check():
        with anyio.fail_after(90):
            await check()

    anyio.run(bounded_check)
