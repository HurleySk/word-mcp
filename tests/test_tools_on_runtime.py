import importlib
import inspect
import json

import pytest
from fastmcp import Client

MODULES = ["edit", "tables", "references", "read", "layout", "screen"]
READ_ONLY = {
    "word_live_list_cross_reference_items",
    "word_live_take_snapshot",
    "word_live_get_diff",
    "word_live_snapshot_status",
    "word_live_get_text",
    "word_live_get_paragraph_format",
    "word_live_get_info",
    "word_live_list_open",
    "word_live_find_text",
    "word_live_get_comments",
    "word_live_list_revisions",
    "word_live_get_page_text",
    "word_live_get_undo_history",
    "word_live_diagnose_layout",
    "word_screen_capture",
}


def public_tools():
    found = {}
    for name in MODULES:
        module = importlib.import_module(f"word_mcp.tools.{name}")
        for attr, value in vars(module).items():
            if attr.startswith("word_") and callable(value):
                found[attr] = value
    return found


def test_all_45_tools_are_live_tools():
    tools = public_tools()
    assert len(tools) == 45
    for name, tool in tools.items():
        assert inspect.iscoroutinefunction(tool), name
        assert callable(tool.sync), name
        assert tool.mutates is (name not in READ_ONLY), name


def test_no_blocking_attach_code_remains():
    import pathlib

    for path in pathlib.Path("word_mcp").rglob("*.py"):
        text = path.read_text(encoding="utf-8")
        assert "word_mcp.word_com" not in text, path
        assert 'Dispatch("Word.Application")' not in text, path
        if path.parent.name == "tools":
            assert "async def" not in text, path
            assert "await " not in text, path


async def test_list_open_through_fastmcp(connector):
    from word_mcp import server

    if not getattr(server, "_registered_for_tests", False):
        server.register_tools()
        server._registered_for_tests = True
    async with Client(server.mcp) as client:
        result = await client.call_tool("word_live_list_open", {})
    payload = json.loads(result.content[0].text)
    assert [d["name"] for d in payload["documents"]] == ["report.docx"]


async def test_unknown_document_gives_structured_error(connector):
    from word_mcp.tools import read

    payload = json.loads(await read.word_live_get_info("missing.docx"))
    assert payload["code"] == "document_not_found"
    assert set(payload) == {"error", "code", "retryable", "hint"}
