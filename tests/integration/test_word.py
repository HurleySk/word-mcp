import json

import pytest

from word_mcp import com_runtime
from word_mcp.tools import edit, read

pytestmark = pytest.mark.word


@pytest.fixture
def scratch():
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = win32com.client.Dispatch("Word.Application")
    app.Visible = True
    doc = app.Documents.Add()
    com_runtime.reset()
    yield doc
    doc.Close(SaveChanges=0)
    com_runtime.reset()


async def test_list_open_sees_the_scratch_document(scratch):
    payload = json.loads(await read.word_live_list_open())
    assert scratch.Name in [d["name"] for d in payload["documents"]]


async def test_insert_replace_undo(scratch):
    name = scratch.Name
    inserted = json.loads(await edit.word_live_insert_text(name, "alpha beta", "end"))
    assert inserted["success"] is True
    assert "alpha beta" in scratch.Content.Text

    replaced = json.loads(await edit.word_live_replace_text(name, "beta", "gamma"))
    assert replaced.get("success") is True
    assert "alpha gamma" in scratch.Content.Text

    undone = json.loads(await edit.word_live_undo(name, 1))
    assert undone.get("success") is True
    assert "alpha beta" in scratch.Content.Text


async def test_missing_document_is_a_structured_error(scratch):
    payload = json.loads(await read.word_live_get_info("no-such-file.docx"))
    assert payload["code"] == "document_not_found"
