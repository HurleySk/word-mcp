import json
import pathlib
import sys

from tests.handshake import handshake, live_tools

SNAPSHOT = pathlib.Path(__file__).parent / "snapshots" / "live_tools.json"
COMMAND = [sys.executable, "-c", "from word_document_server.main import run_server; run_server()"]


def test_live_tool_schemas_match_snapshot():
    expected = json.loads(SNAPSHOT.read_text(encoding="utf-8"))
    actual = live_tools(handshake(COMMAND)["tools"])
    assert [t["name"] for t in actual] == [t["name"] for t in expected]
    assert actual == expected
