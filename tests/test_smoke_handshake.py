import json
import pathlib
import sys

from tests.handshake import handshake, live_tools

SNAPSHOT = pathlib.Path(__file__).parent / "snapshots" / "live_tools.json"
COMMAND = [sys.executable, "-c", "from word_mcp.server import run; run()"]


def test_live_tool_schemas_match_snapshot():
    expected = json.loads(SNAPSHOT.read_text(encoding="utf-8"))
    actual = live_tools(handshake(COMMAND)["tools"])
    assert [t["name"] for t in actual] == [t["name"] for t in expected]
    assert actual == expected


CONSOLE_SCRIPT = [str(pathlib.Path(sys.executable).with_name("word-mcp.exe"))]


def test_console_script_starts_fast_and_exits_clean():
    handshake(CONSOLE_SCRIPT)
    result = handshake(CONSOLE_SCRIPT)
    assert len(result["tools"]) == 45
    assert result["elapsed"] < 10.0
    assert result["exit_code"] == 0
    assert "FastMCP" not in result["stderr"]
