import inspect
import json
import threading

from tests.fakes import FakeComError
from word_mcp.com_runtime import error_json, get_word_app
from word_mcp.errors import RPC_E_CALL_REJECTED
from word_mcp.live_tool import live_tool


async def test_body_runs_on_the_com_thread_and_keeps_its_identity(connector):
    @live_tool(mutates=False)
    def word_live_example(filename: str = None, count: int = 1) -> str:
        """Example docstring."""
        return json.dumps({"thread": threading.current_thread().name, "count": count, "docs": get_word_app().Documents.Count})

    assert inspect.iscoroutinefunction(word_live_example)
    assert word_live_example.__name__ == "word_live_example"
    assert word_live_example.__doc__ == "Example docstring."
    assert list(inspect.signature(word_live_example).parameters) == ["filename", "count"]
    assert word_live_example.mutates is False
    result = json.loads(await word_live_example(None, 7))
    assert result == {"thread": "word-com", "count": 7, "docs": 1}


async def test_sync_is_the_raw_body(connector):
    @live_tool(mutates=False)
    def inner(value):
        return json.dumps({"value": value})

    @live_tool(mutates=False)
    def outer(value):
        return inner.sync(value)

    assert json.loads(await outer(3)) == {"value": 3}


async def test_escaped_exception_becomes_error_json(connector):
    @live_tool(mutates=True)
    def broken():
        raise KeyError("oops")

    result = json.loads(await broken())
    assert set(result) == {"error", "code", "retryable", "hint"}
    assert result["code"] == "internal"


async def test_timeout_becomes_error_json(connector):
    release = threading.Event()

    @live_tool(mutates=False, timeout=0.2)
    def slow():
        release.wait(5)
        return "{}"

    result = json.loads(await slow())
    release.set()
    assert result["code"] == "timeout"
    assert result["retryable"] is True


async def test_read_only_tool_is_retried_after_mid_body_busy(connector):
    calls = []

    @live_tool(mutates=False)
    def flaky():
        calls.append(1)
        if len(calls) < 3:
            try:
                raise FakeComError(RPC_E_CALL_REJECTED)
            except Exception as e:
                return error_json(e)
        return json.dumps({"success": True})

    assert json.loads(await flaky()) == {"success": True}
    assert len(calls) == 3


async def test_mutating_tool_is_never_retried(connector):
    calls = []

    @live_tool(mutates=True)
    def flaky():
        calls.append(1)
        try:
            raise FakeComError(RPC_E_CALL_REJECTED)
        except Exception as e:
            return error_json(e)

    result = json.loads(await flaky())
    assert len(calls) == 1
    assert result["code"] == "busy"
    assert result["retryable"] is False
    assert "undo" in result["hint"].lower()
