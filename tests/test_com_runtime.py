import asyncio
import json
import threading
import time

import pytest

from tests.fakes import FakeApp, FakeClock, FakeComError, FakeConnector
from word_mcp import com_runtime
from word_mcp.com_runtime import current_session, error_json, find_document, get_word_app, run_com
from word_mcp.errors import RPC_E_DISCONNECTED, WordError


async def test_calls_run_on_the_word_com_thread(connector):
    names = [await run_com(lambda s: threading.current_thread().name) for _ in range(3)]
    assert names == ["word-com"] * 3
    assert connector.thread_names == ["word-com"]
    assert connector.attach_count == 1


async def test_event_loop_stays_responsive(connector):
    ticks = 0

    async def ticker():
        nonlocal ticks
        while True:
            await asyncio.sleep(0.01)
            ticks += 1

    task = asyncio.create_task(ticker())
    await run_com(lambda s: time.sleep(0.3))
    task.cancel()
    assert ticks >= 10


async def test_timeout_then_fresh_worker(connector):
    release = threading.Event()
    with pytest.raises(WordError) as caught:
        await run_com(lambda s: release.wait(5), timeout=0.2)
    assert caught.value.code == "timeout"
    assert caught.value.retryable is True
    assert await run_com(lambda s: "ok") == "ok"
    assert len(connector.thread_names) == 2
    release.set()


async def test_queued_call_fails_fast_when_worker_is_poisoned(connector):
    release = threading.Event()
    stuck = asyncio.create_task(run_com(lambda s: release.wait(5), timeout=0.2))
    await asyncio.sleep(0.05)
    started = time.perf_counter()
    with pytest.raises(WordError) as caught:
        await run_com(lambda s: "never", timeout=5)
    assert caught.value.code == "timeout"
    assert time.perf_counter() - started < 1.0
    with pytest.raises(WordError):
        await stuck
    release.set()


async def test_reattaches_after_disconnect(connector):
    await run_com(lambda s: s.app.Documents.Count)
    connector.apps[0].dead = True
    connector.apps.append(FakeApp())
    assert await run_com(lambda s: s.app.Documents.Count) == 1
    assert connector.attach_count == 2


async def test_word_not_running():
    fake = FakeConnector(apps=[])
    com_runtime.configure(lambda: fake, sleep=lambda seconds: None)
    try:
        with pytest.raises(WordError) as caught:
            await run_com(lambda s: s.app)
        assert caught.value.code == "word_not_running"
    finally:
        com_runtime.reset()


async def test_waits_out_a_short_busy_spell(connector):
    connector.apps[0].busy_calls = 3
    assert await run_com(lambda s: s.app.Documents.Count) == 1


async def test_persistent_busy_is_reported_as_modal_dialog():
    fake = FakeConnector()
    fake.apps[0].busy_calls = 10**6
    clock = FakeClock()
    com_runtime.configure(lambda: fake, clock=clock, sleep=clock.sleep)
    try:
        with pytest.raises(WordError) as caught:
            await run_com(lambda s: s.app.Documents.Count)
        assert caught.value.code == "modal_dialog"
        assert clock.now >= 10.0
    finally:
        com_runtime.reset()


async def test_get_word_app_only_works_on_the_com_thread(connector):
    with pytest.raises(WordError):
        get_word_app()
    app = await run_com(lambda s: get_word_app())
    assert app is connector.apps[0]


async def test_find_document(connector):
    name = await run_com(lambda s: find_document(s.app, "REPORT.docx").Name)
    assert name == "report.docx"
    with pytest.raises(WordError) as caught:
        await run_com(lambda s: find_document(s.app, "missing.docx"))
    assert caught.value.code == "document_not_found"
    connector.apps[0].docs.clear()
    with pytest.raises(WordError) as caught:
        await run_com(lambda s: find_document(s.app, None))
    assert caught.value.code == "no_document"


async def test_error_json_drops_the_app_on_disconnect(connector):
    def body(session):
        assert session.app is connector.apps[0]
        text = error_json(FakeComError(RPC_E_DISCONNECTED))
        return text, session._app, current_session() is session

    text, cached, same = await run_com(body)
    assert json.loads(text)["code"] == "disconnected"
    assert cached is None
    assert same is True
