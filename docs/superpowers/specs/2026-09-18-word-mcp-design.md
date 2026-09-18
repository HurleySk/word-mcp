# word-mcp design

Date: 2026-09-18
Base: `HurleySk/word-mcp-live` at `d3ba8fc`, fork of `ykarapazar/word-mcp-live` 1.6.2

## Problem

The server often fails to launch and drops mid-session. Measured on 2026-09-18:

- The committed `uv.lock` has no `fastmcp` entry and `pyproject.toml` declares `fastmcp>=2.8.1`, so each fresh install resolves the newest release. The probe resolved 4.0.5, published one day earlier.
- A cold start downloads 80 packages and exceeds the 30 s MCP connect timeout. A warm start takes 11.5 s.
- `tools/list` returns 120 tools in 90 KB. 44 are `word_live_*`.
- The process exits 1 when stdin closes. `pytest` is a runtime dependency.
- Every live tool is `async def` and makes blocking COM calls on the asyncio loop thread. Nothing calls `CoInitialize`, no COM message filter is registered, and no call has a timeout.
- `get_word_app()` rescans the Running Object Table on every call and falls back to `Dispatch`, which can start a second Word instance.
- Attach, undo-record and track-changes restore logic is copy-pasted across 43 call sites. Errors are returned as `{"error": str(e)}`.
- No tests cover the live tools.

## Goals

1. Launch succeeds deterministically, warm start under 3 s.
2. A busy, hung or closed Word produces a structured tool error. The server process survives.
3. Only the live COM tools remain, on Windows only.
4. Tool names and signatures of the 44 live tools and `screen_capture` are unchanged.

## Non-goals

macOS support, closed-file `.docx` editing, HTTP or SSE transport, hosted deployment, a C# port, redesigning the tool surface.

## Identity

| Item | Value |
|---|---|
| GitHub repo | `HurleySk/word-mcp`, remains a fork of `ykarapazar/word-mcp-live` |
| Local folder | `word-mcp` |
| Python package | `word_mcp` |
| Distribution and console script | `word-mcp` |
| MCP server name | `word-mcp` |
| Version | `0.1.0` |
| Licence | MIT, upstream copyright lines retained |

The README states the upstream origin and what differs.

## Removed

- Package `office_word_mcp_server`
- `word_document_server/tools/`: `comment_tools`, `comment_write_tools`, `content_tools`, `document_tools`, `extended_document_tools`, `footnote_tools`, `format_tools`, `hyperlink_tools`, `layout_tools`, `protection_tools`, `tracked_changes_tools`
- `word_document_server/core/`: `comment_writer`, `comments`, `footnotes`, `hyperlink_writer`, `protection`, `styles`, `tables`, `tracked_changes`, `unprotect`, `word_mac`
- `word_document_server/utils/`: `document_utils`, `extended_document_utils`, `file_utils`, `path_utils`, `save_utils`
- Every `_MAC_AVAILABLE` branch in the live modules
- `Dockerfile`, `RENDER_DEPLOYMENT.md`, `smithery.yaml`, `manifest.json`, `.mcpbignore`, `server.json`, `setup_mcp.py`, `requirements.txt`, `word_mcp_server.py`, root `__init__.py`, `test_formatting.py`, `tests/test_convert_to_pdf.py`
- `.env` loading and the transport switch in `main.py`. Transport is stdio only.

Before deleting a `core` or `utils` module, grep the live modules for imports of it. A module a live tool imports is kept and moved.

## Layout

```
word_mcp/
  __init__.py
  server.py            FastMCP instance, tool registration, run()
  com_runtime.py       STA worker, message filter, run_com, attach cache
  live_tool.py         @live_tool decorator, error shaping
  errors.py            WordError codes and HRESULT mapping
  defaults.py
  text_safety.py
  table_com.py
  tools/
    edit.py            from live_tools.py
    read.py            from live_read_tools.py
    layout.py          from live_layout_tools.py
    screen.py          from screen_capture_tools.py
tests/
  fakes.py
  test_com_runtime.py
  test_live_tool.py
  test_smoke_handshake.py
  integration/         @pytest.mark.word
```

`live_tools.py` is 2,200 lines. When it moves to `tools/edit.py`, split it into `edit.py`, `tables.py` and `references.py` if any file would exceed about 800 lines after the boilerplate is removed.

## Dependencies

Runtime: `fastmcp>=4.0.5,<5` (the current registration code was verified against 4.0.5), `pywin32>=306`, `Pillow>=10`.

Dev group: `pytest`, `pytest-asyncio`.

`uv.lock` is regenerated and committed. CI runs `uv lock --check`.

## com_runtime

One daemon thread named `word-com`, started on first use.

Thread start: `pythoncom.CoInitialize()`. While idle the thread calls `pythoncom.PumpWaitingMessages()` every 100 ms.

pywin32 does not expose `CoRegisterMessageFilter`, so busy handling is explicit. Before each tool body runs, the worker polls `app.Documents.Count` with backoff for up to 10 s until Word accepts a call. A body that is then rejected mid-way is retried from the top only when the tool is read-only. A mutating tool is never re-run: it returns `busy` with a hint that the edit may be partial and can be undone.

Public API:

```python
async def run_com(fn: Callable[[WordSession], T], *, timeout: float = 60.0) -> T
```

`fn` runs on the COM thread and receives a `WordSession` exposing `app` and `document(filename)`. COM objects never leave the COM thread: `fn` returns plain Python data.

Calls are serialised through a queue. On timeout the awaiting coroutine gets `WordError(code="timeout", retryable=True)`. The worker is marked poisoned, a replacement thread is started for later calls, and the stuck thread is abandoned as a daemon. The cached app handle is dropped.

Attach:

1. If a cached app exists, validate with one cheap property read. A disconnect HRESULT (`RPC_E_DISCONNECTED`, `RPC_S_SERVER_UNAVAILABLE`, `CO_E_OBJNOTCONNECTED`) drops the cache.
2. `GetActiveObject("Word.Application")`. If it has documents, use it.
3. ROT scan, logic carried over from `_find_word_with_docs`.
4. Otherwise raise `WordError(code="word_not_running")`.

`Dispatch("Word.Application")` is removed. None of the 45 tools opens a file, so nothing in the server starts Word.

The default per-call timeout is 60 s and can be changed with the `WORD_MCP_TIMEOUT` environment variable. When a worker is poisoned, calls still queued on it fail at once with `timeout`.

`find_document` and `undo_record` move here unchanged in behaviour.

## errors

`WordError(code, message, retryable, hint)`. Codes: `word_not_running`, `no_document`, `document_not_found`, `busy`, `modal_dialog`, `timeout`, `disconnected`, `invalid_argument`, `com_error`, `internal`.

`pywintypes.com_error` is mapped by HRESULT: `RPC_E_CALL_REJECTED` and `RPC_E_SERVERCALL_RETRYLATER` to `busy`, the disconnect set to `disconnected`, anything else to `com_error` with the HRESULT and Word's description in `message`.

`busy` that persists past the 10 s readiness poll is reported as `modal_dialog` with the hint "Word may have a dialog open. Close it and retry."

## live_tool decorator

Tool schemas, titles and annotations live in the registration wrappers carried over from `main.py`, which call the implementations positionally. The implementations therefore keep their current signatures and bodies.

```python
@live_tool(mutates=True)
def word_live_insert_text(filename=None, text="", position="end", bookmark=None, track_changes=False) -> str
```

`live_tool` turns the synchronous body into an `async def` that runs it through `run_com`. The undecorated body is available as `.sync` for tools that call other tools. Inside a body, `get_word_app()` returns the worker's cached app, so the 45 existing call sites are unchanged.

Each body's trailing `except Exception as e: return json.dumps({"error": str(e)})` becomes `return error_json(e)`, which classifies the exception, drops the cached app on a disconnect, and returns `{"error", "code", "retryable", "hint"}`. The decorator applies the same shaping to timeouts and to anything that escapes a body. Nothing raises into FastMCP.

The migration is scripted because the patterns are exact: 45 `async def`, 45 identical blanket handlers, and every `if _MAC_AVAILABLE:` block ends at the following `if sys.platform != "win32":` line.

Folding the per-tool undo and track-changes restore code into the decorator is deferred. It already sits in `try/finally`, there are no live-tool tests to catch a regression, and it is not a cause of the failures.

## Startup

- `pywin32` and `PIL` are imported inside the COM thread and the screen tool, not at module import.
- FastMCP banner disabled with `show_banner=False`. All logging to stderr at WARNING.
- Measured: importing the whole current server takes 1.1 s. The 11.5 s warm start is `uv run` re-syncing the project.
- EOF on stdin exits 0.
- A non-Windows platform exits 1 at startup with one line on stderr.
- README install path: `uv tool install git+https://github.com/HurleySk/word-mcp`, then register the `word-mcp` command. Launch performs no dependency resolution.

## Testing

`tests/fakes.py` provides a fake `WordSession` and fake COM objects, injected through a `com_runtime` factory hook so unit tests run without Word or pywin32.

- `test_com_runtime.py`: calls are serialised on one thread; timeout returns `timeout` and the next call succeeds on a fresh worker; a disconnect HRESULT triggers one re-attach; no Word gives `word_not_running` and launches nothing; the event loop stays responsive during a slow call.
- `test_live_tool.py`: the body runs on the COM thread; `.sync` is the raw body; an escaped exception and a timeout both produce the error JSON shape; a read-only tool is retried after a mid-body `busy` and a mutating tool is not.
- `test_smoke_handshake.py`: spawn the console script, send `initialize` and `tools/list`, assert 45 tools, a response within 5 s, and exit code 0 after stdin closes.
- `tests/integration/`: opt-in, against real Word with a scratch document. Covers list, insert, read back, replace and undo. The modal-dialog case cannot be automated, because showing a dialog blocks the test's own COM call, so it is a manual check in the final task.

CI on `windows-latest`: `uv lock --check`, unit tests, smoke test.

## Work order

1. Rename package, prune files, stdio-only `server.py`. The live tools still run on their old code path.
2. Pin dependencies, regenerate `uv.lock`, move pytest to dev.
3. `errors.py`, `com_runtime.py`, `fakes.py`, runtime tests.
4. `live_tool.py` and its tests.
5. Migrate tools one module at a time: read, layout, screen, edit.
6. Startup cleanup and the smoke test.
7. CI workflow, README, CHANGELOG.
8. Rename the GitHub repo and local folder, update remotes, register in Claude Code config. The GitHub rename needs explicit confirmation when reached.

## Risks

- The scripted migration touches 5,000 lines with no behavioural tests. The smoke test compares `tools/list` against a snapshot taken before migration, and every tool module is compiled and imported after each scripted step.
- An abandoned stuck COM thread holds a reference to Word. This is accepted: the alternative is a dead server.