# Contributing to word-mcp

## Development setup

Requires Windows, Microsoft Word, and [uv](https://docs.astral.sh/uv/).

```
git clone https://github.com/HurleySk/word-mcp.git
cd word-mcp
uv sync
uv run pytest
```

`uv run pytest -m word` runs the integration tests against a real Word.

## Project structure

```
word_mcp/
  server.py         FastMCP instance and the 45 registration wrappers
  com_runtime.py    STA worker thread, Word attach, run_com, find_document, undo_record
  live_tool.py      @live_tool decorator and the busy retry policy
  errors.py         WordError and HRESULT classification
  defaults.py       MCP_AUTHOR and MCP_AUTHOR_INITIALS
  table_com.py      table helpers used by word_live_modify_table
  text_safety.py    control-character validation for find and insert text
  tools/            edit, tables, references, read, layout, screen
```

## Adding a tool

Write a synchronous body in the matching `word_mcp/tools/` module. It runs on the COM thread and returns a JSON string.

```python
@live_tool(mutates=True)
def word_live_my_tool(filename: str = None) -> str:
    """Description shown to the MCP client."""
    try:
        from word_mcp.com_runtime import find_document, get_word_app, undo_record

        app = get_word_app()
        doc = find_document(app, filename)
        with undo_record(app, "MCP: My Tool"):
            pass
        return json.dumps({"success": True})
    except Exception as e:
        return error_json(e)
```

Set `mutates=False` only for a tool that changes nothing, because read-only tools are re-run when Word turns busy. Wrap every edit in `undo_record`. Call another tool from inside a body with `other_tool.sync(...)`.

Register it in `word_mcp/server.py` with an `async def` wrapper that awaits the tool, then refresh the schema snapshot:

```
uv run python scripts/dump_tools.py tests/snapshots/live_tools.json .venv\Scripts\python.exe -c "from word_mcp.server import run; run()"
```

## Pull requests

One feature or fix per PR. Add a changelog entry. `uv lock --check` and `uv run pytest` must pass.
