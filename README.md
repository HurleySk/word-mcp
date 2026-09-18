# word-mcp

MCP server for editing documents that are open in Microsoft Word on Windows. It drives the running Word over COM, so changes appear live and land on Word's undo stack.

This is a hardened fork of [ykarapazar/word-mcp-live](https://github.com/ykarapazar/word-mcp-live), itself based on GongRzhe's Office-Word-MCP-Server. The 44 `word_live_*` tools and `word_screen_capture` keep their upstream names and parameters. See [TOOLS.md](TOOLS.md).

## What differs from upstream

- Windows and live COM tools only. The python-docx file tools, macOS support, HTTP transports and hosted-deployment files are gone.
- Every COM call runs on one dedicated STA thread. The server stays responsive while Word works.
- Each call has a timeout (`WORD_MCP_TIMEOUT`, default 60 seconds). A hung Word produces a `timeout` error and a fresh worker.
- The server waits up to 10 seconds for a busy Word, then reports `modal_dialog`.
- The server never starts Word. With no Word running, tools return `word_not_running`.
- Read-only tools are retried when Word turns busy mid-call. Editing tools are never re-run.
- Errors are JSON: `{"error", "code", "retryable", "hint"}`.
- Dependencies are pinned and locked.

## Install

Requires Windows, Microsoft Word, and [uv](https://docs.astral.sh/uv/).

```
uv tool install --compile-bytecode git+https://github.com/HurleySk/word-mcp
claude mcp add --scope user word-mcp -- word-mcp
```

`uv tool install` builds an isolated environment once, so launching the server resolves nothing. `--compile-bytecode` keeps the first launch under 3 seconds instead of about 25.

To update: `uv tool upgrade word-mcp`.

## Configuration

| Variable | Default | Meaning |
|---|---|---|
| `WORD_MCP_TIMEOUT` | `60` | Seconds before a tool call gives up on Word |
| `MCP_AUTHOR` | `Author` | Author name for tracked changes and comments |
| `MCP_AUTHOR_INITIALS` | empty | Initials for comments |

## Development

```
uv sync
uv run pytest
uv run pytest -m word
```

The second pytest command runs the integration tests against a real Word and opens a scratch document.

## Licence

MIT. See [LICENSE](LICENSE).
