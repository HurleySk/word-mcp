# Privacy Policy — word-mcp

word-mcp runs entirely on your machine and does not collect, transmit, or store user data.

- No telemetry, analytics, or crash reports.
- No network requests. The only transport is stdio to your MCP client.
- No databases, logs, or caches beyond the current session. Paragraph snapshots taken by `word_live_take_snapshot` live in memory and vanish when the server exits.
- The server talks to Microsoft Word through COM automation, which is local inter-process communication.

## Environment variables

Read locally and never transmitted: `WORD_MCP_TIMEOUT`, `MCP_AUTHOR`, `MCP_AUTHOR_INITIALS`.

## Source

MIT licensed. Source: https://github.com/HurleySk/word-mcp. Questions: https://github.com/HurleySk/word-mcp/issues
