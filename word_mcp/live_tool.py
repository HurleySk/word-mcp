import functools
import json
import os

from word_mcp.com_runtime import error_json, run_com
from word_mcp.errors import WordError

READ_ATTEMPTS = 3


def _default_timeout():
    return float(os.environ.get("WORD_MCP_TIMEOUT", "60"))


def _call(session, body, args, kwargs):
    session.last_error = None
    try:
        return body(*args, **kwargs)
    except Exception as exc:
        return error_json(exc)


def _invoke(session, body, args, kwargs, mutates):
    attempts = 1 if mutates else READ_ATTEMPTS
    for attempt in range(attempts):
        if attempt:
            session.pause(0.5 * attempt)
            session.wait_until_ready()
        result = _call(session, body, args, kwargs)
        err = session.last_error
        if err is None or err.code != "busy":
            return result
    if mutates:
        partial = WordError(
            "busy",
            "Word became busy part-way through the edit",
            retryable=False,
            hint="The edit may be partial. Check the document, or call word_live_undo, before retrying.",
        )
        return json.dumps(partial.to_dict(), ensure_ascii=False)
    return result


def live_tool(*, mutates, timeout=None):
    def decorate(body):
        @functools.wraps(body)
        async def tool(*args, **kwargs):
            limit = _default_timeout() if timeout is None else timeout
            try:
                return await run_com(
                    lambda session: _invoke(session, body, args, kwargs, mutates),
                    timeout=limit,
                )
            except Exception as exc:
                return error_json(exc)

        tool.sync = body
        tool.mutates = mutates
        return tool

    return decorate
