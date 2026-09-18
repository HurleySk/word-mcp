RPC_E_CALL_REJECTED = -2147418111
RPC_E_SERVERCALL_RETRYLATER = -2147417846
RPC_E_DISCONNECTED = -2147417848
RPC_S_SERVER_UNAVAILABLE = -2147023174
RPC_S_CALL_FAILED = -2147023170
CO_E_OBJNOTCONNECTED = -2147220995

BUSY_HRESULTS = frozenset({RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER})
DISCONNECT_HRESULTS = frozenset(
    {RPC_E_DISCONNECTED, RPC_S_SERVER_UNAVAILABLE, RPC_S_CALL_FAILED, CO_E_OBJNOTCONNECTED}
)


class WordError(Exception):
    def __init__(self, code, message, *, retryable=False, hint=None):
        super().__init__(message)
        self.code = code
        self.message = message
        self.retryable = retryable
        self.hint = hint

    def to_dict(self):
        return {
            "error": self.message,
            "code": self.code,
            "retryable": self.retryable,
            "hint": self.hint,
        }


def _com_message(exc, hresult):
    info = getattr(exc, "excepinfo", None)
    detail = None
    if info and len(info) > 2:
        detail = info[2]
    detail = detail or getattr(exc, "strerror", None) or "COM error"
    return f"{detail} (HRESULT 0x{hresult & 0xFFFFFFFF:08X})"


def classify(exc):
    if isinstance(exc, WordError):
        return exc
    hresult = getattr(exc, "hresult", None)
    if isinstance(hresult, int):
        if hresult in BUSY_HRESULTS:
            return WordError(
                "busy",
                "Word rejected the call because it is busy",
                retryable=True,
                hint="Wait for Word to finish, close any open dialog, then retry.",
            )
        if hresult in DISCONNECT_HRESULTS:
            return WordError(
                "disconnected",
                "Lost the connection to Word",
                retryable=True,
                hint="Word was closed or restarted. Retry to reconnect.",
            )
        return WordError("com_error", _com_message(exc, hresult))
    if isinstance(exc, ValueError):
        return WordError("invalid_argument", str(exc))
    return WordError("internal", f"{type(exc).__name__}: {exc}")
