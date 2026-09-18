from word_mcp.errors import (
    CO_E_OBJNOTCONNECTED,
    RPC_E_CALL_REJECTED,
    RPC_E_DISCONNECTED,
    RPC_E_SERVERCALL_RETRYLATER,
    RPC_S_SERVER_UNAVAILABLE,
    WordError,
    classify,
)


class ComError(Exception):
    def __init__(self, hresult, strerror="failed", excepinfo=None):
        super().__init__(strerror)
        self.hresult = hresult
        self.strerror = strerror
        self.excepinfo = excepinfo


def test_to_dict_has_the_four_keys():
    err = WordError("busy", "Word is busy", retryable=True, hint="wait")
    assert err.to_dict() == {"error": "Word is busy", "code": "busy", "retryable": True, "hint": "wait"}


def test_word_error_passes_through():
    err = WordError("timeout", "slow")
    assert classify(err) is err


def test_busy_hresults():
    for hresult in (RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER):
        err = classify(ComError(hresult))
        assert err.code == "busy"
        assert err.retryable is True


def test_disconnect_hresults():
    for hresult in (RPC_E_DISCONNECTED, RPC_S_SERVER_UNAVAILABLE, CO_E_OBJNOTCONNECTED):
        err = classify(ComError(hresult))
        assert err.code == "disconnected"
        assert err.retryable is True


def test_other_com_error_uses_word_description():
    exc = ComError(-2147352567, "Exception occurred.", (0, "Microsoft Word", "The requested member does not exist.", None, 0, -2146822347))
    err = classify(exc)
    assert err.code == "com_error"
    assert err.retryable is False
    assert "The requested member does not exist." in err.message
    assert "0x80020009" in err.message


def test_value_error_is_invalid_argument():
    assert classify(ValueError("bad position")).code == "invalid_argument"


def test_anything_else_is_internal():
    err = classify(KeyError("x"))
    assert err.code == "internal"
    assert "KeyError" in err.message
