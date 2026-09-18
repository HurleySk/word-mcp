import asyncio
import concurrent.futures
import json
import os
import queue
import threading
import time
import unicodedata
from contextlib import contextmanager

from word_mcp.errors import WordError, classify

_local = threading.local()
_lock = threading.Lock()
_worker = None
_connector_factory = None
_clock = time.monotonic
_sleep = time.sleep
OBJID_NATIVEOM = 0xFFFFFFF0


class Win32Connector:
    def thread_init(self):
        import pythoncom

        pythoncom.CoInitialize()

    def thread_exit(self):
        import pythoncom

        pythoncom.CoUninitialize()

    def pump(self):
        import pythoncom

        pythoncom.PumpWaitingMessages()

    def attach(self):
        import win32com.client

        try:
            app = win32com.client.GetActiveObject("Word.Application")
        except Exception:
            app = None
        if app is not None:
            try:
                has_docs = app.Documents.Count > 0
            except Exception as exc:
                if classify(exc).code == "busy":
                    return app
                app = None
                has_docs = False
            if has_docs:
                return app
        found = _find_word_with_docs() or _find_word_by_window()
        if found is not None:
            return found
        if app is not None:
            return app
        raise WordError(
            "word_not_running",
            "Microsoft Word is not running",
            hint="Open the document in Word, then retry.",
        )


class WordSession:
    def __init__(self, connector):
        self._connector = connector
        self._app = None
        self.last_error = None

    @property
    def app(self):
        if self._app is None:
            self._app = self._connector.attach()
        return self._app

    def drop(self):
        self._app = None

    def pause(self, seconds):
        _sleep(seconds)

    def wait_until_ready(self, window=10.0):
        deadline = _clock() + window
        delay = 0.1
        while True:
            try:
                self.app.Documents.Count
                return
            except Exception as exc:
                err = classify(exc)
            if err.code == "disconnected":
                self.drop()
            elif err.code != "busy":
                raise err
            if _clock() >= deadline:
                if err.code == "busy":
                    raise WordError(
                        "modal_dialog",
                        f"Word did not accept calls for {window:g} seconds",
                        retryable=True,
                        hint="Word may have a dialog open. Close it and retry.",
                    )
                raise err
            _sleep(delay)
            delay = min(delay * 2, 1.0)


class ComWorker:
    def __init__(self, connector):
        self._connector = connector
        self._queue = queue.SimpleQueue()
        self.poisoned = False
        self._thread = threading.Thread(target=self._run, name="word-com", daemon=True)
        self._thread.start()

    def submit(self, fn):
        future = concurrent.futures.Future()
        self._queue.put((fn, future))
        return future

    def poison(self):
        self.poisoned = True
        while True:
            try:
                _, future = self._queue.get_nowait()
            except queue.Empty:
                return
            if future.set_running_or_notify_cancel():
                future.set_exception(_timeout_error("Word stopped responding"))

    def _run(self):
        self._connector.thread_init()
        session = WordSession(self._connector)
        _local.session = session
        try:
            while not self.poisoned:
                try:
                    fn, future = self._queue.get(timeout=0.1)
                except queue.Empty:
                    self._connector.pump()
                    continue
                if not future.set_running_or_notify_cancel():
                    continue
                try:
                    session.wait_until_ready()
                    future.set_result(fn(session))
                except BaseException as exc:
                    future.set_exception(exc)
        finally:
            _local.session = None
            self._connector.thread_exit()


def _timeout_error(message):
    return WordError(
        "timeout",
        message,
        retryable=True,
        hint="Word may be showing a dialog or processing a large document.",
    )


def configure(connector_factory, *, clock=time.monotonic, sleep=time.sleep):
    global _connector_factory, _clock, _sleep
    reset()
    _connector_factory = connector_factory
    _clock = clock
    _sleep = sleep


def reset():
    global _worker, _connector_factory, _clock, _sleep
    with _lock:
        if _worker is not None:
            _worker.poison()
        _worker = None
    _connector_factory = None
    _clock = time.monotonic
    _sleep = time.sleep


def _get_worker():
    global _worker
    with _lock:
        if _worker is None or _worker.poisoned:
            factory = _connector_factory or Win32Connector
            _worker = ComWorker(factory())
        return _worker


async def run_com(fn, *, timeout=60.0):
    worker = _get_worker()
    future = worker.submit(fn)
    try:
        return await asyncio.wait_for(asyncio.wrap_future(future), timeout)
    except asyncio.TimeoutError:
        still_queued = future.cancel()
        if not still_queued:
            worker.poison()
        raise _timeout_error(f"Word did not respond within {timeout:g} seconds") from None


def current_session():
    return getattr(_local, "session", None)


def get_word_app():
    session = current_session()
    if session is None:
        raise WordError("internal", "get_word_app was called off the COM thread")
    return session.app


def error_json(exc):
    err = classify(exc)
    session = current_session()
    if session is not None:
        session.last_error = err
        if err.code == "disconnected":
            session.drop()
    return json.dumps(err.to_dict(), ensure_ascii=False)


def _find_word_by_window():
    try:
        import ctypes
        from ctypes import wintypes

        import pythoncom
        import win32com.client
        import win32gui

        accessible = ctypes.windll.oleacc.AccessibleObjectFromWindow
        accessible.argtypes = [
            wintypes.HWND,
            wintypes.DWORD,
            ctypes.c_void_p,
            ctypes.POINTER(ctypes.c_void_p),
        ]
        accessible.restype = ctypes.c_long
        iid = (ctypes.c_byte * 16).from_buffer_copy(bytes(pythoncom.IID_IDispatch))
        panes = []

        def on_child(hwnd, _):
            if win32gui.GetClassName(hwnd) == "_WwG":
                panes.append(hwnd)
            return True

        def on_top(hwnd, _):
            if win32gui.GetClassName(hwnd) == "OpusApp":
                try:
                    win32gui.EnumChildWindows(hwnd, on_child, None)
                except Exception:
                    pass
            return True

        win32gui.EnumWindows(on_top, None)
        for hwnd in panes:
            pointer = ctypes.c_void_p()
            if accessible(hwnd, OBJID_NATIVEOM, iid, ctypes.byref(pointer)) != 0 or not pointer.value:
                continue
            try:
                window = win32com.client.Dispatch(
                    pythoncom.ObjectFromAddress(pointer.value, pythoncom.IID_IDispatch)
                )
                return window.Application
            except Exception:
                continue
    except Exception:
        pass
    return None


def _find_word_with_docs():
    """Scan the Running Object Table for a Word.Application with open docs.

    Handles Office 365 / OneDrive scenarios where GetActiveObject returns an
    empty Application proxy.  In these cases, documents are registered in the
    ROT as file monikers (.docx paths or https://d.docs.live.net/... URLs).
    We grab the Document COM object from such a moniker and reach the real
    Application via ``doc.Application``.

    Returns the Word.Application COM object if found, or None.
    """
    try:
        import pythoncom
        import win32com.client

        rot = pythoncom.GetRunningObjectTable(0)
        enum = rot.EnumRunning()

        monikers_to_retry = []
        while True:
            batch = enum.Next(1)
            if not batch:
                break
            moniker = batch[0]
            try:
                ctx = pythoncom.CreateBindCtx(0)
                name = moniker.GetDisplayName(ctx, None)
                obj = rot.GetObject(moniker)
                dispatch = obj.QueryInterface(pythoncom.IID_IDispatch)
                com_obj = win32com.client.Dispatch(dispatch)
                if hasattr(com_obj, "Documents") and hasattr(com_obj, "ActiveDocument"):
                    if com_obj.Documents.Count > 0:
                        return com_obj
                if name and (name.lower().endswith(".docx") or name.lower().endswith(".doc")):
                    monikers_to_retry.append((name, moniker))
            except Exception:
                try:
                    ctx = pythoncom.CreateBindCtx(0)
                    name = moniker.GetDisplayName(ctx, None)
                    if name and (name.lower().endswith(".docx") or name.lower().endswith(".doc")):
                        monikers_to_retry.append((name, moniker))
                except Exception:
                    pass
                continue

        for name, moniker in monikers_to_retry:
            try:
                obj = rot.GetObject(moniker)
                dispatch = obj.QueryInterface(pythoncom.IID_IDispatch)
                doc = win32com.client.Dispatch(dispatch)
                app = doc.Application
                if app.Documents.Count > 0:
                    return app
            except Exception:
                continue
    except Exception:
        pass
    return None


def find_document(app, filename: str = None):
    """Find an open document by filename.

    Args:
        app: Word.Application COM object.
        filename: Document name (basename) or full path.
                  If None or empty, returns the active document.

    Returns:
        Document COM object.

    Raises:
        ValueError: If the document is not found or no documents are open.
    """
    if app.Documents.Count == 0:
        raise WordError("no_document", "No documents are open in Word", hint="Open a document in Word, then retry.")

    if not filename:
        return app.ActiveDocument

    target_basename = unicodedata.normalize('NFC', os.path.basename(filename)).lower()
    target_fullpath = (
        unicodedata.normalize('NFC', os.path.normpath(filename)).lower()
        if os.path.isabs(filename) else None
    )

    for i in range(1, app.Documents.Count + 1):
        doc = app.Documents(i)
        if unicodedata.normalize('NFC', doc.Name).lower() == target_basename:
            return doc
        if target_fullpath and unicodedata.normalize('NFC', os.path.normpath(doc.FullName)).lower() == target_fullpath:
            return doc

    open_docs = [app.Documents(i).Name for i in range(1, app.Documents.Count + 1)]
    raise WordError(
        "document_not_found",
        f"Document '{filename}' is not open in Word. Open documents: {open_docs}",
        hint="Pass one of the open document names, or omit filename to use the active document.",
    )


@contextmanager
def undo_record(app, name: str):
    """Wrap a block of COM mutations in a single Word UndoRecord.

    Groups all changes into one Ctrl+Z entry in Word's undo stack.
    The undo record name appears in Edit > Undo and in the undo history.
    Degrades gracefully on Word 2007 or earlier (no UndoRecord support).

    Args:
        app: Word.Application COM object.
        name: Label for the undo entry (truncated to 64 chars by Word).

    Usage::

        with undo_record(app, "MCP: Insert Text"):
            doc.Range(0, 0).InsertBefore("Hello")
    """
    rec = None
    try:
        rec = app.UndoRecord
        if rec.IsRecordingCustomRecord:
            try:
                rec.EndCustomRecord()
            except Exception:
                pass
        rec.StartCustomRecord(name[:64])
    except Exception:
        rec = None
    try:
        yield
    finally:
        if rec is not None:
            try:
                rec.EndCustomRecord()
            except Exception:
                pass
