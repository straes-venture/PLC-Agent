import os
import time
import psutil
from pywinauto.application import Application
from pywinauto import Desktop

RS500_EXE = r"C:\Program Files (x86)\Rockwell Software\RSLogix 500 English\RS500.exe"
RS500_PROC = "rs500.exe"

POLL = 0.25
START_TIMEOUT = 30.0

# ------------------------------------------------------------------
# Debug sink plumbing (UI can inject logger)
# ------------------------------------------------------------------

_debug_sink = None


def set_debug_sink(fn):
    """
    Set a callable that receives debug strings.
    Example: set_debug_sink(ui.log_line)
    """
    global _debug_sink
    _debug_sink = fn


def dbg(msg: str):
    line = f"[RSLOGIX] {msg}"
    if _debug_sink:
        try:
            _debug_sink(line)
            return
        except Exception:
            pass
    print(line, flush=True)


# ------------------------------------------------------------------
# Detection helpers
# ------------------------------------------------------------------

def _find_rslogix_process():
    dbg("Checking for existing RSLogix processes")
    for p in psutil.process_iter(["pid", "name"]):
        try:
            name = (p.info.get("name") or "").lower()
            if name == RS500_PROC:
                dbg(f"Found RSLogix process: pid={p.info['pid']}")
                return p.info["pid"]
        except Exception:
            pass
    dbg("No RSLogix process found")
    return None


def _find_rslogix_window(timeout=0.5):
    dbg("Checking for RSLogix windows")
    try:
        win = Desktop(backend="win32").window(title_re=r".*RSLogix 500.*")
        if win.exists(timeout=timeout):
            title = (win.window_text() or "").strip()
            dbg(f"Found RSLogix window: '{title}'")
            return win
    except Exception:
        pass

    dbg("No RSLogix window found")
    return None


def _dismiss_activation_popup(timeout=5):
    end = time.time() + timeout
    while time.time() < end:
        try:
            dlg = Desktop(backend="win32").window(
                title_re=r".*Product Activation Failed.*"
            )
            if dlg.exists(timeout=0.2):
                dbg("Activation popup detected, dismissing")
                try:
                    dlg.set_focus()
                    dlg.type_keys("{ENTER}")
                except Exception:
                    pass
                return True
        except Exception:
            pass
        time.sleep(0.2)
    return False


# ------------------------------------------------------------------
# Core lifecycle logic
# ------------------------------------------------------------------

def ensure_rslogix_running():
    """
    Ensure RSLogix is running.
    Never intentionally starts a second instance.
    """
    pid = _find_rslogix_process()

    if pid:
        dbg("RSLogix already running; attaching to existing process")
        app = Application(backend="win32").connect(process=pid)

        end = time.time() + START_TIMEOUT
        while time.time() < end:
            _dismiss_activation_popup(timeout=1)
            win = _find_rslogix_window(timeout=0.2)
            if win:
                try:
                    win.set_focus()
                except Exception:
                    pass
                dbg("Attached to existing RSLogix window")
                return app, win
            time.sleep(POLL)

        raise TimeoutError("RSLogix process found but no window appeared")

    # ------------------------------------------------------------
    # No process found -> safe to start
    # ------------------------------------------------------------

    if not os.path.exists(RS500_EXE):
        raise FileNotFoundError(f"RSLogix executable not found: {RS500_EXE}")

    dbg("No RSLogix process detected; starting RSLogix")
    app = Application(backend="win32").start(f'"{RS500_EXE}"')

    end = time.time() + START_TIMEOUT
    while time.time() < end:
        _dismiss_activation_popup(timeout=1)
        win = _find_rslogix_window(timeout=0.2)
        if win:
            try:
                win.set_focus()
            except Exception:
                pass
            dbg("RSLogix started and window is ready")
            return app, win
        time.sleep(POLL)

    raise TimeoutError("RSLogix did not open a window within timeout")


# ------------------------------------------------------------------
# Open program
# ------------------------------------------------------------------

def _find_open_dialog(timeout=15):
    end = time.time() + timeout
    while time.time() < end:
        try:
            dlg = Desktop(backend="win32").window(
                class_name="#32770",
                title_re=r".*(Open/Import|Open).*"
            )
            if dlg.exists(timeout=0.2):
                dbg("Open dialog detected")
                return dlg
        except Exception:
            pass
        time.sleep(0.2)

    raise RuntimeError("Open dialog not found")


def _set_open_filename(dlg, path):
    try:
        dlg.Edit.set_edit_text(path)
        return
    except Exception:
        pass

    dlg.set_focus()
    dlg.type_keys("^a{BACKSPACE}")
    dlg.type_keys(path, with_spaces=True)


def open_program(rss_path: str):
    if not os.path.exists(rss_path):
        raise FileNotFoundError(rss_path)

    dbg(f"Preparing to open RSS: {rss_path}")

    app, main_win = ensure_rslogix_running()

    dbg("Sending Ctrl+O to open file")
    try:
        main_win.set_focus()
    except Exception:
        pass
    time.sleep(0.2)
    main_win.type_keys("^o")

    dlg = _find_open_dialog()
    dlg.set_focus()
    _set_open_filename(dlg, rss_path)
    dlg.type_keys("{ENTER}")

    _dismiss_activation_popup(timeout=5)

    dbg("RSS open sequence complete")
    time.sleep(2.0)

    return app
