import time
from pywinauto import Desktop
from pywinauto.keyboard import send_keys


def _dbg(msg: str):
    print(f"[CLOSE] {msg}", flush=True)


def close_acrobat(timeout=10):
    """
    Close Acrobat windows that have the PDF open.
    """
    try:
        win = Desktop(backend="win32").window(title_re=r".* - Adobe Acrobat.*")
        if win.exists(timeout=0.2):
            _dbg("Closing Acrobat")
            win.set_focus()
            win.type_keys("%{F4}")  # Alt+F4
            time.sleep(0.8)

            # If "Save changes?" prompts appear, choose No
            prompt = Desktop(backend="win32").window(class_name="#32770", title_re=r".*(Save|Adobe Acrobat).*")
            if prompt.exists(timeout=0.8):
                try:
                    prompt.type_keys("n")
                except Exception:
                    pass

            _dbg("Acrobat closed cleanly")
            return
    except Exception:
        pass

    _dbg("Acrobat not detected (or already closed)")


def close_rslogix_program_only(timeout=10):
    """
    Close the currently open RSS inside RSLogix WITHOUT exiting the app.
    Sequence: Alt+F, then C (File -> Close)
    Dismiss save prompts if they appear.
    """
    end = time.time() + timeout

    while time.time() < end:
        try:
            win = Desktop(backend="win32").window(title_re=r".*RSLogix 500.*")
            if not win.exists(timeout=0.2):
                _dbg("RSLogix not detected")
                return

            win.set_focus()
            time.sleep(0.15)

            _dbg("RSLogix File->Close (Alt+F, C)")
            send_keys("%f")
            time.sleep(0.10)
            send_keys("c")
            time.sleep(0.6)

            # Handle "Save changes?" prompt if it appears
            prompt = Desktop(backend="win32").window(class_name="#32770", title_re=r".*(RSLogix|Save|Confirm).*")
            if prompt.exists(timeout=0.8):
                _dbg("Save prompt detected -> selecting No")
                try:
                    prompt.type_keys("n")
                except Exception:
                    try:
                        prompt.child_window(title_re=r".*No.*").click()
                    except Exception:
                        pass
                time.sleep(0.4)

            _dbg("Program closed (RSLogix still running)")
            return

        except Exception:
            pass

        time.sleep(0.25)

    _dbg("Program-only close timed out")


def close_rslogix(timeout=12):
    """
    Full application exit (kept for optional use).
    RSLogix closes reliably with Alt+F4.
    Dismiss save prompts if they appear.
    """
    end = time.time() + timeout
    closed = False

    while time.time() < end:
        try:
            win = Desktop(backend="win32").window(title_re=r".*RSLogix 500.*")
            if win.exists(timeout=0.2):
                _dbg("RSLogix detected -> Alt+F4 (exit app)")
                win.set_focus()
                win.type_keys("%{F4}")
                time.sleep(0.8)

                # Handle "Save changes?" prompt
                prompt = Desktop(backend="win32").window(class_name="#32770", title_re=r".*(RSLogix|Save|Confirm).*")
                if prompt.exists(timeout=0.8):
                    try:
                        prompt.type_keys("n")
                    except Exception:
                        try:
                            prompt.child_window(title_re=r".*No.*").click()
                        except Exception:
                            pass

                time.sleep(0.6)
                continue
            else:
                closed = True
                break
        except Exception:
            pass

        time.sleep(0.3)

    if closed:
        _dbg("RSLogix closed cleanly")
    else:
        _dbg("RSLogix close timed out (may already be closed)")
