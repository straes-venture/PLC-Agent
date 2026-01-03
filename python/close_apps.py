import time
import psutil
from pywinauto import Desktop
from pywinauto.keyboard import send_keys
import util.logger as logger


def _find_acrobat_candidate_windows():
    """
    Return list of top-level windows that look like Acrobat/Reader PDF windows.
    Heuristics: title contains '.pdf' OR title mentions 'adobe'/'acrobat' OR class name hints.
    """
    candidates = []
    try:
        all_wins = Desktop(backend="win32").windows()
    except Exception:
        return candidates

    for w in all_wins:
        try:
            txt = (w.window_text() or "").strip()
            cls = (w.class_name() or "").strip().lower()
            t = txt.lower()
            if ".pdf" in t or "adobe" in t or "acrobat" in t or "acro" in cls:
                candidates.append((w, txt))
        except Exception:
            pass
    return candidates


def _log_acrobat_processes():
    names = set()
    for p in psutil.process_iter(["name"]):
        try:
            n = (p.info.get("name") or "").lower()
            if "acro" in n or "adobe" in n:
                names.add(p.info.get("name"))
        except Exception:
            pass
    if names:
        logger.dbg(f"[CLOSE] Acrobat-like processes present: {', '.join(sorted(names))}")
    else:
        logger.dbg("[CLOSE] No Acrobat-like processes detected")


def close_acrobat(timeout=10, kill_if_still_running=False):
    """
    Close Acrobat windows that have the PDF open.
    More robust detection: tries strict match first, then heuristics ('.pdf' in title or process/class hints).
    If windows close but processes remain, we log processes. Optionally kill processes when kill_if_still_running=True.
    """
    end = time.time() + timeout

    try:
        # Try the original strict pattern first (covers typical title formats)
        win = Desktop(backend="win32").window(title_re=r".* - Adobe Acrobat.*")
        if not win.exists(timeout=0.2):
            # Fallback: scan candidate windows
            candidates = _find_acrobat_candidate_windows()
            if not candidates:
                logger.dbg("[CLOSE] Acrobat not detected (or already closed)")
                _log_acrobat_processes()
                return
            # pick the best candidate (prefer a window with .pdf in title)
            win = None
            for w, txt in candidates:
                if ".pdf" in (txt or "").lower():
                    win = w
                    break
            if win is None:
                win = candidates[0][0]

        # Attempt to close the found window
        logger.dbg("[CLOSE] Closing Acrobat")
        try:
            win.set_focus()
        except Exception:
            pass

        try:
            # Try Alt+F4 first
            win.type_keys("%{F4}")
        except Exception:
            try:
                send_keys("%{F4}")
            except Exception:
                pass

        time.sleep(0.8)

        # If "Save changes?" prompts appear, choose No
        prompt = Desktop(backend="win32").window(class_name="#32770", title_re=r".*(Save|Adobe Acrobat|Acrobat).*")
        if prompt.exists(timeout=0.8):
            try:
                prompt.type_keys("n")
            except Exception:
                try:
                    prompt.child_window(title_re=r".*No.*").click()
                except Exception:
                    pass

        # Wait briefly to confirm the specific window went away
        wait_end = time.time() + 2.0
        while time.time() < wait_end:
            try:
                if not win.exists(timeout=0.2):
                    logger.dbg("[CLOSE] Acrobat closed cleanly")
                    _log_acrobat_processes()
                    return
            except Exception:
                # If querying existence fails, just break to process-level check
                break
            time.sleep(0.2)

    except Exception:
        # fall through to process-level logging
        pass

    # If we reach here, the window may still exist or process still running
    logger.dbg("[CLOSE] Acrobat may not have closed (window still present or process lingering)")
    _log_acrobat_processes()

    if kill_if_still_running:
        # Dangerous: only used if caller explicitly requests force-kill
        for p in psutil.process_iter(["pid", "name"]):
            try:
                name = (p.info.get("name") or "").lower()
                if "acro" in name or "adobe" in name:
                    logger.dbg(f"[CLOSE] Attempting to terminate process: {p.info.get('name')} pid={p.info.get('pid')}")
                    try:
                        proc = psutil.Process(p.info.get("pid"))
                        proc.terminate()
                        proc.wait(timeout=3)
                        logger.dbg(f"[CLOSE] Terminated pid={p.info.get('pid')}")
                    except Exception as e:
                        logger.dbg(f"[CLOSE] Failed to terminate pid={p.info.get('pid')}: {e}")
            except Exception:
                pass


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
                logger.dbg("[CLOSE] RSLogix not detected")
                return

            win.set_focus()
            time.sleep(0.15)

            logger.dbg("[CLOSE] RSLogix File->Close (Alt+F, C)")
            send_keys("%f")
            time.sleep(0.10)
            send_keys("c")
            time.sleep(0.6)

            # Handle "Save changes?" prompt if it appears
            prompt = Desktop(backend="win32").window(class_name="#32770", title_re=r".*(RSLogix|Save|Confirm).*")
            if prompt.exists(timeout=0.8):
                logger.dbg("[CLOSE] Save prompt detected -> selecting No")
                try:
                    prompt.type_keys("n")
                except Exception:
                    try:
                        prompt.child_window(title_re=r".*No.*").click()
                    except Exception:
                        pass
                time.sleep(0.4)

            logger.dbg("[CLOSE] Program closed (RSLogix still running)")
            return

        except Exception:
            pass

        time.sleep(0.25)

    logger.dbg("[CLOSE] RSLogix close timed out")


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
                logger.dbg("[CLOSE] RSLogix detected -> Alt+F4 (exit app)")
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
        logger.dbg("[CLOSE] RSLogix closed cleanly")
    else:
        logger.dbg("[CLOSE] RSLogix close timed out (may already be closed)")
