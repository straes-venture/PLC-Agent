import os
import re
import json
import time
from datetime import datetime

try:
    import fitz  # type: ignore
except Exception:
    fitz = None

try:
    import pdfplumber
except Exception:
    pdfplumber = None

# Use unified logger so messages go to the UI when a sink is installed
import util.logger as logger


# Close Acrobat if it is holding the PDF (best-effort)
def _ensure_acrobat_closed():
    try:
        from python.close_apps import close_acrobat
        try:
            close_acrobat()
        except Exception:
            pass
    except Exception:
        pass


def _wait_for_file_stable(path: str, timeout=120, stable_secs=2.0):
    """
    Wait until file size is unchanged for `stable_secs` seconds or timeout.
    Returns True if stable, False on timeout.
    """
    end = time.time() + timeout
    last_size = -1
    last_change = time.time()

    while time.time() < end:
        try:
            if os.path.exists(path):
                size = os.path.getsize(path)
                if size != last_size:
                    last_size = size
                    last_change = time.time()
                elif size > 0 and (time.time() - last_change) >= stable_secs:
                    return True
        except Exception:
            pass
        time.sleep(0.25)
    return False


def extract_pdf_text(pdf_path: str, progress_callback=None):
    """
    Extract text from PDF and return (text, page_count).
    If provided, progress_callback(current_page, total_pages) is called as pages are processed.
    Logs page count immediately after opening the file.
    """
    # Ensure Acrobat is not holding the file
    _ensure_acrobat_closed()

    # Log and wait for the file to be present/stable before opening
    logger.dbg(f"[SNAPSHOT] Waiting for PDF to stabilize: {pdf_path}")
    stable = _wait_for_file_stable(pdf_path, timeout=120, stable_secs=2.0)
    if not stable:
        logger.dbg(f"[SNAPSHOT] WARNING: PDF did not fully stabilize within timeout: {pdf_path}")

    # Try PyMuPDF first
    if fitz:
        logger.dbg("[SNAPSHOT] Extracting PDF text using PyMuPDF")
        doc = fitz.open(pdf_path)
        page_count = getattr(doc, "page_count", len(doc))
        logger.dbg(f"[SNAPSHOT] PDF page count: {page_count}")
        out = []
        # notify initial progress (0 / total)
        if progress_callback:
            try:
                progress_callback(0, page_count)
            except Exception:
                pass
        for i, page in enumerate(doc, start=1):
            out.append(page.get_text("text") or "")
            if progress_callback:
                try:
                    progress_callback(i, page_count)
                except Exception:
                    pass
            else:
                logger.dbg(f"[SNAPSHOT] Extracted pages {i} of {page_count}")
        return ("\n".join(out), page_count)

    # Fallback to pdfplumber
    if pdfplumber:
        logger.dbg("[SNAPSHOT] Extracting PDF text using pdfplumber")
        out = []
        try:
            with pdfplumber.open(pdf_path) as pdf:
                page_count = len(pdf.pages)
                logger.dbg(f"[SNAPSHOT] PDF page count: {page_count}")
                # notify initial progress (0 / total)
                if progress_callback:
                    try:
                        progress_callback(0, page_count)
                    except Exception:
                        pass
                for i, page in enumerate(pdf.pages, start=1):
                    try:
                        txt = page.extract_text() or ""
                        out.append(txt)
                        if progress_callback:
                            try:
                                progress_callback(i, page_count)
                            except Exception:
                                pass
                        else:
                            logger.dbg(f"[SNAPSHOT] pdfplumber extracted page {i}")
                    except Exception as e:
                        logger.dbg(f"[SNAPSHOT] pdfplumber failed on page {i} of {page_count}: {e}")
            return ("\n".join(out), page_count)
        except Exception as e:
            logger.dbg(f"[SNAPSHOT] Exception opening PDF with pdfplumber: {e}")
            raise

    raise RuntimeError(
        "No PDF backend available. Install one of:\n"
        "  py -3.10-32 -m pip install pymupdf\n"
        "  py -3.10-32 -m pip install pdfplumber"
    )


def _parse_program_file_list(lines):
    out = []
    in_section = False

    for raw in lines:
        line = (raw or "").strip()
        if not line:
            continue

        if line == "Program File List":
            in_section = True
            continue

        if in_section:
            if line.startswith("Data File List") or line.startswith("Page "):
                break

            if line.startswith("Name Number Type"):
                continue

            parts = line.split()
            # Accept lines with 5 tokens (no name) or more (name present).
            if len(parts) < 5:
                continue

            try:
                bytes_ = int(parts[-1])
                debug = parts[-2]
                rungs = int(parts[-3])
                typ = parts[-4]
                number = int(parts[-5])
                name = " ".join(parts[:-5]).strip()
            except Exception:
                continue

            # Explicitly ignore SYSTEM / non-ladder program files
            if typ != "LADDER" or number < 2:
                continue

            out.append({
                "number": number,
                "name": name,
                "type": typ,
                "rungs": rungs,
                "debug": debug,
                "bytes": bytes_
            })

    return out


def _parse_data_file_list(lines):
    out = []
    in_section = False

    for raw in lines:
        line = (raw or "").strip()
        if not line:
            continue

        if line == "Data File List":
            in_section = True
            continue

        if in_section:
            if line.startswith("Data File Information") or line.startswith("Program File Information"):
                break

            if line.startswith("Name Number Type"):
                continue

            parts = line.split()
            if len(parts) < 9:
                continue

            try:
                last = parts[-1]
                elements = int(parts[-2])
                words = int(parts[-3])
                debug = parts[-4]
                scope = parts[-5]
                typ = parts[-6]
                number = int(parts[-7])
                name = " ".join(parts[:-7]).strip()
            except Exception:
                continue

            out.append({
                "name": name,
                "number": number,
                "type": typ,
                "scope": scope,
                "debug": debug,
                "words": words,
                "elements": elements,
                "last": last
            })

    return out


def _parse_memory(lines):
    mem = {}
    for line in lines[:80]:
        m = re.search(
            r"Total Memory Used:\s*(\d+)\s*Instruction Words Used\s*-\s*(\d+)\s*Data Table Words Used",
            line
        )
        if m:
            mem["instruction_words_used"] = int(m.group(1))
            mem["data_table_words_used"] = int(m.group(2))

        m = re.search(
            r"Total Memory Left:\s*(\d+)\s*Instruction Words Left",
            line
        )
        if m:
            mem["instruction_words_left"] = int(m.group(1))

    return mem


def digest_program_snapshot(pdf_path: str, unit: str, revision: str, rss_hash: str, out_dir: str,
                           progress_callback=None, resume_event=None, stop_event=None) -> dict:
    """
    Digest the PDF and return a snapshot dict.
    progress_callback(current_page, total_pages) is called during extraction (if provided).
    resume_event / stop_event are optional threading.Events used to cooperate with UI pause/stop.
    """
    logger.dbg(f"[SNAPSHOT] Digesting snapshot from {os.path.basename(pdf_path)}")

    text, page_count = extract_pdf_text(pdf_path, progress_callback=progress_callback)
    lines = text.splitlines()

    os.makedirs(out_dir, exist_ok=True)

    debug_path = os.path.join(out_dir, "_debug_extracted_report.txt")
    with open(debug_path, "w", encoding="utf-8", errors="replace") as f:
        f.write(text)
    logger.dbg(f"[SNAPSHOT] Wrote extracted text → {debug_path}")

    snap = {
        "identity": {
            "unit": unit,
            "revision": revision,
            "rss_hash": rss_hash,
            "generated": datetime.utcnow().isoformat() + "Z",
            "processor": None,
            "processor_series": None
        },
        "page_count": page_count,
        "memory": _parse_memory(lines),
        "program_files": _parse_program_file_list(lines),
        "data_files": _parse_data_file_list(lines),
        "symbols": [],
        "io_config": [],
        "channel_config": {}
    }

    try:
        m2 = re.search(r"(Bul\.\s*1766.*MicroLogix.*Series\s*[A-Z])", text)
        if m2:
            snap["identity"]["processor"] = m2.group(1).strip()
    except Exception as e:
        logger.dbg(f"[SNAPSHOT] WARNING: failed to set processor identity: {e} (m2={repr(m2)})")

    try:
        m3 = re.search(r"Series\s*([A-Z])", text)
        if m3:
            snap["identity"]["processor_series"] = m3.group(1)
    except Exception as e:
        logger.dbg(f"[SNAPSHOT] WARNING: failed to set processor_series: {e} (m3={repr(m3)})")

    out_path = os.path.join(out_dir, "program_snapshot.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(snap, f, indent=2)

    logger.dbg(f"[SNAPSHOT] Wrote snapshot → {out_path}")

    # Cleanup: delete PDF after successful extraction
    try:
        os.remove(pdf_path)
        logger.dbg(f"[SNAPSHOT] Deleted source PDF → {pdf_path}")
    except Exception as e:
        logger.dbg(f"[SNAPSHOT] WARNING: Failed to delete PDF ({pdf_path}): {e}")

        # Attempt to close Acrobat and retry once, waiting up to 60s for the file to unlock.
        try:
            from python.close_apps import close_acrobat
            logger.dbg("[SNAPSHOT] Attempting to close Acrobat (retry-delete flow)")
            try:
                close_acrobat(timeout=8, kill_if_still_running=False)
            except Exception:
                pass
        except Exception:
            logger.dbg("[SNAPSHOT] close_acrobat not available for retry")

        end = time.time() + 60.0
        deleted = False
        while time.time() < end:
            # Honor stop request
            if stop_event is not None and stop_event.is_set():
                logger.dbg("[SNAPSHOT] Stop requested during delete-retry — aborting retries")
                break
            # If UI paused, wait until resumed before continuing retries
            if resume_event is not None and not resume_event.is_set():
                logger.dbg("[SNAPSHOT] Waiting for resume to continue delete attempts...")
                resume_event.wait()
                logger.dbg("[SNAPSHOT] Resumed — continuing delete attempts")
            try:
                os.remove(pdf_path)
                logger.dbg(f"[SNAPSHOT] Deleted source PDF after retry → {pdf_path}")
                deleted = True
                break
            except Exception:
                time.sleep(1.0)

        if not deleted:
            # Try one more close attempt then immediate delete
            try:
                from python.close_apps import close_acrobat
                logger.dbg("[SNAPSHOT] Second attempt to close Acrobat before final delete attempt")
                try:
                    close_acrobat(timeout=6, kill_if_still_running=False)
                except Exception:
                    pass
            except Exception:
                logger.dbg("[SNAPSHOT] close_acrobat not available for second attempt")

            try:
                os.remove(pdf_path)
                logger.dbg(f"[SNAPSHOT] Deleted source PDF after second attempt → {pdf_path}")
                deleted = True
            except Exception as e2:
                logger.dbg(f"[SNAPSHOT] PDF still locked after retries: {e2}")

        if not deleted:
            # Pause and await human interaction to resolve the lock. The UI exposes resume_event;
            # block here until resume_event is set by user (so they can close Acrobat and hit Resume).
            logger.dbg("[SNAPSHOT] Pausing and awaiting user to close Acrobat and resume processing")
            if resume_event is not None:
                # Wait until user resumes processing
                resume_event.wait()
                # After resume, try one final delete (best-effort)
                try:
                    os.remove(pdf_path)
                    logger.dbg(f"[SNAPSHOT] Deleted source PDF after user resume → {pdf_path}")
                except Exception as e3:
                    logger.dbg(f"[SNAPSHOT] Failed to delete PDF after user resume: {e3}")
            else:
                logger.dbg("[SNAPSHOT] No resume_event available — cannot pause for human interaction")

    return snap
