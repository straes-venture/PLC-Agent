import os
from pywinauto import Desktop
import time

from python.open_program import open_program
from python.print_report import run as print_report
from python.close_apps import close_acrobat, close_rslogix_program_only

from extract.dump_ladders import dump_ladders_from_snapshot
from digest.program_snapshot import digest_program_snapshot
from util.common import (
    sha256_file,
    parse_unit_and_revision_from_filename,
    ensure_dir,
)
import util.logger as logger


RSS_DIR = r"C:\PLC_Agent\rss"
EXPORTS_DIR = r"C:\PLC_Agent\exports"
REPORTS_DIR = r"C:\PLC_Agent\reports"
TEMPLATE_PATH = r"C:\PLC_Agent\templates\Auto_Parse.RO0"


def process_rss_file(rss_path: str, resume_event=None, stop_event=None, restart_event=None,
                     progress_callback=None, ladder_progress_callback=None):
    import traceback
    def _check_pause_and_stop():
        if resume_event is not None:
            resume_event.wait()
        if stop_event is not None and stop_event.is_set():
            raise RuntimeError("Stop requested")

    unit, revision = parse_unit_and_revision_from_filename(rss_path)
    logger.dbg(f"[AGENT] === {os.path.basename(rss_path)} -> {unit} {revision} ===")

    rss_hash = sha256_file(rss_path)

    export_dir = os.path.join(EXPORTS_DIR, unit, revision)
    ensure_dir(export_dir)

    ladder_out_dir = os.path.join(export_dir, "ladders")
    ensure_dir(ladder_out_dir)

    app = None
    try:
        # --- OPEN PROGRAM ---
        logger.dbg("[AGENT] Opening RSLogix program")
        _check_pause_and_stop()
        app = open_program(
            rss_path,
            resume_event=resume_event,
            stop_event=stop_event,
            restart_event=restart_event,
        )
        logger.dbg("[AGENT] open_program returned")

        main_win = Desktop(backend="win32").window(title_re=r".*RSLogix 500.*")

        # --- PRINT REPORT ---
        logger.dbg("[AGENT] Starting print_report")
        _check_pause_and_stop()
        pdf_path = os.path.join(REPORTS_DIR, f"{unit}_{revision}.report.pdf")
        logger.dbg(f"[AGENT] Using report template: {TEMPLATE_PATH}")
        # print_report creates the PDF; digest will extract and report page progress via progress_callback
        print_report(
            app=app,
            main_win=main_win,
            pdf_out_path=pdf_path,
            template_path=TEMPLATE_PATH,
            resume_event=resume_event,
            stop_event=stop_event,
            restart_event=restart_event,
        )
        logger.dbg(f"[AGENT] print_report finished — pdf_out_path={pdf_path}")

        # --- SNAPSHOT / METADATA ---
        logger.dbg("[AGENT] Starting digest_program_snapshot")
        _check_pause_and_stop()
        snapshot = digest_program_snapshot(
            pdf_path=pdf_path,
            unit=unit,
            revision=revision,
            rss_hash=rss_hash,
            out_dir=export_dir,
            progress_callback=progress_callback,
            resume_event=resume_event,
            stop_event=stop_event,
        )
        logger.dbg("[AGENT] digest_program_snapshot finished")

        # --- LADDER EXTRACTION ---
        logger.dbg("[AGENT] Starting ladder extraction")
        _check_pause_and_stop()
        dump_ladders_from_snapshot(
            snapshot=snapshot,
            out_dir=ladder_out_dir,
            progress_callback=ladder_progress_callback,
        )
        logger.dbg("[AGENT] Ladder extraction finished")

        logger.dbg(f"[AGENT] Completed {unit} {revision}")

    except Exception as e:
        tb = traceback.format_exc()
        logger.dbg(f"[AGENT] FAILED {unit} {revision}: {e}\n{tb}")
        raise
    finally:
        # Best-effort: attempt to close Acrobat, ignore if it doesn't close. Do NOT force-kill.
        try:
            logger.dbg("[AGENT] Attempting to close Acrobat (best-effort)")
            close_acrobat(timeout=8, kill_if_still_running=False)
        except Exception:
            logger.dbg("[AGENT] Exception while attempting to close Acrobat (ignored)")
        try:
            close_rslogix_program_only()
        except Exception:
            logger.dbg("[AGENT] Exception while attempting to close RSLogix program (ignored)")


def main():
    ensure_dir(REPORTS_DIR)
    ensure_dir(EXPORTS_DIR)

    rss_files = [
        os.path.join(RSS_DIR, f)
        for f in os.listdir(RSS_DIR)
        if f.lower().endswith(".rss")
    ]

    if not rss_files:
        logger.dbg("[AGENT] No RSS files found.")
        return

    for rss in rss_files:
        try:
            process_rss_file(rss)
        except Exception:
            logger.dbg("[AGENT] Aborting current file; continuing safely.")


if __name__ == "__main__":
    main()
