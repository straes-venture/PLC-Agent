import os
from pywinauto import Desktop

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


def process_rss_file(rss_path: str):
    unit, revision = parse_unit_and_revision_from_filename(rss_path)
    logger.dbg(f"[AGENT] === {os.path.basename(rss_path)} -> {unit} {revision} ===")

    rss_hash = sha256_file(rss_path)

    export_dir = os.path.join(EXPORTS_DIR, unit, revision)
    ensure_dir(export_dir)

    ladder_out_dir = os.path.join(export_dir, "ladders")
    ensure_dir(ladder_out_dir)

    # --- OPEN PROGRAM ---
    app = open_program(rss_path)
    main_win = Desktop(backend="win32").window(title_re=r".*RSLogix 500.*")

    try:
        # --- PRINT REPORT ---
        pdf_path = os.path.join(REPORTS_DIR, f"{unit}_{revision}.report.pdf")
        logger.dbg(f"[AGENT] Using report template: {TEMPLATE_PATH}")
        print_report(
            app=app,
            main_win=main_win,
            pdf_out_path=pdf_path,
            template_path=TEMPLATE_PATH,
        )

        # --- SNAPSHOT / METADATA ---
        snapshot = digest_program_snapshot(
            pdf_path=pdf_path,
            unit=unit,
            revision=revision,
            rss_hash=rss_hash,
            out_dir=export_dir,
        )

        # --- LADDER EXTRACTION ---
        dump_ladders_from_snapshot(
            snapshot=snapshot,
            out_dir=ladder_out_dir
        )

        logger.dbg(f"[AGENT] Completed {unit} {revision}")

    except Exception as e:
        logger.dbg(f"[AGENT] FAILED {unit} {revision}: {e}")
        raise

    finally:
        # Always recover cleanly
        close_acrobat()
        close_rslogix_program_only()


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
