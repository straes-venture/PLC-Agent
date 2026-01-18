# PlainLogic.py
import logging
import sys
from pathlib import Path

import temp_ui


def setup_logging():
    log_dir = Path(__file__).parent / "logs"
    log_dir.mkdir(exist_ok=True)

    log_file = log_dir / "plainlogic.log"

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )


def main():
    setup_logging()
    logging.info("PlainLogic starting")

    try:
        temp_ui.launch()
    except Exception:
        logging.exception("Fatal error during PlainLogic execution")
        raise


if __name__ == "__main__":
    main()
