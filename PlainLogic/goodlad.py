# goodlad.py
import logging
from pathlib import Path

def normalize_lad_text(raw_text: str) -> list[str]:
    rungs = []
    current = []

    tokens = raw_text.split()

    for token in tokens:
        if token == "SOR":
            current = []
        elif token == "EOR":
            if current:
                rung_text = " ".join(current).strip()
                if rung_text != "END":   # ← filter END rung
                    rungs.append(rung_text)
            current = []
        else:
            current.append(token)

    return rungs

def clean_name_from_raw(path: Path) -> Path:
    name = path.name
    if name.endswith(".raw.txt"):
        base = name[:-8]
    else:
        base = path.stem
    return path.with_name(f"{base}.clean.txt")

def process_lad_file(raw_path: Path):
    text = raw_path.read_text(encoding="utf-8", errors="ignore")
    rungs = normalize_lad_text(text)

    clean_path = clean_name_from_raw(raw_path)

    with open(clean_path, "w", encoding="utf-8") as f:
        for idx, rung in enumerate(rungs):
            f.write(f"{idx:04d} | {rung}\n")

    return len(rungs), clean_path

def run(root_dir: Path, force: bool = False, progress_cb=None):
    logging.info(f"GoodLad scanning root: {root_dir} (force={force})")

    raw_files = sorted(root_dir.rglob("*.raw.txt"))
    total = len(raw_files)

    if total == 0:
        logging.warning("No .raw.txt files found")
        return

    processed = 0

    for idx, raw in enumerate(raw_files, start=1):
        clean = clean_name_from_raw(raw)

        if clean.exists() and not force:
            logging.info(f"Skipping (exists): {clean}")
            if progress_cb:
                progress_cb(idx, total, raw.name, "skipped")
            continue

        logging.info(f"Processing: {raw}")
        rung_count, out_path = process_lad_file(raw)
        processed += 1

        if progress_cb:
            progress_cb(idx, total, raw.name, f"{rung_count} rungs")

    logging.info(f"GoodLad complete — {processed} files generated")
