import os
import re
import json
import hashlib
from datetime import datetime
from pathlib import Path


def sha256_file(path: str, chunk_size: int = 1024 * 1024) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            chunk = f.read(chunk_size)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()


def normalize_unit_name(name: str) -> str:
    s = name.strip().upper()
    s = re.sub(r"[^\w]+", "_", s)
    s = re.sub(r"_+", "_", s).strip("_")
    return s or "UNKNOWN_UNIT"


def parse_unit_and_revision_from_filename(rss_path: str):
    base = Path(rss_path).stem

    for pat in (r"(\d{6})$", r"(\d{8})$"):
        m = re.search(pat, base)
        if m:
            return normalize_unit_name(base[: m.start(1)]), m.group(1)

    return normalize_unit_name(base), datetime.now().strftime("%Y%m%d%H%M%S")


def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)
