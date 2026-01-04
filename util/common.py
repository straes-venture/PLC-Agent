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
    """
    Parse a filename stem into (unit, revision).

    Supports forms like:
      - ERIE_251217_STARTUP -> unit='ERIE', revision='251217_STARTUP'
      - BROSS_W_210101       -> unit='BROSS_W', revision='210101'
      - UNIT251217           -> unit='UNIT', revision='251217'
    Falls back to returning the normalized stem as unit and a timestamp revision.
    """
    base = Path(rss_path).stem

    # Match: <unit><optional-sep><date6|date8><optional-sep><optional-suffix>
    m = re.match(r"^(?P<unit>.*?)(?:[_\-]?)(?P<date>\d{6}|\d{8})(?:[_\-]?(?P<suffix>.*))?$", base)
    if m and m.group("date"):
        unit_raw = m.group("unit") or ""
        date = m.group("date")
        suffix = m.group("suffix")
        unit = normalize_unit_name(unit_raw) if unit_raw else normalize_unit_name(base[: m.start("date")])
        revision = date + (f"_{suffix}" if suffix else "")
        return unit, revision

    return normalize_unit_name(base), datetime.now().strftime("%Y%m%d%H%M%S")


def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)
