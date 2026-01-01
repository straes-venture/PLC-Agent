import os
import sys
import re
import json
import hashlib
from collections import defaultdict
from datetime import datetime

EXPORTS_DIR = r"C:\PLC_Agent\exports"


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def normalize_unit(name: str) -> str:
    name = name.upper()
    name = re.sub(r"[^\w]+", "_", name)
    return name.strip("_")


def parse_unit_and_revision(filename: str):
    base = os.path.splitext(filename)[0]
    m = re.search(r"(\d{6,8})$", base)
    if m:
        return normalize_unit(base[:m.start(1)]), m.group(1)
    return normalize_unit(base), None


def has_been_processed(unit, revision):
    path = os.path.join(EXPORTS_DIR, unit, revision)
    return os.path.exists(os.path.join(path, "program_snapshot.json"))


def scan_rss_directory(rss_dir: str, auto_delete_bak=True):
    print(f"[SCAN] {rss_dir}")

    results = {
        "scanned_utc": datetime.utcnow().isoformat() + "Z",
        "units": {},
        "bak_deleted": [],
        "bak_remaining": [],
    }

    for fname in os.listdir(rss_dir):
        path = os.path.join(rss_dir, fname)

        if not os.path.isfile(path):
            continue

        if fname.lower().endswith(".bak"):
            if auto_delete_bak:
                os.remove(path)
                results["bak_deleted"].append(fname)
            else:
                results["bak_remaining"].append(fname)

    grouped = defaultdict(list)

    for fname in os.listdir(rss_dir):
        if not fname.lower().endswith(".rss"):
            continue

        path = os.path.join(rss_dir, fname)
        unit, rev = parse_unit_and_revision(fname)
        file_hash = sha256(path)

        grouped[unit].append({
            "filename": fname,
            "base_revision": rev,
            "hash": file_hash,
            "mtime": os.path.getmtime(path),
        })

    for unit, files in grouped.items():
        seen_hashes = {}
        enriched = []

        for f in sorted(files, key=lambda x: x["mtime"], reverse=True):
            rev = f["base_revision"] or "UNKNOWN"

            if rev in seen_hashes and seen_hashes[rev] != f["hash"]:
                ts = datetime.fromtimestamp(f["mtime"]).strftime("%Y%m%d_%H%M%S")
                rev = f"{rev}_{ts}"

            seen_hashes[rev] = f["hash"]

            enriched.append({
                **f,
                "revision": rev,
                "processed": has_been_processed(unit, rev),
            })

        results["units"][unit] = enriched

    out_path = os.path.join(rss_dir, "scan_results.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2)

    print(f"[SCAN] wrote {out_path}")
    return results


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: py scan_rss.py <rss_dir>")
        sys.exit(1)

    scan_rss_directory(sys.argv[1])
