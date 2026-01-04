import os
import sys
from pathlib import Path

# Ensure project root is on sys.path so imports like "util.common" work
_here = Path(__file__).resolve().parent
_project_root = str(_here.parent)  # one level up from python/
if _project_root not in sys.path:
    sys.path.insert(0, _project_root)

from util.common import parse_unit_and_revision_from_filename

RSS_DIR = r"C:\PLC_Agent\rss"
EXPORTS_DIR = r"C:\PLC_Agent\exports"
MAX_SHOW = 50

def normalize(s):
    return "".join(c for c in s.lower() if c.isalnum() or c == "_")

def gather(path):
    out = []
    for root, _, files in os.walk(path):
        for fn in files:
            p = Path(root) / fn
            out.append(str(p))
    return out

def main():
    if not os.path.isdir(RSS_DIR):
        print("RSS_DIR not found:", RSS_DIR); return
    if not os.path.isdir(EXPORTS_DIR):
        print("EXPORTS_DIR not found:", EXPORTS_DIR); return

    rss_files = [f for f in os.listdir(RSS_DIR) if f.lower().endswith(".rss")]
    exports = gather(EXPORTS_DIR)

    print("RSS files (count={}):".format(len(rss_files)))
    for i, f in enumerate(sorted(rss_files)[:MAX_SHOW], 1):
        stem = Path(f).stem
        unit, rev = parse_unit_and_revision_from_filename(f)
        print(f"{i:2d}. {f}  -> stem='{stem}' unit='{unit}' rev='{rev}'")

    print("\nExport files (count={}):".format(len(exports)))
    for i, p in enumerate(sorted(exports)[:MAX_SHOW], 1):
        print(f"{i:2d}. {p}")

    print("\nAttempting matching (verbose):")
    export_names = [os.path.basename(p).lower() for p in exports]
    export_stems = [Path(p).stem.lower() for p in exports]

    for f in sorted(rss_files):
        stem = Path(f).stem.lower()
        f_lower = f.lower()
        matched = False
        reasons = []
        # direct equality
        if f_lower in export_names:
            matched = True
            reasons.append("exact filename match")
        if stem in export_stems:
            matched = True
            reasons.append("stem match")
        # contains
        if not matched:
            for en in export_names:
                if stem in en or f_lower in en:
                    matched = True
                    reasons.append(f"contains match against export name: '{en}'")
                    break
        # check unit/rev presence
        unit, rev = parse_unit_and_revision_from_filename(f)
        unit_norm = normalize(unit)
        rev_norm = normalize(rev)
        for en in export_names:
            nen = normalize(en)
            if unit_norm and unit_norm in nen:
                matched = True
                reasons.append(f"unit '{unit}' appears in export '{en}'")
                break
            if rev_norm and rev_norm in nen:
                matched = True
                reasons.append(f"rev '{rev}' appears in export '{en}'")
                break

        print(f"{f}: matched={matched}; reasons={reasons or ['no match']}")

if __name__ == "__main__":
    main()