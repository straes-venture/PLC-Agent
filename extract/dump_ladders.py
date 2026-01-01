import os
import time
import pyperclip
from pywinauto.keyboard import send_keys


SETTLE = 0.25


def dump_ladders_from_snapshot(snapshot: dict, out_dir: str):
    """
    Dump ladder logic using snapshot data.

    HARD RULES:
    - ONLY type == 'LADDER'
    - ONLY ladder numbers >= 2
    - [SYSTEM] is NEVER allowed (even if present upstream)
    """

    os.makedirs(out_dir, exist_ok=True)

    program_files = snapshot.get("program_files", [])

    ladders = [
        pf for pf in program_files
        if pf.get("type") == "LADDER"
        and isinstance(pf.get("number"), int)
        and pf["number"] >= 2
        and pf.get("name", "").upper() != "[SYSTEM]"
    ]

    dumped = []

    for lad in ladders:
        lad_num = lad["number"]
        lad_name = lad["name"]

        print(f"[DUMP] LAD {lad_num} — {lad_name}", flush=True)

        # Go To ladder
        send_keys("^g")
        time.sleep(SETTLE)

        send_keys(f"{lad_num}:0")
        time.sleep(SETTLE)

        send_keys("%g")   # Alt+G (Go)
        time.sleep(SETTLE)

        send_keys("%c")   # Alt+C (Close Go To)
        time.sleep(SETTLE)

        # Select all rungs (your proven method)
        send_keys("^{HOME}")
        time.sleep(SETTLE)

        send_keys("+^{END}")
        time.sleep(SETTLE)

        # Copy
        send_keys("^c")
        time.sleep(SETTLE)

        text = pyperclip.paste()

        if not text.strip():
            print(f"[WARN] Empty clipboard for LAD {lad_num}", flush=True)
            continue

        fname = f"LAD{lad_num:03d}.raw.txt"
        out_path = os.path.join(out_dir, fname)

        with open(out_path, "w", encoding="utf-8") as f:
            f.write(text)

        dumped.append({
            "number": lad_num,
            "name": lad_name,
            "file": fname,
            "bytes": len(text)
        })

    return dumped
