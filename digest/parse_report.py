import os
import re
import json
from datetime import datetime
from util.common import ensure_dir


def utc_now():
    return datetime.utcnow().replace(microsecond=0).isoformat() + "Z"


def extract_text(pdf_path):
    try:
        import fitz  # PyMuPDF
        doc = fitz.open(pdf_path)
        return "\n".join(p.get_text("text") for p in doc)
    except Exception:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            return "\n".join(p.extract_text() or "" for p in pdf.pages)


# -------------------------
# PARSERS
# -------------------------

def parse_processor(text):
    m = re.search(r"(Bul\.\s*\d+\s+MicroLogix.+)", text)
    return m.group(1).strip() if m else None


def parse_program_name(text):
    m = re.search(r"Processor Name\s*:\s*(.+)", text)
    return m.group(1).strip() if m else None


def parse_memory(text):
    mem = {}
    for label, key in [
        ("Instruction Words Used", "instruction_words_used"),
        ("Data Table Words Used", "data_table_words_used"),
        ("Instruction Words Left", "instruction_words_left"),
    ]:
        m = re.search(rf"{label}\s*:\s*(\d+)", text)
        if m:
            mem[key] = int(m.group(1))
    return mem


def parse_program_files(text):
    """
    STRICT ladder rules:
      - Type must be LADDER
      - Number >= 2
      - Name must not be [SYSTEM]
    """
    files = []

    pat = re.compile(
        r"^\s*(\d+)\s+(LADDER|SYS)\s+(.+?)\s+(\d+)\s+(Yes|No)\s+(\d+)",
        re.MULTILINE,
    )

    for m in pat.finditer(text):
        number = int(m.group(1))
        ftype = m.group(2).strip()
        name = m.group(3).strip()

        if ftype != "LADDER":
            continue
        if number < 2:
            continue
        if name.upper() == "[SYSTEM]":
            continue

        files.append(
            {
                "number": number,
                "name": name,
                "type": "LADDER",
                "rungs": int(m.group(4)),
                "debug": m.group(5),
                "bytes": int(m.group(6)),
            }
        )

    return files


def parse_data_files(text):
    files = []
    pat = re.compile(
        r"^(B|N|F|ST|T|C)(\d+)\s+(\w+)\s+(\d+)\s*(.*)$",
        re.MULTILINE,
    )
    for m in pat.finditer(text):
        files.append(
            {
                "file": f"{m.group(1)}{m.group(2)}",
                "type": m.group(3),
                "elements": int(m.group(4)),
                "description": m.group(5).strip() or None,
            }
        )
    return files


# -------------------------
# MAIN ENTRY
# -------------------------

def build_program_snapshot(pdf_path, unit, revision, rss_hash, out_dir):
    ensure_dir(out_dir)

    text = extract_text(pdf_path)

    snapshot = {
        "schema_version": 1,
        "identity": {
            "unit": unit,
            "revision": revision,
            "program_name": parse_program_name(text) or unit,
            "processor": parse_processor(text),
            "generated_utc": utc_now(),
            "rss_hash": rss_hash,
        },
        "memory": parse_memory(text),
        "program_files": parse_program_files(text),
        "data_files": parse_data_files(text),
        "symbols": [],
        "io_config": [],
        "channel_config": {},
    }

    out_path = os.path.join(out_dir, "program_snapshot.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(snapshot, f, indent=2)

    return snapshot
