import os
import re
import json
from datetime import datetime

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

try:
    import pdfplumber
except Exception:
    pdfplumber = None


def dbg(msg: str):
    print(f"[SNAPSHOT] {msg}", flush=True)


def extract_pdf_text(pdf_path: str) -> str:
    if fitz:
        dbg("Extracting PDF text using PyMuPDF")
        doc = fitz.open(pdf_path)
        out = []
        for page in doc:
            out.append(page.get_text("text"))
        return "\n".join(out)

    if pdfplumber:
        dbg("Extracting PDF text using pdfplumber")
        out = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                out.append(page.extract_text() or "")
        return "\n".join(out)

    raise RuntimeError(
        "No PDF backend available. Install one of:\n"
        "  py -3.10-32 -m pip install pymupdf\n"
        "  py -3.10-32 -m pip install pdfplumber"
    )


def _parse_program_file_list(lines):
    out = []
    in_section = False

    for raw in lines:
        line = (raw or "").strip()
        if not line:
            continue

        if line == "Program File List":
            in_section = True
            continue

        if in_section:
            if line.startswith("Data File List") or line.startswith("Page "):
                break

            if line.startswith("Name Number Type"):
                continue

            parts = line.split()
            if len(parts) < 6:
                continue

            try:
                bytes_ = int(parts[-1])
                debug = parts[-2]
                rungs = int(parts[-3])
                typ = parts[-4]
                number = int(parts[-5])
                name = " ".join(parts[:-5]).strip()
            except Exception:
                continue

            # 🚫 Explicitly ignore SYSTEM / non-ladder program files
            if typ != "LADDER" or number < 2:
                continue

            out.append({
                "number": number,
                "name": name,
                "type": typ,
                "rungs": rungs,
                "debug": debug,
                "bytes": bytes_
            })

    return out


def _parse_data_file_list(lines):
    out = []
    in_section = False

    for raw in lines:
        line = (raw or "").strip()
        if not line:
            continue

        if line == "Data File List":
            in_section = True
            continue

        if in_section:
            if line.startswith("Data File Information") or line.startswith("Program File Information"):
                break

            if line.startswith("Name Number Type"):
                continue

            parts = line.split()
            if len(parts) < 9:
                continue

            try:
                last = parts[-1]
                elements = int(parts[-2])
                words = int(parts[-3])
                debug = parts[-4]
                scope = parts[-5]
                typ = parts[-6]
                number = int(parts[-7])
                name = " ".join(parts[:-7]).strip()
            except Exception:
                continue

            out.append({
                "name": name,
                "number": number,
                "type": typ,
                "scope": scope,
                "debug": debug,
                "words": words,
                "elements": elements,
                "last": last
            })

    return out


def _parse_memory(lines):
    mem = {}
    for line in lines[:80]:
        m = re.search(
            r"Total Memory Used:\s*(\d+)\s*Instruction Words Used\s*-\s*(\d+)\s*Data Table Words Used",
            line
        )
        if m:
            mem["instruction_words_used"] = int(m.group(1))
            mem["data_table_words_used"] = int(m.group(2))

        m = re.search(
            r"Total Memory Left:\s*(\d+)\s*Instruction Words Left",
            line
        )
        if m:
            mem["instruction_words_left"] = int(m.group(1))

    return mem


def digest_program_snapshot(pdf_path: str, unit: str, revision: str, rss_hash: str, out_dir: str) -> dict:
    dbg(f"Digesting snapshot from {os.path.basename(pdf_path)}")

    text = extract_pdf_text(pdf_path)
    lines = text.splitlines()

    os.makedirs(out_dir, exist_ok=True)

    debug_path = os.path.join(out_dir, "_debug_extracted_report.txt")
    with open(debug_path, "w", encoding="utf-8", errors="replace") as f:
        f.write(text)
    dbg(f"Wrote extracted text → {debug_path}")

    snapshot = {
        "identity": {
            "unit": unit,
            "revision": revision,
            "rss_hash": rss_hash,
            "generated": datetime.utcnow().isoformat() + "Z",
            "processor": None,
            "processor_series": None
        },
        "memory": _parse_memory(lines),
        "program_files": _parse_program_file_list(lines),
        "data_files": _parse_data_file_list(lines),
        "symbols": [],
        "io_config": [],
        "channel_config": {}
    }

    m2 = re.search(r"(Bul\.\s*1766.*MicroLogix.*Series\s*[A-Z])", text)
    if m2:
        snapshot["identity"]["processor"] = m2.group(1).strip()

    m3 = re.search(r"Series\s*([A-Z])", text)
    if m3:
        snapshot["identity"]["processor_series"] = m3.group(1)

    out_path = os.path.join(out_dir, "program_snapshot.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(snapshot, f, indent=2)

    dbg(f"Wrote snapshot → {out_path}")

    # ✅ CLEANUP: delete PDF after successful extraction
    try:
        os.remove(pdf_path)
        dbg(f"Deleted source PDF → {pdf_path}")
    except Exception as e:
        dbg(f"WARNING: Failed to delete PDF ({pdf_path}): {e}")

    return snapshot
