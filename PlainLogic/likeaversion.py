# -*- coding: utf-8 -*-
# likeaversion.py
#
# Build a structural/parameter/runtime index of ladder logic.
#
# OPTION A IMPLEMENTATION:
# - Read LADxxx.raw.txt directly
# - Split rungs using SOR/EOR
# - Assign rung numbers by physical order (0000, 0001, ...)
#
# TON v1 classification:
#   TON <timer> <timebase> <preset> <accum>
#     structural: TON, timebase
#     parameter : TON_TIMER, TON_PRESET
#     runtime   : TON_ACCUM

import json
import hashlib
from pathlib import Path
from typing import Dict, List, Any, Callable, Optional, Tuple


def _hash_tokens(tokens: List[str]) -> str:
    h = hashlib.sha1()
    for t in tokens:
        h.update(t.encode("utf-8"))
        h.update(b"\0")
    return h.hexdigest()[:8]


def _tokenize(rung_text: str) -> List[str]:
    return rung_text.strip().split()


def _split_rungs(raw_text: str) -> List[str]:
    """
    Split ladder raw text into rung strings using SOR/EOR.
    Keeps order. Drops final 'SOR END EOR' rung if present.
    """
    tokens = raw_text.split()
    rungs: List[List[str]] = []
    cur: Optional[List[str]] = None

    for tok in tokens:
        if tok == "SOR":
            cur = []
        elif tok == "EOR":
            if cur is not None:
                if not (len(cur) == 1 and cur[0] == "END"):
                    rungs.append(cur)
            cur = None
        else:
            if cur is not None:
                cur.append(tok)

    return [" ".join(r) for r in rungs]


def _is_addr_token(t: str) -> bool:
    return (":" in t) or ("/" in t) or t.startswith("#")


def _classify_tokens(tokens: List[str]) -> Tuple[List[str], List[Tuple[str, str]], List[Tuple[str, str]]]:
    """
    Returns:
      structural_tokens: list[str] with '_' placeholders for non-structural operands
      parameter_tokens : list[(key,value)] for configurable operands
      runtime_tokens   : list[(key,value)] for stateful operands
    """
    structural: List[str] = []
    params: List[Tuple[str, str]] = []
    runtime: List[Tuple[str, str]] = []

    i = 0
    n = len(tokens)

    while i < n:
        t = tokens[i]

        # TON <timer> <timebase> <preset> <accum>
        if t == "TON" and i + 4 < n:
            timer = tokens[i + 1]
            timebase = tokens[i + 2]
            preset = tokens[i + 3]
            accum = tokens[i + 4]

            structural.append("TON")
            params.append(("TON_TIMER", timer))
            structural.append("_")

            structural.append(timebase)

            params.append(("TON_PRESET", preset))
            structural.append("_")

            runtime.append(("TON_ACCUM", accum))
            structural.append("_")

            i += 5
            continue

        # Default heuristic behavior (unchanged intent):
        # address-ish tokens are parameters, everything else is structural
        if _is_addr_token(t):
            params.append(("GENERIC", t))
            structural.append("_")
        else:
            structural.append(t)

        i += 1

    return structural, params, runtime


def process_ladder_file(path: Path) -> Dict[str, Any]:
    raw_text = path.read_text(encoding="utf-8", errors="ignore")
    rung_texts = _split_rungs(raw_text)

    ladder: Dict[str, Any] = {"rungs": {}}

    for idx, rung_text in enumerate(rung_texts):
        rung_id = f"{idx:04d}"
        tokens = _tokenize(rung_text)

        structural_tokens, parameter_tokens, runtime_tokens = _classify_tokens(tokens)

        # JSON will serialize tuples as lists; we hash as "K=V" strings to be stable.
        param_sig = [f"{k}={v}" for (k, v) in parameter_tokens]

        ladder["rungs"][rung_id] = {
            "raw": rung_text,
            "tokens": tokens,
            "structural_tokens": structural_tokens,
            "parameter_tokens": parameter_tokens,
            "runtime_tokens": runtime_tokens,
            "structural_hash": _hash_tokens(structural_tokens),
            "parameter_hash": _hash_tokens(param_sig),
        }

    return ladder


def run(root_dir: Path, progress_cb: Optional[Callable] = None):
    """
    Scan all program revisions under root_dir and generate likeaversion.json
    in each revision directory.
    """
    revisions = sorted(p for p in root_dir.rglob("*") if p.is_dir() and (p / "ladders").is_dir())
    total = len(revisions)

    for idx, rev_dir in enumerate(revisions, start=1):
        ladders_dir = rev_dir / "ladders"
        ladder_files = sorted(ladders_dir.glob("LAD*.raw.txt"))
        if not ladder_files:
            continue

        if progress_cb:
            progress_cb(idx, total, "Indexing", str(rev_dir))

        out: Dict[str, Any] = {
            "location": rev_dir.parent.name,
            "revision": rev_dir.name,
            "ladders": {}
        }

        for lad_file in ladder_files:
            ladder_name = lad_file.stem.replace(".raw", "")
            out["ladders"][ladder_name] = process_ladder_file(lad_file)

        (rev_dir / "likeaversion.json").write_text(json.dumps(out, indent=2), encoding="utf-8")
