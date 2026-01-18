# -*- coding: utf-8 -*-
# rung_index.py
#
# Same structure as your working version; only minor tolerance improvements.

import json
from collections import defaultdict
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple


def load_json(path: Path) -> Dict[str, Any]:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def _read_program_snapshot(snapshot_path: Path) -> Optional[Dict[str, Any]]:
    if not snapshot_path.exists():
        return None
    try:
        snap = load_json(snapshot_path)
        program_files = snap.get("program_files", []) or []
        data_files = snap.get("data_files", []) or []
        identity = snap.get("identity", {}) or {}
        return {
            "ladders": len(program_files),
            "data_files": len(data_files),
            "processor": identity.get("processor"),
            "snapshot_path": str(snapshot_path),
        }
    except Exception:
        return None


def _parse_rung_no(rung_id: str) -> Optional[int]:
    try:
        return int(str(rung_id).strip())
    except Exception:
        return None


def build_rung_index_from_files(index_files: List[Path]) -> Dict[str, Any]:
    families: Dict[Tuple[Optional[int], Optional[int]], Dict[str, Any]] = {}
    snapshot_cache: Dict[str, Optional[Dict[str, Any]]] = {}

    def get_family(ladders_count: Optional[int], data_files_count: Optional[int]) -> Dict[str, Any]:
        key = (ladders_count, data_files_count)
        if key not in families:
            families[key] = {
                "footprint": {"ladders": ladders_count, "data_files": data_files_count},
                "programs": set(),
                "program_stats": {},
                "ladders": defaultdict(
                    lambda: {
                        "revisions": defaultdict(
                            lambda: {
                                "programs": set(),
                                "rungs": defaultdict(
                                    lambda: {
                                        "revisions": defaultdict(
                                            lambda: {
                                                "programs": set(),
                                                "parameters": defaultdict(lambda: {"programs": set(), "example": None}),
                                            }
                                        )
                                    }
                                ),
                            }
                        )
                    }
                ),
            }
        return families[key]

    for idx_path in index_files:
        likea = load_json(idx_path)
        program_id = f'{likea["location"]}/{likea["revision"]}'

        revision_dir = idx_path.parent
        cache_key = str(revision_dir.resolve())
        if cache_key not in snapshot_cache:
            snapshot_cache[cache_key] = _read_program_snapshot(revision_dir / "program_snapshot.json")
        snap = snapshot_cache[cache_key]

        ladders_count = snap["ladders"] if snap else None
        data_files_count = snap["data_files"] if snap else None

        fam = get_family(ladders_count, data_files_count)
        fam["programs"].add(program_id)
        if snap:
            fam["program_stats"][program_id] = snap

        ladders = likea.get("ladders", {}) or {}

        for ladder_name, ladder in ladders.items():
            rung_map = (ladder.get("rungs", {}) or {})
            rung_count = len(rung_map)

            lad_rev = fam["ladders"][ladder_name]["revisions"][rung_count]
            lad_rev["programs"].add(program_id)

            for rung_id, rung in rung_map.items():
                rung_no = _parse_rung_no(rung_id)
                if rung_no is None:
                    rung_no = -1

                struct_hash = rung.get("structural_hash")
                param_hash = rung.get("parameter_hash")

                rr = lad_rev["rungs"][rung_no]["revisions"][struct_hash]
                rr["programs"].add(program_id)

                pr = rr["parameters"][param_hash]
                pr["programs"].add(program_id)

                if pr["example"] is None:
                    pr["example"] = {
                        "raw": rung.get("raw", ""),
                        "tokens": rung.get("tokens", []) or [],
                        "structural_tokens": rung.get("structural_tokens", []) or [],
                        "parameter_tokens": rung.get("parameter_tokens", []) or [],
                        "runtime_tokens": rung.get("runtime_tokens", []) or [],
                    }

    out: Dict[str, Any] = {}

    def fam_sort_key(k: Tuple[Optional[int], Optional[int]]):
        ladders_count, data_files_count = k
        known = 0 if (ladders_count is not None and data_files_count is not None) else 1
        return (
            known,
            ladders_count if ladders_count is not None else 10**9,
            data_files_count if data_files_count is not None else 10**9,
        )

    for fam_key in sorted(families.keys(), key=fam_sort_key):
        fam = families[fam_key]
        ladders_count, data_files_count = fam_key

        ladders_out: Dict[str, Any] = {}
        for ladder_name, ladder_obj in fam["ladders"].items():
            revs_out: Dict[str, Any] = {}
            for rung_count, rev in ladder_obj["revisions"].items():
                rungs_out: Dict[str, Any] = {}
                for rung_no, rung_no_obj in rev["rungs"].items():
                    revs2_out: Dict[str, Any] = {}
                    for struct_hash, rr in rung_no_obj["revisions"].items():
                        params_out: Dict[str, Any] = {}
                        for param_hash, pr in rr["parameters"].items():
                            params_out[param_hash] = {
                                "programs": sorted(pr["programs"]),
                                "example": pr["example"],
                            }
                        revs2_out[struct_hash] = {
                            "programs": sorted(rr["programs"]),
                            "parameters": params_out,
                        }
                    rungs_out[str(rung_no)] = {"revisions": revs2_out}
                revs_out[str(rung_count)] = {
                    "programs": sorted(rev["programs"]),
                    "rungs": rungs_out,
                }
            ladders_out[ladder_name] = {"revisions": revs_out}

        out[str(fam_key)] = {
            "footprint": {"ladders": ladders_count, "data_files": data_files_count},
            "programs": sorted(fam["programs"]),
            "program_stats": fam["program_stats"],
            "ladders": ladders_out,
        }

    return out
