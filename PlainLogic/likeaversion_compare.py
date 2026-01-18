# likeaversion_compare.py
#
# Compares two likeaversion_index.json files
# Outputs a structured comparison report
# Comparator output is self-describing and downstream-safe

import json
from pathlib import Path
from typing import Dict, Any


def load_index(path: Path) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def normalize_ladder_name(name: str) -> str:
    if name.endswith(".clean.txt"):
        return name[:-10]
    return name


def compare_revisions(
    left_idx: Dict[str, Any],
    right_idx: Dict[str, Any],
    left_index_path: Path,
    right_index_path: Path,
    include_all_rungs: bool = False,
) -> Dict[str, Any]:
    report = {
        "left": {
            "location": left_idx["location"],
            "revision": left_idx["revision"],
            "path": left_idx["path"],
            "program_id": f'{left_idx["location"]}/{left_idx["revision"]}',
            "likeaversion_index": str(left_index_path.resolve()),
        },
        "right": {
            "location": right_idx["location"],
            "revision": right_idx["revision"],
            "path": right_idx["path"],
            "program_id": f'{right_idx["location"]}/{right_idx["revision"]}',
            "likeaversion_index": str(right_index_path.resolve()),
        },
        "summary": {
            "status": "PASS",
            "ladder_differences": 0,
        },
        "ladders": {},
    }

    left_ladders = left_idx.get("ladders", {})
    right_ladders = right_idx.get("ladders", {})
    all_ladders = sorted(set(left_ladders) | set(right_ladders))

    for ladder_file in all_ladders:
        ladder_id = normalize_ladder_name(ladder_file)

        entry = {
            "ladder_id": ladder_id,
            "file": ladder_file,
            "status": "identical",
            "structural": "same",
            "parameters": "same",
            "left_ladder_struct_hash": None,
            "right_ladder_struct_hash": None,
            "left_ladder_param_hash": None,
            "right_ladder_param_hash": None,
            "rungs": {},
        }

        l = left_ladders.get(ladder_file)
        r = right_ladders.get(ladder_file)

        if l is None:
            entry["status"] = "missing_left"
        elif r is None:
            entry["status"] = "missing_right"
        else:
            entry["left_ladder_struct_hash"] = l.get("ladder_structural_hash")
            entry["right_ladder_struct_hash"] = r.get("ladder_structural_hash")
            entry["left_ladder_param_hash"] = l.get("ladder_parameter_hash")
            entry["right_ladder_param_hash"] = r.get("ladder_parameter_hash")

            if l.get("ladder_structural_hash") != r.get("ladder_structural_hash"):
                entry["structural"] = "different"
                entry["status"] = "logic_difference"
            elif l.get("ladder_parameter_hash") != r.get("ladder_parameter_hash"):
                entry["parameters"] = "different"
                entry["status"] = "parameter_difference"

        # Rung drilldown:
        # - always if include_all_rungs
        # - or if ladder differs/missing
        if include_all_rungs or entry["status"] != "identical":
            left_rungs = (l or {}).get("rungs", {})
            right_rungs = (r or {}).get("rungs", {})
            all_rungs = sorted(set(left_rungs) | set(right_rungs))

            for rung_id in all_rungs:
                lr = left_rungs.get(rung_id)
                rr = right_rungs.get(rung_id)

                if lr is None:
                    entry["rungs"][rung_id] = {"rung": rung_id, "status": "missing_left"}
                    continue
                if rr is None:
                    entry["rungs"][rung_id] = {"rung": rung_id, "status": "missing_right"}
                    continue

                rung_entry = {
                    "rung": rung_id,
                    "status": "identical",
                    "structural": "same",
                    "parameters": "same",
                    "left_structural_hash": lr.get("structural_hash"),
                    "right_structural_hash": rr.get("structural_hash"),
                    "left_parameter_hash": lr.get("parameter_hash"),
                    "right_parameter_hash": rr.get("parameter_hash"),
                    "left_tokens": lr.get("tokens", []),
                    "right_tokens": rr.get("tokens", []),
                }

                if lr.get("structural_hash") != rr.get("structural_hash"):
                    rung_entry["structural"] = "different"
                if lr.get("parameter_hash") != rr.get("parameter_hash"):
                    rung_entry["parameters"] = "different"

                if rung_entry["structural"] != "same" or rung_entry["parameters"] != "same":
                    rung_entry["status"] = "different"

                # If not including all, only keep differences
                if include_all_rungs or rung_entry["status"] != "identical":
                    entry["rungs"][rung_id] = rung_entry

        if entry["status"] != "identical":
            report["summary"]["status"] = "FAIL"
            report["summary"]["ladder_differences"] += 1

        report["ladders"][ladder_id] = entry

    return report


def run(left_index_path: Path, right_index_path: Path, output_path: Path):
    left_idx = load_index(left_index_path)
    right_idx = load_index(right_index_path)

    report = compare_revisions(left_idx, right_idx, left_index_path, right_index_path)

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2)

    return report
