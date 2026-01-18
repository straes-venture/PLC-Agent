# -*- coding: utf-8 -*-
# rung_index_ui.py
#
# Program alignment explorer:
# Family (footprint)
#   Ladder
#     Ladder Revision (rung count)
#       Rung Number
#         Rung Revision (structural hash)
#           Parameter Revision (parameter hash)

import tkinter as tk
from tkinter import ttk
from typing import Dict, Any, List, Tuple

from diff_filter_ui import open_diff_filter_dialog


def _family_label(footprint: Dict[str, Any], program_count: int) -> str:
    ladders = footprint.get("ladders")
    data_files = footprint.get("data_files")
    if ladders is None or data_files is None:
        return f"Family - unknown footprint ({program_count} programs)"
    return f"Family - {ladders} ladders, {data_files} data files ({program_count} programs)"


def _diff_tokens(base: List[str], other: List[str]) -> List[str]:
    out = []
    n = min(len(base), len(other))
    for i in range(n):
        if base[i] != other[i]:
            out.append(f"[{i:03d}] {base[i]} -> {other[i]}")
    if len(base) != len(other):
        out.append(f"[len] {len(base)} -> {len(other)}")
    return out


def _as_kv_pairs(x):
    if not x:
        return []
    if isinstance(x[0], (list, tuple)) and len(x[0]) == 2:
        return [(str(a), str(b)) for a, b in x]
    return [("GENERIC", str(v)) for v in x]


def _effective_param_sig(example: Dict[str, Any], diff_filter: Dict[str, bool]) -> List[str]:
    sig = []
    for k, v in _as_kv_pairs(example.get("parameter_tokens", [])):
        if diff_filter.get(k, True):
            sig.append(f"{k}={v}")
    for k, v in _as_kv_pairs(example.get("runtime_tokens", [])):
        if diff_filter.get(k, False):
            sig.append(f"{k}={v}")
    return sig


def launch(parent, rung_index: Dict[str, Any]):
    win = tk.Toplevel(parent)
    win.title("Rung Index Explorer")
    win.geometry("1400x760")

    diff_filter = {
        "GENERIC": True,
        "TON_TIMER": True,
        "TON_PRESET": True,
        "TON_ACCUM": False,
    }

    show_structural_only = tk.BooleanVar(value=False)

    controls = ttk.Frame(win)
    controls.pack(fill="x", padx=8, pady=(8, 0))

    ttk.Label(controls, text="Filters:").pack(side="left", padx=(0, 8))

    ttk.Checkbutton(
        controls,
        text="Show structural differences only",
        variable=show_structural_only,
        command=lambda: build_tree()
    ).pack(side="left", padx=(0, 12))

    ttk.Button(
        controls,
        text="Difference Filters…",
        command=lambda: open_diff_filter_dialog(win, diff_filter, lambda _: build_tree())
    ).pack(side="left")

    paned = ttk.Panedwindow(win, orient="horizontal")
    paned.pack(fill="both", expand=True, padx=8, pady=8)

    # ---- Tree
    left = ttk.Frame(paned)
    paned.add(left, weight=1)

    tree = ttk.Treeview(left)
    tree.pack(fill="both", expand=True)

    # ---- Detail pane
    right = ttk.Frame(paned)
    paned.add(right, weight=3)

    detail = tk.Text(right, wrap="none", font=("Consolas", 10))
    detail.pack(fill="both", expand=True)
    detail.config(state="disabled")

    # underline tag for differing tokens
    detail.tag_configure("diff_token", underline=True)

    # ---- helper to render token rows with selective underlining
    def _insert_token_row(prefix: str, tokens: List[str], diff_indices: set):
        detail.insert("end", prefix)
        for idx, tok in enumerate(tokens):
            tok_str = f"{tok:<12} "
            start = detail.index("end")
            detail.insert("end", tok_str)
            end = detail.index("end")
            if idx in diff_indices:
                detail.tag_add("diff_token", start, end)
        detail.insert("end", "\n")

    # ---- Build tree
    def build_tree():
        tree.delete(*tree.get_children())

        for fam_key, fam in rung_index.items():
            fam_iid = tree.insert(
                "",
                "end",
                text=_family_label(fam.get("footprint", {}), len(fam.get("programs", []))),
                open=False
            )

            for ladder_name, ladder in fam.get("ladders", {}).items():
                lad_iid = tree.insert(fam_iid, "end", text=ladder_name, open=False)

                for rung_count_str, rev in ladder.get("revisions", {}).items():
                    ladrev_iid = tree.insert(
                        lad_iid,
                        "end",
                        text=f"Ladder Rev - {rung_count_str} rungs",
                        open=False
                    )

                    for rung_no_str, rung_no_obj in rev.get("rungs", {}).items():
                        rung_revs = rung_no_obj.get("revisions", {})
                        structural_rev_ct = len(rung_revs)

                        eff_param_sigs = set()
                        for rr in rung_revs.values():
                            for pr in rr.get("parameters", {}).values():
                                ex = pr.get("example", {})
                                sig = tuple(_effective_param_sig(ex, diff_filter))
                                eff_param_sigs.add(sig)

                        has_struct_diff = structural_rev_ct > 1
                        has_param_diff = len(eff_param_sigs) > 1

                        if show_structural_only.get() and not has_struct_diff:
                            continue

                        rung_label = f"Rung {int(rung_no_str):04d}"
                        if not has_struct_diff and not has_param_diff:
                            rung_label += "  ✓"
                        else:
                            markers = []
                            if has_struct_diff:
                                markers.append(f"{structural_rev_ct} structural")
                            if has_param_diff:
                                markers.append("params/runtime")
                            rung_label += "  ≠ (" + ", ".join(markers) + ")"

                        rung_iid = tree.insert(
                            ladrev_iid,
                            "end",
                            text=rung_label,
                            open=False,
                            values=("rung", fam_key, ladder_name, rung_count_str, rung_no_str)
                        )

                        for idx, (struct_hash, rr) in enumerate(rung_revs.items(), start=1):
                            rr_iid = tree.insert(
                                rung_iid,
                                "end",
                                text=f"Rung Rev {idx} ({len(rr.get('programs', []))} programs)",
                                open=False,
                                values=("rung_rev", fam_key, ladder_name, rung_count_str, rung_no_str, struct_hash)
                            )

                            for pidx, (param_hash, pr) in enumerate(rr.get("parameters", {}).items(), start=1):
                                tree.insert(
                                    rr_iid,
                                    "end",
                                    text=f"Param Rev {pidx} ({len(pr.get('programs', []))} programs)",
                                    values=("param", fam_key, ladder_name, rung_count_str, rung_no_str, struct_hash, param_hash)
                                )

    # ---- Selection handler
    def on_select(event):
        item = tree.focus()
        vals = tree.item(item, "values")
        if not vals:
            return

        detail.config(state="normal")
        detail.delete("1.0", "end")

        # ---------- RUNG REV ----------
        if vals[0] == "rung_rev":
            _, fam_key, ladder_name, rung_count_str, rung_no_str, struct_hash = vals
            fam = rung_index[fam_key]
            rr = fam["ladders"][ladder_name]["revisions"][rung_count_str]["rungs"][rung_no_str]["revisions"][struct_hash]

            params = list(rr["parameters"].items())
            base_param_hash, base_pr = params[0]
            base_ex = base_pr["example"]
            base_tokens = base_ex.get("tokens", [])
            base_sig = _effective_param_sig(base_ex, diff_filter)

            footprint = fam.get("footprint", {})
            detail.insert("end", f"Family Footprint: {footprint.get('ladders')} ladders, {footprint.get('data_files')} data files\n")
            detail.insert("end", f"Ladder: {ladder_name}\n")
            detail.insert("end", f"Ladder Revision: {rung_count_str} rungs\n")
            detail.insert("end", f"Rung Number: {rung_no_str}\n")
            detail.insert("end", f"Structural Hash: {struct_hash}\n\n")

            detail.insert("end", "Revisions:\n")
            for idx, (param_hash, pr) in enumerate(params, start=1):
                progs = ", ".join(pr.get("programs", []))
                detail.insert("end", f"  Rev {idx}  {param_hash}  ({progs})\n")

            detail.insert("end", "\nToken alignment:\n\n")

            for idx, (param_hash, pr) in enumerate(params, start=1):
                ex = pr["example"]
                tokens = ex.get("tokens", [])

                diff_indices = set()
                if idx > 1:
                    for i, (a, b) in enumerate(zip(base_tokens, tokens)):
                        if a != b:
                            diff_indices.add(i)
                    if len(tokens) != len(base_tokens):
                        diff_indices.update(range(min(len(tokens), len(base_tokens)), max(len(tokens), len(base_tokens))))

                _insert_token_row(f"Rev {idx}: ", tokens, diff_indices)

            detail.insert("end", "\nParameter/runtime differences vs baseline (respecting Difference Filters):\n")
            for idx, (param_hash, pr) in enumerate(params, start=1):
                if idx == 1:
                    continue
                sig = _effective_param_sig(pr["example"], diff_filter)
                diffs = _diff_tokens(base_sig, sig)
                for d in diffs:
                    detail.insert("end", f"  {d}\n")

            for idx, (param_hash, pr) in enumerate(params, start=1):
                raw = pr["example"].get("raw") or " ".join(pr["example"].get("tokens", []))
                detail.insert("end", f"\nRaw Rung, Rev {idx}:\n  {raw}\n")

            detail.config(state="disabled")
            return

        # ---------- PARAM ----------
        if vals[0] == "param":
            _, fam_key, ladder_name, rung_count_str, rung_no_str, struct_hash, param_hash = vals
            fam = rung_index[fam_key]
            pr = fam["ladders"][ladder_name]["revisions"][rung_count_str]["rungs"][rung_no_str]["revisions"][struct_hash]["parameters"][param_hash]

            detail.insert("end", "Raw Rung:\n")
            detail.insert("end", f"  {pr['example'].get('raw')}\n")

            detail.config(state="disabled")
            return

    tree.bind("<<TreeviewSelect>>", on_select)

    build_tree()
