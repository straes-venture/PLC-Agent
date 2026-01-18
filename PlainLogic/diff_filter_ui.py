# -*- coding: utf-8 -*-
# diff_filter_ui.py
#
# Popup UI for controlling which parameter/runtime differences
# are considered meaningful during comparison.
#
# Encoding: UTF-8 (no BOM)

import tkinter as tk
from tkinter import ttk


def open_diff_filter_dialog(parent, diff_filter, on_change):
    win = tk.Toplevel(parent)
    win.title("Difference Filters")
    win.geometry("520x420")
    win.transient(parent)
    win.grab_set()

    frm = ttk.Frame(win, padding=12)
    frm.pack(fill="both", expand=True)

    ttk.Label(frm, text="Structural logic (always considered):", font=("Segoe UI", 10, "bold")).pack(anchor="w")
    ttk.Label(
        frm,
        text="- Instruction sequence\n- Branching and flow\n- Operand positions",
    ).pack(anchor="w", padx=12, pady=(2, 8))

    ttk.Separator(frm).pack(fill="x", pady=8)

    ttk.Label(frm, text="Parameter differences (per feature):", font=("Segoe UI", 10, "bold")).pack(anchor="w")

    param_keys = [k for k in diff_filter.keys() if k != "RUNTIME_GROUP" and not k.endswith("_ACCUM") and not k.startswith("RT_")]
    if not param_keys:
        ttk.Label(frm, text="(No parameter features registered yet)").pack(anchor="w", padx=12)
    else:
        for key in sorted(param_keys):
            var = tk.BooleanVar(value=bool(diff_filter.get(key, False)))

            def _update(k=key, v=var):
                diff_filter[k] = v.get()
                on_change(diff_filter)

            ttk.Checkbutton(frm, text=key.replace("_", " "), variable=var, command=_update).pack(anchor="w", padx=12)

    ttk.Separator(frm).pack(fill="x", pady=8)

    ttk.Label(frm, text="Runtime/state differences (per feature):", font=("Segoe UI", 10, "bold")).pack(anchor="w")

    runtime_keys = [k for k in diff_filter.keys() if k.endswith("_ACCUM") or k.startswith("RT_")]
    if not runtime_keys:
        ttk.Label(frm, text="(No runtime features registered yet)").pack(anchor="w", padx=12)
    else:
        for key in sorted(runtime_keys):
            var = tk.BooleanVar(value=bool(diff_filter.get(key, False)))

            def _update(k=key, v=var):
                diff_filter[k] = v.get()
                on_change(diff_filter)

            ttk.Checkbutton(frm, text=key.replace("_", " "), variable=var, command=_update).pack(anchor="w", padx=12)

    ttk.Separator(frm).pack(fill="x", pady=10)

    ttk.Button(frm, text="Close", command=win.destroy).pack(anchor="e")
