# updated ui.py — full UI with cooperative processing, progress UI, and robust logging
import os
import threading
import ctypes
from ctypes import wintypes
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import json
from pathlib import Path
import re
import time
import traceback
import faulthandler

from util.common import parse_unit_and_revision_from_filename
import util.logger as logger

# perform console-hiding side-effect on import (can be commented while debugging)
import python.hide_console  # module executes and hides console on import

# Ensure crash directory exists and enable faulthandler to persist C-level tracebacks
try:
    os.makedirs(r"C:\PLC_Agent", exist_ok=True)
    _crash_fh = open(r"C:\PLC_Agent\plc-agent-crash.log", "a", buffering=1, encoding="utf-8")
    faulthandler.enable(file=_crash_fh)
except Exception:
    _crash_fh = None

# Defer heavy/fragile imports until needed (lazy)
process_rss_file = None
open_program = None
print_report = None


def _lazy_imports():
    """
    Import agent/open_program/print_report on first use.
    Any import error is logged via logger.dbg so startup doesn't crash silently.
    """
    global process_rss_file, open_program, print_report
    if process_rss_file is None:
        try:
            from agent import process_rss_file as _pr
            process_rss_file = _pr
        except Exception as e:
            logger.dbg(f"Lazy import failed: agent.process_rss_file: {e}")
    if open_program is None:
        try:
            import python.open_program as _op
            open_program = _op
        except Exception as e:
            logger.dbg(f"Lazy import failed: python.open_program: {e}")
    if print_report is None:
        try:
            import python.print_report as _pr2
            print_report = _pr2
        except Exception as e:
            logger.dbg(f"Lazy import failed: python.print_report: {e}")


# ------------------------------------------------------------------
# Native Windows Folder Picker (STA COM)
# ------------------------------------------------------------------

_BIF_RETURNONLYFSDIRS = 0x00000001
_BIF_NEWDIALOGSTYLE = 0x00000040

_shell32 = ctypes.WinDLL("shell32", use_last_error=True)
_ole32 = ctypes.WinDLL("ole32", use_last_error=True)

LPITEMIDLIST = ctypes.c_void_p
HRESULT = ctypes.c_long


class BROWSEINFO(ctypes.Structure):
    _fields_ = [
        ("hwndOwner", wintypes.HWND),
        ("pidlRoot", LPITEMIDLIST),
        ("pszDisplayName", wintypes.LPWSTR),
        ("lpszTitle", wintypes.LPCWSTR),
        ("ulFlags", wintypes.UINT),
        ("lpfn", ctypes.c_void_p),
        ("lParam", wintypes.LPARAM),
        ("iImage", ctypes.c_int),
    ]


_shell32.SHBrowseForFolderW.argtypes = [ctypes.POINTER(BROWSEINFO)]
_shell32.SHBrowseForFolderW.restype = LPITEMIDLIST

_shell32.SHGetPathFromIDListW.argtypes = [LPITEMIDLIST, wintypes.LPWSTR]
_shell32.SHGetPathFromIDListW.restype = wintypes.BOOL

_ole32.CoTaskMemFree.argtypes = [ctypes.c_void_p]
_ole32.CoTaskMemFree.restype = None

_ole32.CoInitializeEx.argtypes = [ctypes.c_void_p, wintypes.DWORD]
_ole32.CoInitializeEx.restype = HRESULT

_ole32.CoUninitialize.argtypes = []
_ole32.CoUninitialize.restype = None

COINIT_APARTMENTTHREADED = 0x2


def _browse_for_folder_blocking(title="Select Folder"):
    _ole32.CoInitializeEx(None, COINIT_APARTMENTTHREADED)

    buf = ctypes.create_unicode_buffer(260)
    display = ctypes.create_unicode_buffer(260)

    bi = BROWSEINFO()
    bi.hwndOwner = None
    bi.pidlRoot = None
    bi.pszDisplayName = ctypes.cast(display, wintypes.LPWSTR)
    bi.lpszTitle = title
    bi.ulFlags = _BIF_RETURNONLYFSDIRS | _BIF_NEWDIALOGSTYLE
    bi.lpfn = None
    bi.lParam = 0
    bi.iImage = 0

    pidl = _shell32.SHBrowseForFolderW(ctypes.byref(bi))
    if not pidl:
        _ole32.CoUninitialize()
        return None

    ok = _shell32.SHGetPathFromIDListW(pidl, ctypes.cast(buf, wintypes.LPWSTR))
    _ole32.CoTaskMemFree(pidl)
    _ole32.CoUninitialize()

    return buf.value if ok else None


# ------------------------------------------------------------------
# Main UI
# ------------------------------------------------------------------

class PLCUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PLC Agent")
        self.geometry("1200x700")

        # primary paths (moved to settings)
        self.rss_dir = tk.StringVar()
        self.exports_dir = tk.StringVar()
        self.reports_dir = tk.StringVar()
        self.templates_dir = tk.StringVar()
        self.auto_delete_bak = tk.BooleanVar(value=True)

        # resolve and prepare overrides path
        self._overrides_path = self._resolve_overrides_path()
        self.overrides = self._load_overrides()

        # convenience references
        self.file_overrides = self.overrides.setdefault("files", {})
        self.unit_map = self.overrides.setdefault("unit_map", {})
        self.rev_map = self.overrides.setdefault("rev_map", {})
        self.paths = self.overrides.setdefault("paths", {})

        # populate path vars from persisted paths (fallbacks kept)
        self.rss_dir.set(self.paths.get("rss_dir", r"C:\PLC_Agent\rss"))
        self.exports_dir.set(self.paths.get("exports_dir", r"C:\PLC_Agent\exports"))
        self.reports_dir.set(self.paths.get("reports_dir", r"C:\PLC_Agent\reports"))
        self.templates_dir.set(self.paths.get("templates_dir", r"C:\PLC_Agent\templates"))
        self.auto_delete_bak.set(self.paths.get("auto_delete_bak", True))

        # show verbose logging (persisted)
        self.show_verbose = tk.BooleanVar(value=self.paths.get("show_verbose", False))

        # processing control primitives
        self._resume_event = threading.Event()
        self._resume_event.set()
        self._stop_requested = threading.Event()
        self._restart_current_request = threading.Event()
        self._current_item = None
        self._current_fname = None

        # single-run guard: current processing thread and a lock protecting assignment/check
        self._processing_thread = None
        self._processing_lock = threading.Lock()

        # missing-pdf prompt synchronization
        self._missing_pdf_event = None
        self._missing_pdf_action = None

        self._build_ui()

        # install logger sink that filters verbose messages for UI
        try:
            logger.set_debug_sink(self._logger_sink)
        except Exception:
            # best-effort: ignore if logger not ready
            pass

        try:
            exists = os.path.isfile(self._overrides_path)
            self.log_line(f"Overrides path: {self._overrides_path} (exists={exists})")
        except Exception:
            logger.dbg(f"Overrides path: {self._overrides_path}")

    # ----------------------------------------------------------------
    # UI logger sink (filters verbose messages from showing in UI)
    # Messages prefixed with "[VERBOSE]" are considered verbose.
    # They are still written to the durable log file by util.logger.dbg.
    # ----------------------------------------------------------------
    def _logger_sink(self, msg: str):
        try:
            if isinstance(msg, str) and msg.startswith("[VERBOSE]") and not bool(self.show_verbose.get()):
                return
            self.log_line(msg)
        except Exception:
            # never allow logging to raise
            pass

    # -------------------------
    # Overrides persistence
    # -------------------------
    def _resolve_overrides_path(self):
        try:
            if os.name == "nt":
                common_base = os.getenv("PROGRAMDATA") or os.path.join(os.sep, "ProgramData")
                common_dir = os.path.join(common_base, "PLC-Agent")
                try:
                    os.makedirs(common_dir, exist_ok=True)
                    test_path = os.path.join(common_dir, ".write_test")
                    with open(test_path, "w", encoding="utf-8") as fh:
                        fh.write("ok")
                    os.remove(test_path)
                    return os.path.join(common_dir, "overrides.json")
                except Exception:
                    pass
                base = os.getenv("LOCALAPPDATA") or os.path.expanduser("~")
                cfg_dir = os.path.join(base, "PLC-Agent")
                os.makedirs(cfg_dir, exist_ok=True)
                return os.path.join(cfg_dir, "overrides.json")
            else:
                candidates = ["/etc/plc-agent", "/var/lib/plc-agent"]
                for d in candidates:
                    try:
                        os.makedirs(d, exist_ok=True)
                        test_path = os.path.join(d, ".write_test")
                        with open(test_path, "w", encoding="utf-8") as fh:
                            fh.write("ok")
                        os.remove(test_path)
                        return os.path.join(d, "overrides.json")
                    except Exception:
                        continue
                base = os.getenv("XDG_CONFIG_HOME") or os.path.expanduser("~/.config")
                cfg_dir = os.path.join(base, "plc-agent")
                os.makedirs(cfg_dir, exist_ok=True)
                return os.path.join(cfg_dir, "overrides.json")
        except Exception:
            try:
                return os.path.join(os.path.dirname(__file__), "overrides.json")
            except Exception:
                return "overrides.json"

    def _load_overrides(self):
        try:
            if os.path.isfile(self._overrides_path):
                with open(self._overrides_path, "r", encoding="utf-8") as fh:
                    data = json.load(fh)
                    if isinstance(data, dict) and ("files" in data or "unit_map" in data or "rev_map" in data or "paths" in data):
                        return data
                    return {"files": data, "unit_map": {}, "rev_map": {}, "paths": {}}
        except Exception as e:
            try:
                self.log_line(f"Failed to load overrides from {self._overrides_path}: {e}")
            except Exception:
                logger.dbg(f"Failed to load overrides: {e}")
        return {"files": {}, "unit_map": {}, "rev_map": {}, "paths": {}}

    def _save_overrides(self):
        try:
            self.paths["rss_dir"] = self.rss_dir.get()
            self.paths["exports_dir"] = self.exports_dir.get()
            self.paths["reports_dir"] = self.reports_dir.get()
            self.paths["templates_dir"] = self.templates_dir.get()
            self.paths["auto_delete_bak"] = bool(self.auto_delete_bak.get())
            self.paths["show_verbose"] = bool(self.show_verbose.get())
        except Exception:
            pass
        try:
            tmp = self._overrides_path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as fh:
                json.dump(self.overrides, fh, indent=2, ensure_ascii=False)
            os.replace(tmp, self._overrides_path)
        except Exception as e:
            try:
                self.log_line(f"ERROR saving overrides to {self._overrides_path}: {e}")
            except Exception:
                logger.dbg(f"ERROR saving overrides: {e}")

    # -------------------------
    # UI construction
    # -------------------------
    def _build_ui(self):
        root = ttk.Frame(self)
        root.pack(fill="both", expand=True)

        left = ttk.Frame(root, width=260)
        left.pack(side="left", fill="y", padx=8, pady=8)
        left.pack_propagate(False)

        ttk.Label(left, text="Directories are managed in Settings").pack(anchor="w", pady=(0, 6))

        actions = ttk.Frame(left)
        actions.pack(fill="both", expand=True, pady=(6, 0))

        ttk.Button(actions, text="Scan", command=self.scan).pack(fill="x", pady=(2, 2))
        ttk.Button(actions, text="Scan Exports...", command=self.scan_exports).pack(fill="x", pady=(2, 2))

        # keep references so we can disable while processing
        self.process_selected_btn = ttk.Button(actions, text="Process Selected", command=self.process_selected)
        self.process_selected_btn.pack(fill="x", pady=(2, 2))
        self.process_new_btn = ttk.Button(actions, text="Process NEW Only", command=self.process_new_only)
        self.process_new_btn.pack(fill="x", pady=(2, 8))

        self.pause_btn = ttk.Button(actions, text="Pause", command=self._toggle_pause)
        self.pause_btn.pack(fill="x", pady=(2, 2))
        self.cancel_btn = ttk.Button(actions, text="Cancel", command=self._on_cancel_clicked)
        self.cancel_btn.pack(fill="x", pady=(2, 8))
        ttk.Button(actions, text="Settings...", command=self._open_settings_dialog).pack(fill="x", pady=(2, 2))

        # Progress widgets (stacked top->bottom): programs, pdf pages, ladders
        # placed at the bottom of the left column per request
        progress_frame = ttk.Frame(left)
        progress_frame.pack(side="bottom", fill="x", pady=(6, 0), padx=(4, 0))

        # Programs
        prog_row = ttk.Frame(progress_frame)
        prog_row.pack(fill="x", pady=(2,2))
        self.prog_label = ttk.Label(prog_row, text="", width=28, anchor="w")
        self.prog_label.pack(side="left", fill="x", expand=True)
        self.prog_bar = ttk.Progressbar(prog_row, orient="horizontal", length=180, mode="determinate")
        self.prog_bar.pack(side="right")

        # PDF pages
        pdf_row = ttk.Frame(progress_frame)
        pdf_row.pack(fill="x", pady=(2,2))
        self.pdf_label = ttk.Label(pdf_row, text="", width=28, anchor="w")
        self.pdf_label.pack(side="left", fill="x", expand=True)
        self.pdf_bar = ttk.Progressbar(pdf_row, orient="horizontal", length=180, mode="determinate")
        self.pdf_bar.pack(side="right")

        # Ladders
        ladder_row = ttk.Frame(progress_frame)
        ladder_row.pack(fill="x", pady=(2,2))
        self.ladder_label = ttk.Label(ladder_row, text="", width=28, anchor="w")
        self.ladder_label.pack(side="left", fill="x", expand=True)
        self.ladder_bar = ttk.Progressbar(ladder_row, orient="horizontal", length=180, mode="determinate")
        self.ladder_bar.pack(side="right")

        # Right area contains tree + log
        right = ttk.Frame(root)
        right.pack(side="left", fill="both", expand=True, padx=8, pady=8)

        tree_frame = ttk.LabelFrame(right, text="RSS Files")
        tree_frame.pack(fill="both", expand=True)

        columns = ("unit", "revision", "status", "filename")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="extended")
        self.tree["displaycolumns"] = ("unit", "revision", "status")

        for col, text, w in [
            ("unit", "Unit", 240),
            ("revision", "Revision", 120),
            ("status", "Status", 120),
            ("filename", "Filename", 0),
        ]:
            self.tree.heading(col, text=text)
            if col == "filename":
                self.tree.column(col, width=0, stretch=False)
            else:
                self.tree.column(col, width=w, anchor="w")

        yscroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=yscroll.set)

        self.tree.pack(side="left", fill="both", expand=True)
        yscroll.pack(side="right", fill="y")

        self.tree.bind("<Double-1>", self._on_tree_double_click)

        log_frame = ttk.LabelFrame(right, text="Log")
        log_frame.pack(fill="both", pady=(8, 0))

        self.log = tk.Text(log_frame, height=12, state="disabled", wrap="word")
        self.log.pack(fill="both", expand=True)

    # -------------------------
    # Settings dialog and helpers
    # -------------------------
    def _open_settings_dialog(self):
        dlg = tk.Toplevel(self)
        dlg.title("Settings")
        dlg.transient(self)
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.geometry("+%d+%d" % (self.winfo_rootx() + 120, self.winfo_rooty() + 120))

        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill="both", expand=True)

        tmp_rss = tk.StringVar(value=self.rss_dir.get())
        tmp_exports = tk.StringVar(value=self.exports_dir.get())
        tmp_reports = tk.StringVar(value=self.reports_dir.get())
        tmp_templates = tk.StringVar(value=self.templates_dir.get())
        tmp_auto_delete = tk.BooleanVar(value=self.auto_delete_bak.get())
        tmp_verbose = tk.BooleanVar(value=self.show_verbose.get())

        ttk.Label(frm, text="Directories and preferences").grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(frm, text=" ").grid(row=1, column=0)

        ttk.Label(frm, text="RSS Directory:").grid(row=2, column=0, sticky="w")
        ttk.Entry(frm, textvariable=tmp_rss, width=60).grid(row=2, column=1, sticky="we")
        ttk.Button(frm, text="Browse...", command=lambda: self._browse_and_set(tmp_rss, "Select RSS Folder")).grid(row=2, column=2, padx=(8, 0))

        ttk.Label(frm, text="Exports Directory:").grid(row=3, column=0, sticky="w", pady=(6,0))
        ttk.Entry(frm, textvariable=tmp_exports, width=60).grid(row=3, column=1, sticky="we", pady=(6,0))
        ttk.Button(frm, text="Browse...", command=lambda: self._browse_and_set(tmp_exports, "Select Exports Folder")).grid(row=3, column=2, padx=(8, 0), pady=(6,0))

        ttk.Label(frm, text="Reports Directory:").grid(row=4, column=0, sticky="w", pady=(6,0))
        ttk.Entry(frm, textvariable=tmp_reports, width=60).grid(row=4, column=1, sticky="we", pady=(6,0))
        ttk.Button(frm, text="Browse...", command=lambda: self._browse_and_set(tmp_reports, "Select Reports Folder")).grid(row=4, column=2, padx=(8, 0), pady=(6,0))

        ttk.Label(frm, text="Templates Directory:").grid(row=5, column=0, sticky="w", pady=(6,0))
        ttk.Entry(frm, textvariable=tmp_templates, width=60).grid(row=5, column=1, sticky="we", pady=(6,0))
        ttk.Button(frm, text="Browse...", command=lambda: self._browse_and_set(tmp_templates, "Select Templates Folder")).grid(row=5, column=2, padx=(8, 0), pady=(6,0))

        ttk.Checkbutton(frm, text="Auto-delete _BAK files on scan", variable=tmp_auto_delete).grid(row=6, column=0, columnspan=3, sticky="w", pady=(8,0))
        ttk.Checkbutton(frm, text="Show Verbose Logging", variable=tmp_verbose).grid(row=7, column=0, columnspan=3, sticky="w", pady=(4,0))

        btn_frame = ttk.Frame(frm)
        btn_frame.grid(row=8, column=0, columnspan=3, pady=(12, 0), sticky="e")

        def on_ok():
            self.rss_dir.set(tmp_rss.get())
            self.exports_dir.set(tmp_exports.get())
            self.reports_dir.set(tmp_reports.get())
            self.templates_dir.set(tmp_templates.get())
            self.auto_delete_bak.set(bool(tmp_auto_delete.get()))
            self.show_verbose.set(bool(tmp_verbose.get()))
            self.paths["rss_dir"] = self.rss_dir.get()
            self.paths["exports_dir"] = self.exports_dir.get()
            self.paths["reports_dir"] = self.reports_dir.get()
            self.paths["templates_dir"] = self.templates_dir.get()
            self.paths["auto_delete_bak"] = bool(self.auto_delete_bak.get())
            self.paths["show_verbose"] = bool(self.show_verbose.get())
            self._save_overrides()
            self.log_line("Settings saved")
            dlg.destroy()

        def on_cancel():
            dlg.destroy()

        ttk.Button(btn_frame, text="OK", command=on_ok).pack(side="right", padx=(6, 0))
        ttk.Button(btn_frame, text="Cancel", command=on_cancel).pack(side="right")

    # -------------------------
    # Helpers for browse, logging, tree edit
    # -------------------------
    def _browse_and_set(self, var: tk.StringVar, title: str):
        def worker():
            path = _browse_for_folder_blocking(title)
            self.after(0, lambda: self._on_browse_done(var, path))
        threading.Thread(target=worker, daemon=True).start()

    def _on_browse_done(self, var: tk.StringVar, path: str):
        if path:
            var.set(path)
            self._save_overrides()
            self.log_line(f"Saved path: {path}")

    def log_line(self, msg: str):
        stamp = datetime.now().strftime("[%H:%M:%S] ")
        line = stamp + msg
        def append():
            self.log.configure(state="normal")
            self.log.insert("end", line + "\n")
            self.log.see("end")
            self.log.configure(state="disabled")
        self.after(0, append)

    def _on_tree_double_click(self, event):
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return
        col = self.tree.identify_column(event.x)
        if not col:
            return
        try:
            col_index = int(col.strip("#")) - 1
        except Exception:
            return
        cols = self.tree["columns"]
        if col_index < 0 or col_index >= len(cols):
            return
        col_name = cols[col_index]
        if col_name not in ("unit", "revision"):
            return
        self._open_edit_dialog(row_id)

    def _open_edit_dialog(self, row_id):
        vals = list(self.tree.item(row_id, "values"))
        if len(vals) < 4:
            return
        cur_unit, cur_rev, status, filename = vals
        try:
            orig_unit, orig_rev = parse_unit_and_revision_from_filename(filename)
        except Exception:
            orig_unit = orig_rev = None

        dlg = tk.Toplevel(self)
        dlg.title("Edit Unit / Revision")
        dlg.transient(self)
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.geometry("+%d+%d" % (self.winfo_rootx() + 100, self.winfo_rooty() + 100))

        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Filename:").grid(row=0, column=0, sticky="w")
        ttk.Label(frm, text=filename).grid(row=0, column=1, sticky="w")

        ttk.Label(frm, text="Unit:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        unit_entry = ttk.Entry(frm, width=60)
        unit_entry.grid(row=1, column=1, sticky="we", pady=(8, 0))
        unit_entry.insert(0, cur_unit)

        ttk.Label(frm, text="Revision:").grid(row=2, column=0, sticky="w", pady=(8, 0))
        rev_entry = ttk.Entry(frm, width=20)
        rev_entry.grid(row=2, column=1, sticky="w", pady=(8, 0))
        rev_entry.insert(0, cur_rev)

        note = ttk.Label(frm, text="Note: changes will be applied to all files with the same original parsed unit/revision and persisted.")
        note.grid(row=3, column=0, columnspan=2, pady=(8, 0), sticky="w")

        btn_frame = ttk.Frame(frm)
        btn_frame.grid(row=4, column=0, columnspan=2, pady=(12, 0), sticky="e")

        def _apply_unit_map(orig_u, new_u):
            if not orig_u or orig_u == new_u:
                return
            self.unit_map[orig_u] = new_u
            for child in self.tree.get_children():
                child_vals = list(self.tree.item(child, "values"))
                if len(child_vals) >= 4:
                    cfn = child_vals[3]
                    try:
                        c_orig_unit, _ = parse_unit_and_revision_from_filename(cfn)
                    except Exception:
                        c_orig_unit = None
                    if c_orig_unit == orig_u:
                        child_vals[0] = new_u
                        self.tree.item(child, values=child_vals)

        def _apply_rev_map(orig_r, new_r):
            if not orig_r or orig_r == new_r:
                return
            self.rev_map[orig_r] = new_r
            for child in self.tree.get_children():
                child_vals = list(self.tree.item(child, "values"))
                if len(child_vals) >= 4:
                    cfn = child_vals[3]
                    try:
                        _, c_orig_rev = parse_unit_and_revision_from_filename(cfn)
                    except Exception:
                        c_orig_rev = None
                    if c_orig_rev == orig_r:
                        child_vals[1] = new_r
                        self.tree.item(child, values=child_vals)

        def on_ok():
            new_unit = unit_entry.get().strip()
            new_rev = rev_entry.get().strip()
            vals[0] = new_unit
            vals[1] = new_rev
            self.tree.item(row_id, values=vals)
            fo = self.file_overrides.get(filename, {})
            fo.update({"unit": new_unit, "revision": new_rev})
            self.file_overrides[filename] = fo
            try:
                if orig_unit:
                    _apply_unit_map(orig_unit, new_unit)
                if orig_rev:
                    _apply_rev_map(orig_rev, new_rev)
            except Exception as e:
                self.log_line(f"Error applying global mapping: {e}")
            self._save_overrides()
            self.log_line(f"Override saved for {filename}: unit={new_unit}, rev={new_rev}")
            dlg.destroy()

        def on_cancel():
            dlg.destroy()

        ttk.Button(btn_frame, text="OK", command=on_ok).pack(side="right", padx=(6, 0))
        ttk.Button(btn_frame, text="Cancel", command=on_cancel).pack(side="right")

    # ------------------------------------------------------------
    # Scan exports (uses self.exports_dir.get())
    # ------------------------------------------------------------
    def scan_exports(self):
        def worker():
            path = self.exports_dir.get() or _browse_for_folder_blocking("Select Exports Folder")
            self.after(0, lambda: self._on_scan_exports_done(path))
        threading.Thread(target=worker, daemon=True).start()

    def _on_scan_exports_done(self, path):
        if not path:
            self.log_line("Scan Exports cancelled")
            return
        self.log_line(f"Scanning exports in: {path} ...")
        threading.Thread(target=self._scan_exports_worker, args=(path,), daemon=True).start()

    # ------------------------------------------------------------
    # Scan (unchanged)
    # ------------------------------------------------------------
    def scan(self):
        self.tree.delete(*self.tree.get_children())

        rss_dir = self.rss_dir.get()
        if not os.path.isdir(rss_dir):
            messagebox.showerror("Invalid Folder", rss_dir)
            return

        files = [f for f in os.listdir(rss_dir) if f.lower().endswith(".rss")]

        if self.auto_delete_bak.get():
            for f in list(files):
                if "_bak" in f.lower():
                    try:
                        os.remove(os.path.join(rss_dir, f))
                        self.log_line(f"Deleted backup file: {f}")
                        files.remove(f)
                    except Exception as e:
                        self.log_line(f"ERROR deleting {f}: {e}")

        for f in sorted(files):
            unit, rev = parse_unit_and_revision_from_filename(f)

            fo = self.file_overrides.get(f, {})
            if fo:
                unit = fo.get("unit", unit)
                rev = fo.get("revision", rev)
                status = fo.get("status", "NEW")
            else:
                if unit in self.unit_map:
                    unit = self.unit_map[unit]
                if rev in self.rev_map:
                    rev = self.rev_map[rev]
                status = "NEW"

            if unit and rev:
                self.tree.insert("", "end", values=(unit, rev, status, f))

        self.log_line("Scan complete")

    # ------------------------------------------------------------
    # Pause / Cancel behavior (unchanged)
    # ------------------------------------------------------------
    def _toggle_pause(self):
        if self._resume_event.is_set():
            self._resume_event.clear()
            self.pause_btn.config(text="Resume")
            self.log_line("Processing paused")
        else:
            self.pause_btn.config(state="disabled")
            self.log_line("Resuming in 5 seconds...")
            def resume_after():
                time.sleep(5)
                self._resume_event.set()
                self.after(0, lambda: self.pause_btn.config(text="Pause", state="normal"))
                self.log_line("Processing resumed")
            threading.Thread(target=resume_after, daemon=True).start()

    def _on_cancel_clicked(self):
        self._resume_event.clear()
        self.pause_btn.config(text="Resume")
        dlg = tk.Toplevel(self)
        dlg.title("Cancel / Restart")
        dlg.transient(self)
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.geometry("+%d+%d" % (self.winfo_rootx() + 150, self.winfo_rooty() + 150))

        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Processing paused. Choose an action:").pack(anchor="w", pady=(0,8))
        ttk.Label(frm, text="Cancel current program and restart it, or stop remaining items?").pack(anchor="w")

        btn_frame = ttk.Frame(frm)
        btn_frame.pack(fill="x", pady=(12,0))

        def do_cancel_restart():
            self._restart_current_request.set()
            try:
                if hasattr(open_program, "cancel_current"):
                    threading.Thread(target=open_program.cancel_current, daemon=True).start()
                    self.log_line("Requested cancel of current program (via open_program.cancel_current())")
                else:
                    self.log_line("No cancel_current() hook available in open_program; will restart after current finishes")
            except Exception as e:
                self.log_line(f"Error requesting cancel: {e}")
            dlg.destroy()
            self._delayed_resume_with_log(5)

        def do_stop_all():
            self._stop_requested.set()
            self.log_line("Stop requested — remaining files will not be processed")
            dlg.destroy()

        def do_resume_only():
            dlg.destroy()
            self._delayed_resume_with_log(5)

        ttk.Button(btn_frame, text="Cancel & Restart Current", command=do_cancel_restart).pack(side="left", padx=6)
        ttk.Button(btn_frame, text="Stop All", command=do_stop_all).pack(side="left", padx=6)
        ttk.Button(btn_frame, text="Resume", command=do_resume_only).pack(side="right", padx=6)

    # ------------------------------------------------------------
    # Helpers for PDF detection & cancellation (unchanged)
    # ------------------------------------------------------------
    def _delayed_resume_with_log(self, seconds: int):
        self.pause_btn.config(state="disabled")
        def resume_after():
            time.sleep(seconds)
            self._resume_event.set()
            self.after(0, lambda: self.pause_btn.config(text="Pause", state="normal"))
            self.log_line("Processing resumed")
        threading.Thread(target=resume_after, daemon=True).start()

    def _list_pdfs(self, reports_dir: str):
        out = set()
        try:
            for root, _, files in os.walk(reports_dir):
                for fn in files:
                    if fn.lower().endswith(".pdf"):
                        rel = os.path.relpath(os.path.join(root, fn), reports_dir)
                        out.add(rel.replace("\\", "/"))
        except Exception:
            pass
        return out

    def _prompt_missing_pdf(self, filename: str):
        self._missing_pdf_event = threading.Event()
        self._missing_pdf_action = None

        def show_dialog():
            dlg = tk.Toplevel(self)
            dlg.title("Missing PDF")
            dlg.transient(self)
            dlg.resizable(False, False)
            dlg.grab_set()
            dlg.geometry("+%d+%d" % (self.winfo_rootx() + 160, self.winfo_rooty() + 160))

            frm = ttk.Frame(dlg, padding=12)
            frm.pack(fill="both", expand=True)

            ttk.Label(frm, text=f"No PDF was produced for: {filename}").pack(anchor="w", pady=(0,8))
            ttk.Label(frm, text="Choose an action: Continue (skip) or Cancel Program (attempt to cancel/close RSLogix)").pack(anchor="w")

            def do_continue():
                self._missing_pdf_action = "continue"
                dlg.destroy()
                self._missing_pdf_event.set()

            def do_cancel_program():
                self._missing_pdf_action = "cancel_program"
                dlg.destroy()
                self._missing_pdf_event.set()

            btns = ttk.Frame(frm)
            btns.pack(fill="x", pady=(12,0))
            ttk.Button(btns, text="Continue", command=do_continue).pack(side="right", padx=6)
            ttk.Button(btns, text="Cancel Program", command=do_cancel_program).pack(side="right", padx=6)

            dlg.wait_window(dlg)

        self.after(0, show_dialog)
        self._missing_pdf_event.wait()
        return self._missing_pdf_action

    def _attempt_cancel_rslogix(self):
        try:
            if hasattr(open_program, "cancel_current"):
                try:
                    open_program.cancel_current()
                    self.log_line("Requested cancel of current program (open_program.cancel_current())")
                except Exception as e:
                    self.log_line(f"open_program.cancel_current() raised: {e}")
            else:
                self.log_line("open_program.cancel_current() not available")
            if hasattr(open_program, "close_rslogix"):
                try:
                    open_program.close_rslogix()
                    self.log_line("Requested RSLogix close (open_program.close_rslogix())")
                except Exception as e:
                    self.log_line(f"open_program.close_rslogix() raised: {e}")
            else:
                self.log_line("open_program.close_rslogix() not available")
        except Exception as e:
            self.log_line(f"Error attempting RSLogix cancellation/close: {e}")

    # ------------------------------------------------------------
    # Processing with progress rows
    # ------------------------------------------------------------
    def process_selected(self):
        self._run_items(self.tree.selection())

    def process_new_only(self):
        items = [i for i in self.tree.get_children() if self.tree.item(i)["values"][2] == "NEW"]
        self._run_items(items)

    def _mark_file_processed(self, filename: str):
        now = datetime.now().isoformat()
        fo = self.file_overrides.get(filename, {})
        fo.update({"status": "PROCESSED", "processed_at": now})
        self.file_overrides[filename] = fo
        self._save_overrides()

    def _set_programs_progress(self, cur, total):
        try:
            self.prog_bar["maximum"] = int(total) if total and int(total) > 0 else 1
            self.prog_bar["value"] = int(cur)
            self.prog_label.config(text=f"Programs: {cur} of {total}")
        except Exception:
            pass

    def _set_pdf_progress(self, cur, total):
        try:
            self.pdf_bar["maximum"] = int(total) if total and int(total) > 0 else 1
            self.pdf_bar["value"] = int(cur)
            self.pdf_label.config(text=f"PDF pages: {cur} of {total}")
        except Exception:
            pass

    def _set_ladder_progress(self, cur, total):
        try:
            self.ladder_bar["maximum"] = int(total) if total and int(total) > 0 else 1
            self.ladder_bar["value"] = int(cur)
            self.ladder_label.config(text=f"Ladders: {cur} of {total}")
        except Exception:
            pass

    def _run_items(self, items):
        items = list(items)
        if not items:
            self.log_line("No items to process")
            return

        # Prevent concurrent processing runs
        with self._processing_lock:
            if self._processing_thread and self._processing_thread.is_alive():
                self.log_line("Processing already running — new request ignored")
                return

        def worker():
            # Ensure lazy imports are available
            _lazy_imports()
            if process_rss_file is None or open_program is None or print_report is None:
                self.log_line("Required modules failed to import — aborting processing. Check logs.")
                return

            try:
                logger.set_debug_sink(self.log_line)
                try:
                    open_program.set_debug_sink(self.log_line)
                except Exception:
                    pass
                try:
                    print_report.set_debug_sink(self.log_line)
                except Exception:
                    pass
                self.log_line("Debug sinks enabled")

                try:
                    self.process_selected_btn.config(state="disabled")
                    self.process_new_btn.config(state="disabled")
                except Exception:
                    pass

                reports_dir = self.reports_dir.get() or r"C:\PLC_Agent\reports"
                os.makedirs(reports_dir, exist_ok=True)

                # initialize programs progress
                total_programs = len(items)
                self.after(0, lambda: self._set_programs_progress(0, total_programs))

                for prog_index, item in enumerate(items, start=1):
                    if self._stop_requested.is_set():
                        self.log_line("Processing stopped by user request")
                        break

                    # update programs progress at start of each item
                    self.after(0, lambda cur=prog_index, tot=total_programs: self._set_programs_progress(cur, tot))

                    if not self._resume_event.is_set():
                        self.log_line("Processing paused (before starting item)...")
                        self._resume_event.wait()
                        self.log_line("Processing resumed")

                    self.log_line("Waiting for resume (if paused)...")
                    self._resume_event.wait()

                    # before PDFs snapshot (not used for final decision)
                    before_pdfs = self._list_pdfs(reports_dir)

                    self._current_item = item
                    vals = self.tree.item(item)["values"]
                    if not vals or len(vals) < 4:
                        self._current_item = None
                        continue
                    unit, rev, _, fname = vals
                    self._current_fname = fname
                    full = os.path.join(self.rss_dir.get(), fname)

                    try:
                        self.log_line(f"Processing {fname} (unit={unit}, rev={rev})")
                        if not os.path.isfile(full):
                            self.log_line(f"ERROR {fname}: file not found at {full}")
                            continue

                        proc_result = {"exception": None}

                        def run_processing():
                            self.log_line(f"proc_thread START for {fname}")
                            try:
                                # PDF progress callback -> update pdf widgets
                                def _pdf_progress(cur, total):
                                    try:
                                        self.after(0, lambda c=cur, t=total: self._set_pdf_progress(c, t))
                                        self.after(0, lambda: self.tree.set(item, "status", f"EXTRACT {cur}/{total}"))
                                    except Exception:
                                        pass

                                # Ladder progress callback -> update ladder widgets
                                def _ladder_progress(cur, total):
                                    try:
                                        self.after(0, lambda c=cur, t=total: self._set_ladder_progress(c, t))
                                    except Exception:
                                        pass

                                process_rss_file(
                                    full,
                                    resume_event=self._resume_event,
                                    stop_event=self._stop_requested,
                                    restart_event=self._restart_current_request,
                                    progress_callback=_pdf_progress,
                                    ladder_progress_callback=_ladder_progress,
                                )
                                self.log_line(f"proc_thread FINISH for {fname}")
                            except Exception as ex:
                                tb = traceback.format_exc()
                                self.log_line(f"proc_thread EXCEPTION for {fname}: {ex}\n{tb}")
                                proc_result["exception"] = ex

                        proc_thread = threading.Thread(target=run_processing, daemon=True)
                        proc_thread.start()

                        # Wait for processing thread to finish, honoring pause/stop; cap wait to avoid infinite hang
                        max_processing_wait = 600.0  # overall safety cap per item (seconds)
                        proc_poll = 0.5
                        waited = 0.0
                        while proc_thread.is_alive() and waited < max_processing_wait:
                            if not self._resume_event.is_set():
                                self.log_line("Processing paused (waiting for program to finish)...")
                                self._resume_event.wait()
                                self.log_line("Processing resumed (waiting for program)")
                            if self._stop_requested.is_set():
                                break
                            time.sleep(proc_poll)
                            waited += proc_poll

                        if proc_thread.is_alive():
                            self.log_line(f"Processing thread for {fname} did not finish in time")
                            try:
                                if hasattr(open_program, "cancel_current"):
                                    open_program.cancel_current()
                                    self.log_line("Requested cancel of current program due to timeout")
                            except Exception as e:
                                self.log_line(f"Error requesting cancel: {e}")
                            proc_thread.join(timeout=5.0)

                        # Finalize status based on result
                        if proc_result.get("exception"):
                            self.log_line(f"Processing {fname} raised exception: {proc_result['exception']}")
                            action = self._prompt_missing_pdf(fname)
                            if action == "continue":
                                fo = self.file_overrides.get(fname, {})
                                fo.update({"status": "ERROR_NO_PDF", "processed_at": datetime.now().isoformat()})
                                self.file_overrides[fname] = fo
                                self._save_overrides()
                                self.tree.set(item, "status", "ERROR_NO_PDF")
                                self.log_line(f"Marked {fname} as ERROR_NO_PDF and continuing")
                            elif action == "cancel_program":
                                self._attempt_cancel_rslogix()
                                fo = self.file_overrides.get(fname, {})
                                fo.update({"status": "CANCELLED", "processed_at": datetime.now().isoformat()})
                                self.file_overrides[fname] = fo
                                self._save_overrides()
                                self.tree.set(item, "status", "CANCELLED")
                                self.log_line(f"Cancelled program for {fname}; continuing to next item")
                            else:
                                self.log_line("Unknown action from missing-PDF dialog; continuing")
                        else:
                            self.log_line(f"Processing complete for {fname} (no exception) — marking PROCESSED")
                            self.tree.set(item, "status", "PROCESSED")
                            self._mark_file_processed(fname)

                        # clear progress UI after item finished
                        try:
                            self.after(0, lambda: (self.pdf_label.config(text=""), self.pdf_bar.configure(value=0, maximum=1),
                                                  self.ladder_label.config(text=""), self.ladder_bar.configure(value=0, maximum=1)))
                        except Exception:
                            pass

                        # clear current
                        self._current_item = None
                        self._current_fname = None

                    except Exception as e:
                        self.log_line(f"ERROR {fname}: {e}")
                        self._current_item = None
                        self._current_fname = None

                    if self._stop_requested.is_set():
                        self.log_line("Processing stopped by user request (post-item)")
                        break

                # cleanup
                self._current_item = None
                self._current_fname = None
                self._stop_requested.clear()
                self.log_line("Processing run complete")
            finally:
                # clear processing-thread guard and re-enable UI on main thread
                def _on_done():
                    try:
                        self.process_selected_btn.config(state="normal")
                        self.process_new_btn.config(state="normal")
                    except Exception:
                        pass
                    with self._processing_lock:
                        self._processing_thread = None
                self.after(0, _on_done)

        # create and register the thread, start it
        t = threading.Thread(target=worker, daemon=True)
        with self._processing_lock:
            if self._processing_thread and self._processing_thread.is_alive():
                self.log_line("Processing already running — new request ignored")
                return
            self._processing_thread = t

        self.log_line(f"Starting processing thread for {len(items)} item(s)")
        t.start()


if __name__ == "__main__":
    import traceback, sys, os
    log_dir = r"C:\PLC_Agent"
    os.makedirs(log_dir, exist_ok=True)
    startup_log = os.path.join(log_dir, "startup.log")

    def _append_startup(msg: str):
        try:
            with open(startup_log, "a", encoding="utf-8") as fh:
                fh.write(f"{datetime.now().isoformat()} {msg}\n")
                fh.flush()
        except Exception:
            pass

    _append_startup("ABOUT TO CREATE APP")
    try:
        app = PLCUI()
        _append_startup("APP CREATED")
        _append_startup("ENTERING MAINLOOP")
        try:
            app.mainloop()
            _append_startup("MAINLOOP EXITED (normal)")
        except Exception:
            tb = traceback.format_exc()
            _append_startup(f"MAINLOOP EXCEPTION: {tb}")
            raise
    except Exception:
        tb = traceback.format_exc()
        try:
            crash_path = os.path.join(log_dir, "plc-agent-crash.log")
            with open(crash_path, "w", encoding="utf-8") as fh:
                fh.write(tb)
        except Exception:
            pass
        _append_startup(f"APP CREATION FAILED: {tb}")
        raise