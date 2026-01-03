import os
import threading
import ctypes
from ctypes import wintypes
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

from agent import process_rss_file
from util.common import parse_unit_and_revision_from_filename

import python.open_program as open_program
import python.print_report as print_report

# Add: import unified logger so we can route all module logs into the UI
import util.logger as logger

import python.hide_console  # module executes and hides console on import


# ------------------------------------------------------------------
# Native Windows Folder Picker (STA COM)
# ------------------------------------------------------------------

_BIF_RETURNONLYFSDIRS = 0x00000001
_BIF_NEWDIALOGSTYLE = 0x00000040

_shell32 = ctypes.WinDLL("shell32", use_last_error=True)
_ole32 = ctypes.WinDLL("ole32", use_last_error=True)

LPITEMIDLIST = ctypes.c_void_p
HRESULT = ctypes.c_long  # ✅ FIX: define HRESULT explicitly


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


def _browse_for_folder_blocking(title="Select RSS Folder"):
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

    # Cast the output buffer as LPWSTR when passing to SHGetPathFromIDListW
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

        self.rss_dir = tk.StringVar(value=r"C:\PLC_Agent\rss")
        self.auto_delete_bak = tk.BooleanVar(value=True)

        self._build_ui()

    # ------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------

    def _build_ui(self):
        root = ttk.Frame(self)
        root.pack(fill="both", expand=True)

        left = ttk.Frame(root)
        left.pack(side="left", fill="y", padx=8, pady=8)

        ttk.Label(left, text="RSS Directory:").pack(anchor="w")
        dir_row = ttk.Frame(left)
        dir_row.pack(fill="x", pady=(0, 8))
        ttk.Entry(dir_row, textvariable=self.rss_dir, width=40).pack(side="left", fill="x", expand=True)
        ttk.Button(dir_row, text="Browse...", command=self.browse).pack(side="left", padx=(6, 0))

        ttk.Checkbutton(
            left,
            text="Auto-delete _BAK files on scan",
            variable=self.auto_delete_bak
        ).pack(anchor="w", pady=(0, 8))

        actions = ttk.Frame(left)
        actions.pack(fill="x", pady=(12, 8))
        ttk.Button(actions, text="Scan", command=self.scan).pack(side="left")
        ttk.Button(actions, text="Process Selected", command=self.process_selected).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Process NEW Only", command=self.process_new_only).pack(side="left", padx=(8, 0))

        right = ttk.Frame(root)
        right.pack(side="left", fill="both", expand=True, padx=8, pady=8)

        tree_frame = ttk.LabelFrame(right, text="RSS Files")
        tree_frame.pack(fill="both", expand=True)

        # NOTE: add a hidden 'filename' column so we store the original filename
        columns = ("unit", "revision", "status", "filename")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="extended")
        # Only display the human columns; keep filename available in values
        self.tree["displaycolumns"] = ("unit", "revision", "status")

        for col, text, w in [
            ("unit", "Unit", 240),
            ("revision", "Revision", 120),
            ("status", "Status", 120),
            ("filename", "Filename", 0),
        ]:
            self.tree.heading(col, text=text)
            # make filename non-stretching and effectively hidden
            if col == "filename":
                self.tree.column(col, width=0, stretch=False)
            else:
                self.tree.column(col, width=w, anchor="w")

        yscroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=yscroll.set)

        self.tree.pack(side="left", fill="both", expand=True)
        yscroll.pack(side="right", fill="y")

        log_frame = ttk.LabelFrame(right, text="Log")
        log_frame.pack(fill="both", pady=(8, 0))

        self.log = tk.Text(log_frame, height=12, state="disabled", wrap="word")
        self.log.pack(fill="both", expand=True)

    # ------------------------------------------------------------
    # Logging (thread-safe)
    # ------------------------------------------------------------

    def log_line(self, msg: str):
        stamp = datetime.now().strftime("[%H:%M:%S] ")
        line = stamp + msg

        def append():
            self.log.configure(state="normal")
            self.log.insert("end", line + "\n")
            self.log.see("end")
            self.log.configure(state="disabled")

        self.after(0, append)

    # ------------------------------------------------------------
    # Browse
    # ------------------------------------------------------------

    def browse(self):
        def worker():
            path = _browse_for_folder_blocking("Select RSS Folder")
            self.after(0, lambda: self._on_browse_done(path))

        threading.Thread(target=worker, daemon=True).start()

    def _on_browse_done(self, path):
        if path:
            self.rss_dir.set(path)
            self.log_line(f"Selected directory: {path}")

    # ------------------------------------------------------------
    # Scan
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
            if unit and rev:
                # store original filename in the hidden 'filename' value
                self.tree.insert("", "end", values=(unit, rev, "NEW", f))

        self.log_line("Scan complete")

    # ------------------------------------------------------------
    # Processing
    # ------------------------------------------------------------

    def process_selected(self):
        self._run_items(self.tree.selection())

    def process_new_only(self):
        items = [
            i for i in self.tree.get_children()
            if self.tree.item(i)["values"][2] == "NEW"
        ]
        self._run_items(items)

    def _run_items(self, items):
        def worker():
            # Route all module logs to the UI text widget
            logger.set_debug_sink(self.log_line)

            open_program.set_debug_sink(self.log_line)
            print_report.set_debug_sink(self.log_line)
            self.log_line("Debug sinks enabled")

            for item in items:
                vals = self.tree.item(item)["values"]
                unit, rev, _, fname = vals
                full = os.path.join(self.rss_dir.get(), fname)

                try:
                    self.log_line(f"Processing {fname}")
                    if not os.path.isfile(full):
                        self.log_line(f"ERROR {fname}: file not found at {full}")
                        continue
                    process_rss_file(full)
                    self.tree.set(item, "status", "PROCESSED")
                except Exception as e:
                    self.log_line(f"ERROR {fname}: {e}")

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    app = PLCUI()
    app.mainloop()
