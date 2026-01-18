# temp_ui.py
#
# Minimal UI for PlainLogic
# Scan revisions and open rung index

import logging
import threading
import json
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import likeaversion
from rung_index import build_rung_index_from_files
import rung_index_ui

SETTINGS_FILE = Path(__file__).parent / "plainlogic_settings.json"


def load_settings():
    if SETTINGS_FILE.exists():
        try:
            return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def save_settings(settings: dict):
    SETTINGS_FILE.write_text(json.dumps(settings, indent=2), encoding="utf-8")


def launch():
    logging.info("Temp UI launched")

    root = tk.Tk()
    root.title("PlainLogic – Scan & Index")
    root.geometry("820x420")
    root.resizable(False, False)

    root_dir = tk.StringVar()
    status_text = tk.StringVar(value="Idle")

    settings = load_settings()
    if "last_root_dir" in settings:
        root_dir.set(settings["last_root_dir"])

    def select_root():
        path = filedialog.askdirectory(title="Select Top-Level Directory")
        if path:
            root_dir.set(path)
            settings["last_root_dir"] = path
            save_settings(settings)

    def update_progress(cur, total, phase, detail):
        progress["maximum"] = total
        progress["value"] = cur
        status_text.set(f"{phase}: {detail} ({cur}/{total})")
        root.update_idletasks()

    def run_scan():
        if not root_dir.get():
            messagebox.showerror("Missing Directory", "Select a directory")
            return

        def task():
            try:
                status_text.set("Scanning...")
                likeaversion.run(Path(root_dir.get()), progress_cb=update_progress)
                status_text.set("Scan complete")
            except Exception as e:
                logging.exception("Scan failed")
                messagebox.showerror("Error", str(e))
                status_text.set("Scan failed")

        threading.Thread(target=task, daemon=True).start()

    def open_rung_index():
        if not root_dir.get():
            messagebox.showerror("Missing Directory", "Select a directory")
            return

        root_path = Path(root_dir.get())
        index_files = sorted(root_path.rglob("likeaversion.json"))

        if not index_files:
            messagebox.showwarning("No Index Files", "Run Scan first.")
            return

        rung_index = build_rung_index_from_files(index_files)
        logging.info("Alignment index built: %d families", len(rung_index))
        rung_index_ui.launch(root, rung_index)

    top = ttk.Frame(root)
    top.pack(fill="x", padx=12, pady=12)

    ttk.Label(top, text="Top-Level Directory:").pack(side="left")
    ttk.Entry(top, textvariable=root_dir, width=56).pack(side="left", padx=6)
    ttk.Button(top, text="Browse", command=select_root).pack(side="left")

    ttk.Button(top, text="Scan", command=run_scan).pack(side="left", padx=6)
    ttk.Button(top, text="Rung Index", command=open_rung_index).pack(side="left")

    progress = ttk.Progressbar(root, length=780)
    progress.pack(padx=12, pady=10)

    ttk.Label(root, textvariable=status_text).pack()

    root.mainloop()
