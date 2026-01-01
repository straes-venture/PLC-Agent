import os
import re
import time
from pywinauto import Desktop
from pywinauto.keyboard import send_keys


SETTLE = 0.25
PDF_PRINTER = "Microsoft Print to PDF"

# ------------------------------------------------------------
# Debug sink plumbing
# ------------------------------------------------------------

_DEBUG_SINK = None

def set_debug_sink(fn):
    global _DEBUG_SINK
    _DEBUG_SINK = fn

def _dbg(msg: str):
    if _DEBUG_SINK:
        try:
            _DEBUG_SINK(msg)
            return
        except Exception:
            pass
    print(msg, flush=True)

def _t_dbg(msg: str):
    _dbg(f"[TEMPLATE] {msg}")

def _key(seq: str, label: str = ""):
    if label:
        _dbg(f"[KEY] {label}: {seq}")
    else:
        _dbg(f"[KEY] {seq}")
    send_keys(seq)

def _type(win, seq: str, label: str = ""):
    if label:
        _dbg(f"[KEY] {label}: {seq}")
    else:
        _dbg(f"[KEY] {seq}")
    win.type_keys(seq)

# ------------------------------------------------------------
# Template path resolver (RO0 vs R00 defense)
# ------------------------------------------------------------

def resolve_template_path(template_path: str) -> str:
    template_path = os.path.abspath(template_path)
    if os.path.exists(template_path):
        return template_path

    base_dir = os.path.dirname(template_path)
    base_name = os.path.splitext(os.path.basename(template_path))[0]

    pat = re.compile(rf"^{re.escape(base_name)}\.r[0o][0o]$", re.I)
    matches = [
        os.path.join(base_dir, f)
        for f in os.listdir(base_dir)
        if pat.match(f)
    ]

    if len(matches) == 1:
        _dbg(f"[TEMPLATE] WARNING: Using '{os.path.basename(matches[0])}'")
        return matches[0]

    if not matches:
        raise FileNotFoundError(f"Template not found: {template_path}")

    raise RuntimeError(f"Multiple template candidates: {matches}")

# ------------------------------------------------------------
# Dialog helpers
# ------------------------------------------------------------

def _find_dialog(title_re, timeout=20):
    end = time.time() + timeout
    while time.time() < end:
        dlg = Desktop(backend="win32").window(
            class_name="#32770",
            title_re=title_re
        )
        if dlg.exists(timeout=0.2):
            return dlg
        time.sleep(0.2)
    raise RuntimeError(f"Dialog not found: {title_re}")

def _find_print_dialog(timeout=15):
    return _find_dialog(r"^Print$", timeout)

def _find_file_dialog_open_or_save(timeout=20):
    return _find_dialog(r"(Open|Save|Save Print Output As)", timeout)

def _set_file_dialog_filename(dlg, full_path: str):
    dlg.set_focus()
    time.sleep(0.15)
    _type(dlg, "^a{BACKSPACE}", "Clear filename")
    _type(dlg, full_path, "Type filename")

# ------------------------------------------------------------
# File stability
# ------------------------------------------------------------

def _wait_for_file_stable(path: str, timeout=180, stable_secs=2.0):
    end = time.time() + timeout
    last_size = -1
    last_change = time.time()

    while time.time() < end:
        if os.path.exists(path):
            size = os.path.getsize(path)
            if size != last_size:
                last_size = size
                last_change = time.time()
            elif size > 0 and (time.time() - last_change) >= stable_secs:
                return True
        time.sleep(0.25)

    raise RuntimeError(f"PDF did not stabilize: {path}")

# ------------------------------------------------------------
# Apply report template (VERIFIED SEQUENCE)
# ------------------------------------------------------------

def apply_report_template(main_win, template_path: str):
    template_path = resolve_template_path(template_path)
    _t_dbg(f"Applying report template: {template_path}")

    main_win.set_focus()
    time.sleep(0.2)

    _t_dbg("File → Report Options")
    _key("%f", "Alt+F")
    time.sleep(0.15)
    _key("t", "Report Options")
    time.sleep(0.8)

    rpt_opts = _find_dialog(r"^Report Options$")
    rpt_opts.set_focus()
    time.sleep(0.2)

    _t_dbg("Open Load/Save")
    _key("+{TAB}" * 6, "Shift+Tab x6")
    _key("{ENTER}", "Enter")
    time.sleep(0.6)

    load_save = _find_dialog(r"^Report Options Setup Load/Save$")
    load_save.set_focus()
    time.sleep(0.2)

    _t_dbg("Import template")
    _key("+{TAB}", "Shift+Tab")
    _key("{ENTER}", "Enter")
    time.sleep(0.6)

    fd = _find_file_dialog_open_or_save(timeout=20)
    fd.set_focus()
    time.sleep(0.2)

    _t_dbg("Typing template path")
    _set_file_dialog_filename(fd, template_path)
    _key("{ENTER}", "Confirm template")
    time.sleep(0.8)

    load_save.wait("ready", timeout=10)
    load_save.set_focus()
    time.sleep(0.2)

    _t_dbg("Select template entry")
    _key("{TAB}" * 2, "Tab x2")
    _key("{SPACE}", "Space")
    time.sleep(0.25)

    _t_dbg("Load template")
    _key("{TAB}" * 2, "Tab x2")
    _key("{ENTER}", "Enter")
    time.sleep(0.3)
    _key("{ENTER}", "Confirm")
    time.sleep(0.3)

    _t_dbg("Close Load/Save")
    _key("+{TAB}", "Shift+Tab")
    _key("{ENTER}", "Enter")
    time.sleep(0.4)

    rpt_opts.wait("ready", timeout=10)
    rpt_opts.set_focus()
    time.sleep(0.2)

    _t_dbg("Close Report Options")
    _key("{TAB}" * 3, "Tab x3")
    _key("{ENTER}", "Enter")
    time.sleep(0.4)

# ------------------------------------------------------------
# Print sequence (VERIFIED)
# ------------------------------------------------------------

def run(app, main_win, pdf_out_path: str, template_path: str) -> str:
    pdf_out_path = os.path.abspath(pdf_out_path)
    os.makedirs(os.path.dirname(pdf_out_path), exist_ok=True)

    _dbg(f"[PRINT] Target PDF: {pdf_out_path}")

    if os.path.exists(pdf_out_path):
        os.remove(pdf_out_path)

    apply_report_template(main_win, template_path)

    main_win.set_focus()
    time.sleep(0.2)

    _dbg("[PRINT] Ctrl+R (Print)")
    _key("^r", "Ctrl+R")
    time.sleep(0.6)

    print_dlg = _find_print_dialog(timeout=15)
    print_dlg.set_focus()
    time.sleep(0.2)

    _dbg("[PRINT] Confirm Print dialog")
    _key("{ENTER}", "Enter")
    time.sleep(0.6)

    save_dlg = _find_file_dialog_open_or_save(timeout=25)
    save_dlg.set_focus()
    time.sleep(0.2)

    _dbg("[PRINT] Saving PDF")
    _set_file_dialog_filename(save_dlg, pdf_out_path)
    _key("{ENTER}", "Enter")

    _dbg("[PRINT] Waiting for PDF to stabilize")
    _wait_for_file_stable(pdf_out_path)

    _dbg("[PRINT] Report print complete")
    return pdf_out_path
