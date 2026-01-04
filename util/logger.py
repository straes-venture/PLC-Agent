from typing import Callable, Optional
import os

"""
Simple unified logger used by the PLC Agent modules.

Provides:
- set_debug_sink(fn): route log lines into a UI (callable that accepts a single str)
- dbg(msg): emit a debug line (forwards to sink or stdout) and append to durable log file.
"""
_debug_sink: Optional[Callable[[str], None]] = None

LOG_PATH = r"C:\PLC_Agent\plc-agent.log"


def _append_to_file(msg: str) -> None:
    try:
        os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)
        with open(LOG_PATH, "a", encoding="utf-8") as fh:
            fh.write(msg + "\n")
    except Exception:
        # best-effort only; never raise
        pass


def set_debug_sink(fn: Optional[Callable[[str], None]]) -> None:
    """
    Set a callable to receive debug strings.

    Pass None to clear the sink and revert to printing to stdout.
    Example: set_debug_sink(ui.log_line)
    """
    global _debug_sink
    _debug_sink = fn


def dbg(msg: str) -> None:
    """
    Emit a debug message.

    If a debug sink is installed it will be called with the exact message.
    If the sink raises, or no sink is set, the message is printed to stdout.
    """
    if _debug_sink:
        try:
            _debug_sink(msg)
        except Exception:
            # Fail-safe; fall back to printing
            pass
    try:
        print(msg, flush=True)
    except Exception:
        pass
    # Always append a durable log copy
    _append_to_file(msg)