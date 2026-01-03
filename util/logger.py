"""
Simple unified logger used by the PLC Agent modules.

Provides:
- set_debug_sink(fn): route log lines into a UI (callable that accepts a single str)
- dbg(msg): emit a debug line (forwards to sink or stdout)

This mirrors the small logger API used across the codebase (modules call `logger.dbg(...)`
and the UI calls `logger.set_debug_sink(...)`).
"""
from typing import Callable, Optional

_debug_sink: Optional[Callable[[str], None]] = None


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
            return
        except Exception:
            # Fail-safe to ensure logging never raises
            pass
    try:
        print(msg, flush=True)
    except Exception:
        # Last resort: ignore any printing errors
        pass