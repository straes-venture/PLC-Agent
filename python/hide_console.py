# hide_console.py
# Hide the console window on Windows if the process has one. Safe no-op on other platforms.

import sys

if sys.platform == "win32":
    try:
        import ctypes

        GetConsoleWindow = ctypes.windll.kernel32.GetConsoleWindow
        ShowWindow = ctypes.windll.user32.ShowWindow

        whnd = GetConsoleWindow()
        if whnd:
            SW_HIDE = 0
            ShowWindow(whnd, SW_HIDE)
            # Optionally detach the console entirely:
            # ctypes.windll.kernel32.FreeConsole()
    except Exception:
        # Never crash the app due to console-hiding
        pass