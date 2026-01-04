import os, json, traceback

def test_resolve_and_create():
    try:
        if os.name == "nt":
            base = os.getenv("LOCALAPPDATA") or os.path.expanduser("~")
            cfg_dir = os.path.join(base, "PLC-Agent")
        else:
            base = os.getenv("XDG_CONFIG_HOME") or os.path.expanduser("~/.config")
            cfg_dir = os.path.join(base, "plc-agent")
        print("Resolved cfg_dir:", repr(cfg_dir))
        # attempt to create dir
        os.makedirs(cfg_dir, exist_ok=True)
        if os.path.isdir(cfg_dir):
            print("Directory exists / created OK")
        else:
            print("Directory NOT present after os.makedirs")
        # attempt to write a test file
        test_path = os.path.join(cfg_dir, "overrides_test.json")
        with open(test_path, "w", encoding="utf-8") as fh:
            json.dump({"ok": True}, fh)
        print("Created test file:", test_path)
    except Exception as e:
        print("Exception:", e)
        traceback.print_exc()
        try:
            fallback = os.path.join(os.path.dirname(__file__), "overrides.json")
            print("Fallback path would be:", fallback)
        except Exception as e2:
            print("Fallback resolution failed:", e2)

if __name__ == "__main__":
    test_resolve_and_create()