from pywinauto import Application, Desktop
from pathlib import Path
import traceback

OUT = Path(r"C:\EladalahExport\logs\controls.txt")

def safe(x):
    try:
        return str(x)
    except Exception:
        return ""

def dump_window_controls(window, f, backend_name):
    f.write(f"\n\n=== {backend_name} WINDOW ===\n")
    f.write(f"Title: {safe(window.window_text())}\n")
    f.write(f"Class: {safe(window.class_name())}\n")
    f.write(f"Rect: {safe(window.rectangle())}\n")

    try:
        controls = window.descendants()
        f.write(f"\nTotal descendants: {len(controls)}\n\n")

        for i, c in enumerate(controls):
            try:
                info = c.element_info

                name = safe(getattr(info, "name", ""))
                control_type = safe(getattr(info, "control_type", ""))
                class_name = safe(getattr(info, "class_name", ""))
                auto_id = safe(getattr(info, "automation_id", ""))
                rect = safe(c.rectangle())

                f.write(
                    f"[{i}] "
                    f"name='{name}' | "
                    f"type='{control_type}' | "
                    f"class='{class_name}' | "
                    f"auto_id='{auto_id}' | "
                    f"rect={rect}\n"
                )
            except Exception as e:
                f.write(f"[{i}] ERROR reading control: {e}\n")

    except Exception:
        f.write("\nFAILED TO READ DESCENDANTS:\n")
        f.write(traceback.format_exc())


with OUT.open("w", encoding="utf-8") as f:
    f.write("=== DESKTOP WINDOWS ===\n")

    for w in Desktop(backend="uia").windows():
        title = safe(w.window_text())
        if title.strip():
            f.write(title + "\n")

    f.write("\n\n=== CONNECT UIA ===\n")
    try:
        app = Application(backend="uia").connect(path="eladalah2025.exe", timeout=10)
        for w in app.windows():
            dump_window_controls(w, f, "UIA")
    except Exception:
        f.write(traceback.format_exc())

    f.write("\n\n=== CONNECT WIN32 ===\n")
    try:
        app = Application(backend="win32").connect(path="eladalah2025.exe", timeout=10)
        for w in app.windows():
            dump_window_controls(w, f, "WIN32")
    except Exception:
        f.write(traceback.format_exc())

print(f"Done. Saved to: {OUT}")