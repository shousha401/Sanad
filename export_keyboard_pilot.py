from pywinauto import Application, Desktop
from pathlib import Path
import pyautogui
import pyperclip
import time
import hashlib
import csv
import traceback

OUT_DIR = Path(r"C:\EladalahExport\docs")
LOG_FILE = Path(r"C:\EladalahExport\logs\keyboard_pilot_log.csv")

OUT_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE.parent.mkdir(parents=True, exist_ok=True)

EXPORT_X_REL = 458
EXPORT_Y_REL = 98

TREE_X_REL = 1540
FIRST_ITEM_Y_REL = 278
TREE_ROW_STEP = 28

MAX_ITEMS = 5

def short_hash(text):
    return hashlib.md5(text.encode("utf-8", errors="ignore")).hexdigest()[:8]

def get_eladalah_rect():
    app = Application(backend="uia").connect(path="eladalah2025.exe", timeout=10)

    candidates = []
    for w in app.windows():
        try:
            title = w.window_text()
            rect = w.rectangle()
            area = (rect.right - rect.left) * (rect.bottom - rect.top)
            if "العدالة" in title or "2025" in title:
                candidates.append((area, w, rect, title))
        except Exception:
            pass

    if not candidates:
        raise RuntimeError("Could not find Eladalah window.")

    candidates.sort(reverse=True, key=lambda x: x[0])
    _, w, rect, title = candidates[0]

    print(f"Using window: {title}")
    w.set_focus()
    time.sleep(0.5)

    try:
        w.maximize()
        time.sleep(1)
    except Exception:
        pass

    rect = w.rectangle()
    print(f"Rect: {rect}")
    return rect

def wait_for_word_window():
    for _ in range(50):
        wins = Desktop(backend="uia").windows()
        for w in wins:
            title = w.window_text()
            if "Word" in title or "Document" in title or "Compatibility Mode" in title:
                try:
                    w.set_focus()
                    time.sleep(1)
                    print(f"Word window found: {title}")
                    return True
                except Exception:
                    pass
        time.sleep(0.5)

    return False

def save_word_with_keyboard(out_path):
    if not wait_for_word_window():
        raise RuntimeError("Word window did not appear.")

    time.sleep(1)

    # Open Save As dialog
    pyautogui.press("f12")
    time.sleep(2)

    # Paste full path using clipboard
    pyperclip.copy(str(out_path))
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.5)

    pyautogui.press("enter")
    time.sleep(3)

    # Confirm overwrite if popup appears
    pyautogui.press("enter")
    time.sleep(1)

    # Close Word
    pyautogui.hotkey("alt", "f4")
    time.sleep(1)

def main():
    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 0.25

    rect = get_eladalah_rect()

    export_x = rect.left + EXPORT_X_REL
    export_y = rect.top + EXPORT_Y_REL
    tree_x = rect.left + TREE_X_REL

    print(f"Export button click: {export_x}, {export_y}")
    print(f"Tree click X: {tree_x}")

    with LOG_FILE.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["status", "index", "click_x", "click_y", "output_file", "error"])

        for i in range(MAX_ITEMS):
            item_y = rect.top + FIRST_ITEM_Y_REL + (i * TREE_ROW_STEP)
            out_path = OUT_DIR / f"keyboard-pilot-{i+1:03d}-{short_hash(str(item_y))}.docx"

            print(f"\nExporting row {i+1}")
            print(f"Click tree: {tree_x}, {item_y}")
            print(f"Target save path: {out_path}")

            try:
                pyautogui.click(tree_x, item_y)
                time.sleep(1.5)

                pyautogui.click(export_x, export_y)
                time.sleep(3)

                save_word_with_keyboard(out_path)

                print(f"Saved: {out_path}")
                writer.writerow(["OK", i + 1, tree_x, item_y, str(out_path), ""])

            except Exception as e:
                print("FAILED")
                print(traceback.format_exc())
                writer.writerow(["FAILED", i + 1, tree_x, item_y, str(out_path), str(e)])

            f.flush()
            time.sleep(1)

    print("\nFinished.")
    print(f"Docs: {OUT_DIR}")
    print(f"Log: {LOG_FILE}")

if __name__ == "__main__":
    main()