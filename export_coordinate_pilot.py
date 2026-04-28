from pywinauto import Application
from pathlib import Path
import pyautogui
import win32com.client
import pythoncom
import time
import hashlib
import csv
import re
import traceback

OUT_DIR = Path(r"C:\EladalahExport\docs")
LOG_FILE = Path(r"C:\EladalahExport\logs\coordinate_pilot_log.csv")

OUT_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE.parent.mkdir(parents=True, exist_ok=True)

# These are relative to the Eladalah main window.
# Based on your controls dump:
EXPORT_X_REL = 458
EXPORT_Y_REL = 98

TREE_X_REL = 1540
FIRST_ITEM_Y_REL = 250
TREE_ROW_STEP = 28

MAX_ITEMS = 5

INVALID = r'[\\/:*?"<>|]'

def clean_filename(text, max_len=80):
    text = re.sub(INVALID, " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return (text or "eladalah-export")[:max_len]

def short_hash(text):
    return hashlib.md5(text.encode("utf-8", errors="ignore")).hexdigest()[:8]

def get_eladalah_rect():
    app = Application(backend="uia").connect(path="eladalah2025.exe", timeout=10)
    windows = app.windows()

    candidates = []
    for w in windows:
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
    print(f"Rect: {rect}")
    w.set_focus()
    time.sleep(1)
    return rect

def get_word():
    try:
        return win32com.client.GetActiveObject("Word.Application")
    except Exception:
        return win32com.client.Dispatch("Word.Application")

def close_all_word_docs():
    try:
        word = get_word()
        while word.Documents.Count > 0:
            word.ActiveDocument.Close(False)
    except Exception:
        pass

def save_active_word_doc(path):
    word = get_word()
    word.Visible = True

    doc = None
    for _ in range(60):
        try:
            if word.Documents.Count > 0:
                doc = word.ActiveDocument
                break
        except Exception:
            pass
        time.sleep(0.5)

    if doc is None:
        raise RuntimeError("No Word document appeared after export.")

    doc.SaveAs2(str(path), FileFormat=16)  # docx
    doc.Close(False)

def main():
    pythoncom.CoInitialize()
    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 0.3

    close_all_word_docs()

    rect = get_eladalah_rect()

    export_x = rect.left + EXPORT_X_REL
    export_y = rect.top + EXPORT_Y_REL
    tree_x = rect.left + TREE_X_REL

    print(f"Export button click: {export_x}, {export_y}")
    print(f"Tree click X: {tree_x}")

    with LOG_FILE.open("a", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        if LOG_FILE.stat().st_size == 0:
            writer.writerow(["status", "index", "click_x", "click_y", "output_file", "error"])

        for i in range(MAX_ITEMS):
            item_y = rect.top + FIRST_ITEM_Y_REL + (i * TREE_ROW_STEP)
            filename = f"pilot-{i+1:03d}-{short_hash(str(item_y))}.docx"
            out_path = OUT_DIR / filename

            print(f"\nExporting visible row {i+1}")
            print(f"Click tree: {tree_x}, {item_y}")

            try:
                # Focus/click tree item
                pyautogui.click(tree_x, item_y)
                time.sleep(1.5)

                # Click Convert to Word
                pyautogui.click(export_x, export_y)
                time.sleep(3)

                # Save Word document
                save_active_word_doc(out_path)

                print(f"Saved: {out_path}")
                writer.writerow(["OK", i + 1, tree_x, item_y, str(out_path), ""])

            except Exception as e:
                print("FAILED")
                print(traceback.format_exc())
                writer.writerow(["FAILED", i + 1, tree_x, item_y, str(out_path), str(e)])

            f.flush()
            time.sleep(1)

    print("\nCoordinate pilot finished.")
    print(f"Check docs here: {OUT_DIR}")
    print(f"Log here: {LOG_FILE}")

if __name__ == "__main__":
    main()