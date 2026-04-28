from pywinauto import Application
from pathlib import Path
import win32com.client
import pythoncom
import time
import re
import hashlib
import csv
import traceback

OUT_DIR = Path(r"C:\EladalahExport\docs")
LOG_FILE = Path(r"C:\EladalahExport\logs\export_pilot_log.csv")

START_AT = 4
MAX_ITEMS = 5

OUT_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE.parent.mkdir(parents=True, exist_ok=True)

INVALID = r'[\\/:*?"<>|]'

def clean_filename(text, max_len=90):
    text = re.sub(INVALID, " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return (text or "untitled")[:max_len]

def short_hash(text):
    return hashlib.md5(text.encode("utf-8", errors="ignore")).hexdigest()[:8]

def info_value(control, attr):
    try:
        return getattr(control.element_info, attr, "") or ""
    except Exception:
        return ""

def control_name(control):
    try:
        return control.window_text() or info_value(control, "name")
    except Exception:
        return info_value(control, "name")

def get_word():
    try:
        return win32com.client.GetActiveObject("Word.Application")
    except Exception:
        return win32com.client.Dispatch("Word.Application")

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

    doc.SaveAs2(str(path), FileFormat=16)  # 16 = docx
    doc.Close(False)

def find_working_window(app):
    for w in app.windows():
        try:
            print(f"Checking window: {w.window_text()}")

            all_controls = w.descendants()

            tree_el = None
            export_btn_el = None

            for c in all_controls:
                auto_id = info_value(c, "automation_id")
                ctype = info_value(c, "control_type")
                name = control_name(c)

                if auto_id == "tv_Subjects" and ctype == "Tree":
                    tree_el = c

                if name.strip() == "تحويل لنص وورد" and ctype == "Button":
                    export_btn_el = c

            if tree_el and export_btn_el:
                print("FOUND working encyclopedia window.")
                return w, tree_el, export_btn_el

        except Exception as e:
            print(f"Skipped window due to error: {e}")

    raise RuntimeError(
        "Could not find tv_Subjects tree + تحويل لنص وورد button. "
        "Make sure an encyclopedia page is open, not the home screen."
    )

def get_visible_tree_items(tree):
    all_controls = tree.descendants()
    items = []

    for c in all_controls:
        ctype = info_value(c, "control_type")
        name = control_name(c).strip()

        if ctype == "TreeItem" and name:
            items.append(c)

    return items

def main():
    pythoncom.CoInitialize()

    app = Application(backend="uia").connect(path="eladalah2025.exe", timeout=10)

    main_win, tree, export_btn = find_working_window(app)

    main_win.set_focus()
    time.sleep(1)

    items = get_visible_tree_items(tree)

    print(f"Found {len(items)} visible tree items.")

    for i, item in enumerate(items[:25]):
        print(f"{i}: {control_name(item)}")

    selected_items = items[START_AT:START_AT + MAX_ITEMS]

    if not selected_items:
        raise RuntimeError("No selected tree items found. Try expanding the tree first.")

    with LOG_FILE.open("a", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        if LOG_FILE.stat().st_size == 0:
            writer.writerow(["status", "index", "title", "output_file", "error"])

        for idx, item in enumerate(selected_items, start=START_AT):
            title = control_name(item).strip()
            safe_title = clean_filename(title)
            filename = f"{idx:04d}-{safe_title}-{short_hash(title)}.docx"
            out_path = OUT_DIR / filename

            print(f"\n[{idx}] Exporting: {title}")

            try:
                main_win.set_focus()
                time.sleep(0.5)

                item.click_input()
                time.sleep(1.5)

                export_btn.click_input()
                time.sleep(3)

                save_active_word_doc(out_path)

                print(f"Saved: {out_path}")
                writer.writerow(["OK", idx, title, str(out_path), ""])

            except Exception as e:
                err = traceback.format_exc()
                print(f"FAILED: {title}")
                print(err)
                writer.writerow(["FAILED", idx, title, str(out_path), str(e)])

            f.flush()
            time.sleep(1)

    print("\nPilot finished.")
    print(f"Docs saved in: {OUT_DIR}")
    print(f"Log saved in: {LOG_FILE}")

if __name__ == "__main__":
    main()