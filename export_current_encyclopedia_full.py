from pywinauto import Application
from pathlib import Path
from docx import Document
import pyautogui
import time
import hashlib
import csv
import traceback
import re

BASE_DIR = Path(r"C:\EladalahExport")
DOCX_DIR = BASE_DIR / "docs_full"
TXT_DIR = BASE_DIR / "txt_full"
LOG_FILE = BASE_DIR / "logs" / "full_export_log.csv"

DOCX_DIR.mkdir(parents=True, exist_ok=True)
TXT_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE.parent.mkdir(parents=True, exist_ok=True)

# Coordinates after Eladalah is maximized
TREE_X_REL = 1540
FIRST_ITEM_Y_REL = 278
TREE_ROW_STEP = 28

ROWS_PER_SCREEN = 18
MAX_SCREENS = 500
SCROLL_AMOUNT = -10

MIN_CHARS = 40
STOP_AFTER_NO_NEW_SCREENS = 4

def text_hash(text):
    return hashlib.md5(text.encode("utf-8", errors="ignore")).hexdigest()

def short_hash(text):
    return text_hash(text)[:10]

def clean_filename(text, max_len=80):
    text = re.sub(r'[\\/:*?"<>|]', " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return (text or "untitled")[:max_len]

def get_info(control, attr):
    try:
        return getattr(control.element_info, attr, "") or ""
    except Exception:
        return ""

def get_eladalah_window():
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
    return w, rect

def read_rt_txt(main_window):
    controls = main_window.descendants()

    for c in controls:
        auto_id = get_info(c, "automation_id")
        class_name = get_info(c, "class_name")

        if auto_id == "rt_txt" or "RichEdit20W" in class_name:
            tries = []

            try:
                tries.append(c.window_text())
            except Exception:
                pass

            try:
                texts = c.texts()
                if texts:
                    tries.append("\n".join(texts))
            except Exception:
                pass

            try:
                tries.append(c.get_value())
            except Exception:
                pass

            try:
                tries.append(c.iface_value.CurrentValue)
            except Exception:
                pass

            best = ""
            for t in tries:
                if t and len(t.strip()) > len(best):
                    best = t.strip()

            return best

    return ""

def load_done_hashes():
    done = set()

    if not LOG_FILE.exists():
        return done

    with LOG_FILE.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row.get("status") == "OK" and row.get("hash"):
                done.add(row["hash"])

    return done

def save_docx(text, docx_path, title):
    doc = Document()
    doc.add_heading(title, level=1)

    for line in text.splitlines():
        line = line.strip()
        if line:
            doc.add_paragraph(line)

    doc.save(str(docx_path))

def main():
    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 0.2

    done_hashes = load_done_hashes()
    print(f"Already exported hashes: {len(done_hashes)}")

    main_win, rect = get_eladalah_window()

    tree_x = rect.left + TREE_X_REL
    scroll_x = tree_x
    scroll_y = rect.top + 650

    print(f"Tree click X: {tree_x}")

    file_exists = LOG_FILE.exists()

    with LOG_FILE.open("a", encoding="utf-8-sig", newline="") as f:
        fieldnames = [
            "status",
            "screen",
            "row",
            "click_x",
            "click_y",
            "chars",
            "hash",
            "title",
            "docx_file",
            "txt_file",
            "error",
        ]
        writer = csv.DictWriter(f, fieldnames=fieldnames)

        if not file_exists or LOG_FILE.stat().st_size == 0:
            writer.writeheader()

        no_new_screens = 0

        for screen in range(1, MAX_SCREENS + 1):
            print(f"\n========== SCREEN {screen} ==========")

            new_this_screen = 0

            for row in range(ROWS_PER_SCREEN):
                item_y = rect.top + FIRST_ITEM_Y_REL + (row * TREE_ROW_STEP)

                try:
                    main_win.set_focus()
                    time.sleep(0.2)

                    pyautogui.click(tree_x, item_y)
                    time.sleep(1.1)

                    text = read_rt_txt(main_win)

                    if len(text.strip()) < MIN_CHARS:
                        print(f"SKIP screen {screen}, row {row+1}: too short")
                        writer.writerow({
                            "status": "SKIP_SHORT",
                            "screen": screen,
                            "row": row + 1,
                            "click_x": tree_x,
                            "click_y": item_y,
                            "chars": len(text.strip()),
                            "hash": "",
                            "title": "",
                            "docx_file": "",
                            "txt_file": "",
                            "error": "too short",
                        })
                        f.flush()
                        continue

                    h = text_hash(text)

                    if h in done_hashes:
                        print(f"DUPLICATE screen {screen}, row {row+1}")
                        writer.writerow({
                            "status": "DUPLICATE",
                            "screen": screen,
                            "row": row + 1,
                            "click_x": tree_x,
                            "click_y": item_y,
                            "chars": len(text),
                            "hash": h,
                            "title": "",
                            "docx_file": "",
                            "txt_file": "",
                            "error": "",
                        })
                        f.flush()
                        continue

                    lines = [x.strip() for x in text.splitlines() if x.strip()]
                    title = lines[0][:100] if lines else f"screen-{screen}-row-{row+1}"
                    safe_title = clean_filename(title)

                    base_name = f"s{screen:04d}-r{row+1:02d}-{short_hash(text)}-{safe_title}"
                    docx_path = DOCX_DIR / f"{base_name}.docx"
                    txt_path = TXT_DIR / f"{base_name}.txt"

                    txt_path.write_text(text, encoding="utf-8")
                    save_docx(text, docx_path, title)

                    done_hashes.add(h)
                    new_this_screen += 1

                    print(f"OK screen {screen}, row {row+1}: {len(text)} chars")
                    print(f"Saved: {docx_path.name}")

                    writer.writerow({
                        "status": "OK",
                        "screen": screen,
                        "row": row + 1,
                        "click_x": tree_x,
                        "click_y": item_y,
                        "chars": len(text),
                        "hash": h,
                        "title": title,
                        "docx_file": str(docx_path),
                        "txt_file": str(txt_path),
                        "error": "",
                    })
                    f.flush()

                except Exception as e:
                    print(f"FAILED screen {screen}, row {row+1}")
                    print(traceback.format_exc())

                    writer.writerow({
                        "status": "FAILED",
                        "screen": screen,
                        "row": row + 1,
                        "click_x": tree_x,
                        "click_y": item_y,
                        "chars": 0,
                        "hash": "",
                        "title": "",
                        "docx_file": "",
                        "txt_file": "",
                        "error": str(e),
                    })
                    f.flush()

            print(f"New items this screen: {new_this_screen}")

            if new_this_screen == 0:
                no_new_screens += 1
                print(f"No-new screen count: {no_new_screens}/{STOP_AFTER_NO_NEW_SCREENS}")
            else:
                no_new_screens = 0

            if no_new_screens >= STOP_AFTER_NO_NEW_SCREENS:
                print("Stopping: no new content after several screens.")
                break

            pyautogui.scroll(SCROLL_AMOUNT, x=scroll_x, y=scroll_y)
            time.sleep(1.5)

    print("\nFull export finished.")
    print(f"DOCX: {DOCX_DIR}")
    print(f"TXT:  {TXT_DIR}")
    print(f"LOG:  {LOG_FILE}")

if __name__ == "__main__":
    main()