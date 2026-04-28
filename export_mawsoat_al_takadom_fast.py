from __future__ import annotations

from pathlib import Path
from datetime import datetime, timezone
import csv
import hashlib
import json
import re
import time

import pyautogui
from pywinauto import Application

# =========================
# Settings
# =========================
SAVE_TXT = True
SAVE_DOCX = False
FAST_MODE = True
MAX_ITEMS = None

BASE_DIR = Path(r"C:\EladalahExport")
OUT_DIR = BASE_DIR / "mawsoat_al_takadom_output"
TXT_DIR = OUT_DIR / "txt"
LOG_FILE = OUT_DIR / "logs" / "full_export_log.csv"
MANIFEST_FILE = OUT_DIR / "manifest.json"

MIN_TEXT_CHARS = 30
CLICK_WAIT_SECONDS = 0.20 if FAST_MODE else 0.40
READ_WAIT_SECONDS = 0.12 if FAST_MODE else 0.25
PAGE_SETTLE_SECONDS = 0.22 if FAST_MODE else 0.50
NEXT_PAGE_KEYS = "{PGDN}"
MAX_PAGES = 200


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def sanitize_filename(text: str, max_len: int = 90) -> str:
    text = re.sub(r'[\\/:*?"<>|]', " ", (text or "").strip())
    text = re.sub(r"\s+", " ", text).strip()
    return (text or "untitled")[:max_len]


def text_hash(text: str) -> str:
    return hashlib.md5(text.encode("utf-8", errors="ignore")).hexdigest()


def get_attr(control, attr: str) -> str:
    try:
        return getattr(control.element_info, attr, "") or ""
    except Exception:
        return ""


def get_name(control) -> str:
    try:
        name = control.window_text() or get_attr(control, "name")
        return (name or "").strip()
    except Exception:
        return (get_attr(control, "name") or "").strip()


def connect_main_window():
    app = Application(backend="uia").connect(path="eladalah2025.exe", timeout=10)
    candidates = []
    for w in app.windows():
        try:
            title = w.window_text() or ""
            rect = w.rectangle()
            area = max((rect.right - rect.left), 0) * max((rect.bottom - rect.top), 0)
            if area > 0 and ("العدالة" in title or "2025" in title or "Eladalah" in title):
                candidates.append((area, w))
        except Exception:
            pass

    if not candidates:
        raise RuntimeError("Could not find Eladalah main window.")

    candidates.sort(key=lambda x: x[0], reverse=True)
    win = candidates[0][1]
    try:
        win.set_focus()
    except Exception:
        pass
    return win


def find_tree_control(main_win):
    # One-time scan + cache.
    for c in main_win.descendants():
        auto_id = get_attr(c, "automation_id")
        class_name = get_attr(c, "class_name")
        ctrl_type = get_attr(c, "control_type")
        if auto_id == "tv_Subjects" or "SysTreeView32" in class_name or ctrl_type == "Tree":
            return c
    return None


def find_rt_control(main_win):
    # One-time scan + cache.
    for c in main_win.descendants():
        auto_id = get_attr(c, "automation_id")
        class_name = get_attr(c, "class_name")
        if auto_id == "rt_txt" or "RichEdit20W" in class_name:
            return c
    return None


def read_rt_text(rt_ctrl) -> str:
    tries = []
    try:
        tries.append(rt_ctrl.window_text())
    except Exception:
        pass
    try:
        txts = rt_ctrl.texts()
        if txts:
            tries.append("\n".join(txts))
    except Exception:
        pass
    try:
        tries.append(rt_ctrl.get_value())
    except Exception:
        pass
    try:
        tries.append(rt_ctrl.iface_value.CurrentValue)
    except Exception:
        pass

    best = ""
    for t in tries:
        val = (t or "").strip()
        if len(val) > len(best):
            best = val
    return best


def get_visible_rows_once(tree_ctrl):
    """Read visible TreeItem rows once for current page/list viewport."""
    items = []
    tree_rect = tree_ctrl.rectangle()

    for item in tree_ctrl.descendants(control_type="TreeItem"):
        try:
            if not item.is_visible():
                continue
        except Exception:
            pass

        try:
            r = item.rectangle()
            if r.bottom <= tree_rect.top or r.top >= tree_rect.bottom:
                continue
        except Exception:
            continue

        name = get_name(item)
        items.append((item, name, r))

    items.sort(key=lambda x: (x[2].top, x[1]))
    if MAX_ITEMS is not None:
        items = items[:MAX_ITEMS]
    return items


def page_signature(rows) -> tuple:
    sig = []
    for _, name, rect in rows:
        sig.append(((name or "(blank)"), int(rect.top / 16), int(rect.bottom / 16)))
    return tuple(sig)


def click_row(row_ctrl):
    try:
        row_ctrl.click_input()
        return True
    except Exception:
        try:
            r = row_ctrl.rectangle()
            pyautogui.click((r.left + r.right) // 2, (r.top + r.bottom) // 2)
            return True
        except Exception:
            return False


def ensure_dirs():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    TXT_DIR.mkdir(parents=True, exist_ok=True)
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)


def write_manifest(data: dict):
    MANIFEST_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def main():
    ensure_dirs()
    pyautogui.PAUSE = 0.0

    main_win = connect_main_window()
    tree_ctrl = find_tree_control(main_win)
    rt_ctrl = find_rt_control(main_win)

    if tree_ctrl is None:
        raise RuntimeError("Could not find tv_Subjects tree/list control.")
    if rt_ctrl is None:
        raise RuntimeError("Could not find rt_txt text control.")

    rows_written = 0
    ok_items = 0
    skipped_items = 0
    seen_hashes = set()
    seen_signatures = set()
    repeated_page_stop = False

    log_exists = LOG_FILE.exists() and LOG_FILE.stat().st_size > 0
    with LOG_FILE.open("a", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=[
                "ts",
                "status",
                "page",
                "row",
                "visible_title",
                "chars",
                "hash",
                "txt_file",
                "error",
            ],
        )
        if not log_exists:
            writer.writeheader()

        for page in range(1, MAX_PAGES + 1):
            # Read page rows once.
            visible_rows = get_visible_rows_once(tree_ctrl)
            sig = page_signature(visible_rows)

            if sig in seen_signatures:
                repeated_page_stop = True
                writer.writerow(
                    {
                        "ts": now_iso(),
                        "status": "STOP_REPEATED_PAGE_SIGNATURE",
                        "page": page,
                        "row": 0,
                        "visible_title": "",
                        "chars": len(visible_rows),
                        "hash": "",
                        "txt_file": "",
                        "error": "Detected repeated visible page signature.",
                    }
                )
                rows_written += 1
                f.flush()
                break

            seen_signatures.add(sig)

            for row_index, (row_ctrl, row_name, _) in enumerate(visible_rows, start=1):
                title = row_name or f"row-{row_index}"
                status = "OK"
                err = ""
                txt_rel = ""
                body = ""

                try:
                    clicked = click_row(row_ctrl)
                    if not clicked:
                        raise RuntimeError("row_click_failed")

                    time.sleep(CLICK_WAIT_SECONDS)
                    body = read_rt_text(rt_ctrl)
                    if not body:
                        time.sleep(READ_WAIT_SECONDS)
                        body = read_rt_text(rt_ctrl)

                    if len(body.strip()) < MIN_TEXT_CHARS:
                        status = "SKIP_SHORT"
                        skipped_items += 1
                    else:
                        h = text_hash(body)
                        if h in seen_hashes:
                            status = "DUPLICATE_TEXT"
                        else:
                            seen_hashes.add(h)
                            if SAVE_TXT:
                                filename = f"p{page:03d}-r{row_index:03d}-{h[:10]}-{sanitize_filename(title)}.txt"
                                txt_path = TXT_DIR / filename
                                txt_path.write_text(body, encoding="utf-8")
                                txt_rel = f"txt/{filename}"
                            ok_items += 1
                except Exception as ex:
                    status = "ERROR"
                    err = str(ex)

                writer.writerow(
                    {
                        "ts": now_iso(),
                        "status": status,
                        "page": page,
                        "row": row_index,
                        "visible_title": title,
                        "chars": len(body or ""),
                        "hash": text_hash(body) if body else "",
                        "txt_file": txt_rel,
                        "error": err,
                    }
                )
                rows_written += 1

            f.flush()

            try:
                tree_ctrl.set_focus()
            except Exception:
                pass
            from pywinauto.keyboard import send_keys

            send_keys(NEXT_PAGE_KEYS)
            time.sleep(PAGE_SETTLE_SECONDS)

    manifest = {
        "script": "export_mawsoat_al_takadom_fast.py",
        "started_assumption": "Eladalah open and موسوعة التقادم already opened manually.",
        "output_dir": str(OUT_DIR),
        "save_txt": SAVE_TXT,
        "save_docx": SAVE_DOCX,
        "fast_mode": FAST_MODE,
        "max_items": MAX_ITEMS,
        "rows_logged": rows_written,
        "ok_items": ok_items,
        "skipped_items": skipped_items,
        "unique_text_hashes": len(seen_hashes),
        "seen_page_signatures": len(seen_signatures),
        "stopped_on_repeated_page_signature": repeated_page_stop,
        "finished_at_utc": now_iso(),
    }
    write_manifest(manifest)


if __name__ == "__main__":
    main()
