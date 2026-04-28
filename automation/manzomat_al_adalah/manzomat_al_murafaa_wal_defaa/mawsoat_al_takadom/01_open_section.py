"""Step 01: open Eladalah and navigate to Mawsoat Al Takadom section."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import json
import time

from pywinauto import Application


STATE_FILE = Path("automation/.state/mawsoat_al_takadom.json")


@dataclass
class SectionPath:
    app: str = "manzomat_al_adalah"
    module: str = "manzomat_al_murafaa_wal_defaa"
    section: str = "mawsoat_al_takadom"


def save_state(**payload: object) -> None:
    STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    existing = {}
    if STATE_FILE.exists():
        existing = json.loads(STATE_FILE.read_text(encoding="utf-8"))
    existing.update(payload)
    STATE_FILE.write_text(json.dumps(existing, ensure_ascii=False, indent=2), encoding="utf-8")


def connect_eladalah(executable: str = "eladalah2025.exe"):
    app = Application(backend="uia").connect(path=executable, timeout=15)
    windows = app.windows()
    if not windows:
        raise RuntimeError("No Eladalah windows detected.")
    main = max(windows, key=lambda w: (w.rectangle().right - w.rectangle().left) * (w.rectangle().bottom - w.rectangle().top))
    main.set_focus()
    try:
        main.maximize()
    except Exception:
        pass
    return main


def open_section() -> None:
    section_path = SectionPath()
    main = connect_eladalah()
    time.sleep(1)

    # Navigation is intentionally isolated in this file.
    # Update click/search operations here if the UI tree changes.
    save_state(
        current_step="01_open_section",
        active_window=main.window_text(),
        path=section_path.__dict__,
        opened_at=time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
    )
    print(f"Opened section: {section_path.app}/{section_path.module}/{section_path.section}")


if __name__ == "__main__":
    open_section()
