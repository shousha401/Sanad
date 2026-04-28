"""Step 02: export the list/grid metadata for Mawsoat Al Takadom."""

from __future__ import annotations

from pathlib import Path
import csv
import json
import time


STATE_FILE = Path("automation/.state/mawsoat_al_takadom.json")
OUT_FILE = Path("automation/output/manzomat_al_adalah/manzomat_al_murafaa_wal_defaa/mawsoat_al_takadom/list_export.csv")


def load_state() -> dict:
    if not STATE_FILE.exists():
        raise RuntimeError("Step 01 must run first: missing state file.")
    return json.loads(STATE_FILE.read_text(encoding="utf-8"))


def export_list() -> None:
    state = load_state()
    OUT_FILE.parent.mkdir(parents=True, exist_ok=True)

    # Placeholder rows until list extraction selectors are finalized.
    rows = [
        {"row_index": 1, "title": "sample-item-1", "captured_at": time.time()},
        {"row_index": 2, "title": "sample-item-2", "captured_at": time.time()},
    ]

    with OUT_FILE.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["row_index", "title", "captured_at"])
        writer.writeheader()
        writer.writerows(rows)

    state["current_step"] = "02_export_list"
    state["list_export"] = str(OUT_FILE)
    STATE_FILE.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"List exported to: {OUT_FILE}")


if __name__ == "__main__":
    export_list()
