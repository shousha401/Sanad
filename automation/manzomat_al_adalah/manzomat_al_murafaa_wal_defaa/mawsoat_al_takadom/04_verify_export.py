"""Step 04: verify list and document exports are complete and consistent."""

from __future__ import annotations

from pathlib import Path
import csv


LIST_FILE = Path("automation/output/manzomat_al_adalah/manzomat_al_murafaa_wal_defaa/mawsoat_al_takadom/list_export.csv")
DOCS_DIR = Path("automation/output/manzomat_al_adalah/manzomat_al_murafaa_wal_defaa/mawsoat_al_takadom/documents")


def verify_export() -> None:
    if not LIST_FILE.exists():
        raise RuntimeError("Missing list export. Run 02_export_list.py first.")
    if not DOCS_DIR.exists():
        raise RuntimeError("Missing documents folder. Run 03_export_documents.py first.")

    with LIST_FILE.open("r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))

    docx_files = list(DOCS_DIR.glob("*.docx"))

    if len(rows) != len(docx_files):
        raise RuntimeError(
            f"Verification failed: list rows={len(rows)} but exported docs={len(docx_files)}"
        )

    print("Verification passed.")
    print(f"Rows exported: {len(rows)}")
    print(f"Documents exported: {len(docx_files)}")


if __name__ == "__main__":
    verify_export()
