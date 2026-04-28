# Mawsoat Al Takadom automation

This folder maps 1:1 to the Eladalah app path:

- `manzomat_al_adalah`
- `manzomat_al_murafaa_wal_defaa`
- `mawsoat_al_takadom`

## Script flow

1. `01_open_section.py`: connect to Eladalah and open the target section.
2. `02_export_list.py`: export list/grid metadata to CSV.
3. `03_export_documents.py`: export each selected row to a separate DOCX file.
4. `04_verify_export.py`: verify that exported document count matches list rows.

## Why split scripts

Each automation/export responsibility is isolated in its own file so changes in one step do not break unrelated steps.

## Run order

```bash
python automation/manzomat_al_adalah/manzomat_al_murafaa_wal_defaa/mawsoat_al_takadom/01_open_section.py
python automation/manzomat_al_adalah/manzomat_al_murafaa_wal_defaa/mawsoat_al_takadom/02_export_list.py
python automation/manzomat_al_adalah/manzomat_al_murafaa_wal_defaa/mawsoat_al_takadom/03_export_documents.py
python automation/manzomat_al_adalah/manzomat_al_murafaa_wal_defaa/mawsoat_al_takadom/04_verify_export.py
```
