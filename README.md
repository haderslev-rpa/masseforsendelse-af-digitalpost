# masseforsendelse-af-digitalpost

Hver udsendelse styres af sin egen JSON-fil. `queue_feeder.py` indeholder ingen konkret udsendelsesopsætning.

## Ny udsendelse

1. Kopiér `forsendelse_template.json`.
2. Læg kopien i `forsendelser/`.
3. Omdøb kopien, fx `orienteringsbrev-2026-10-20.json`.
4. Udfyld alle værdier.
5. Behold først `"dry_run": true`.
6. Kør `uv run python queue_feeder.py`.
7. Skriv JSON-filens navn i terminalen.
8. Kontrollér udskriften.
9. Skift til `"dry_run": false` og kør samme JSON igen.

Mappen `forsendelser/` er i `.gitignore`. Den blanke `forsendelse_template.json` ligger i repositoryets rod og commits som den fælles skabelon.

Den fulde feltforklaring står øverst i `queue_feeder.py`.

## JSON-eksempel

```json
{
  "batch_id": "orienteringsbrev-2026-10-20",
  "sharepoint_folder_name": "20-10-2026",
  "excel": {
    "file_name": "modtagere.xlsx",
    "sheet_name": null,
    "cpr_column": "A",
    "first_data_row": 2
  },
  "document": {
    "mode": "PDF",
    "file_name": "Fællesbrev.pdf",
    "output_file_name": "Fællesbrev.pdf"
  },
  "subject": "Orienteringsbrev fra Haderslev Kommune",
  "word": {
    "bookmark_values": {},
    "fill_name_and_address_bookmarks": false
  },
  "attachments": {
    "Vejledning.pdf": "Vejledning version 3.pdf"
  },
  "delivery_method": "ONLY_DIGITAL_POST",
  "dry_run": true
}
```
