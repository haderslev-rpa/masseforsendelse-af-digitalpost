"""Manuel feeder: læs CPR-numre fra Excel og opret work items.

Kør:
    uv run python queue_feeder.py

Filen opretter rigtige work items i processens Automation Server-kø.
Ret kun værdierne i afsnittet MANUEL KONFIGURATION før hver kørsel.
"""

from __future__ import annotations

import re
from pathlib import Path

from automation_server_client import AutomationServer
from openpyxl import load_workbook

from configuration import WORKQUEUE_ID


# ============================================================================
# MANUEL KONFIGURATION - RET FØR HVER KØRSEL
# ============================================================================

# Excel-filen må kun indeholde CPR-numre i den valgte kolonne.
EXCEL_FILE = Path(__file__).resolve().parent / "tests_local" / "cpr.xlsx"
EXCEL_SHEET_NAME = None  # None bruger det aktive ark.
CPR_COLUMN = "A"
FIRST_DATA_ROW = 2

# Unikt ID for hele kørslen. Vælg en ny værdi for hver udsendelse.
# Eksempel: "budgetbrev-2026-10-20".
BATCH_ID = "RET-MIG-FOR-HVER-KOERSEL"

# Emnet, som modtageren ser i Digital Post.
SUBJECT = "RET MIG"
OUTPUT_FILE_NAME = "brev.pdf"

# True henter navn og adresse via q-datafordeleren og udfylder:
# [[NAVN]], [[ADRESSE]] og [[POSTNUMMER_OG_BY]]. #Så husk at tilføje disse bogmærker i docx. filen
FILL_NAME_AND_ADDRESS_BOOKMARKS = True

# Faste pladsholdere til Word-skabelonen for denne kørsel.
# CPR gemmes kun i work itemet og indsættes ikke automatisk i brevet.
BOOKMARK_VALUES = {
    "DATO": "20. oktober 2026",
    # "SAGSNUMMER": "123456",
}

# ONLY_DIGITAL_POST sender kun til personer tilmeldt Digital Post.
# DIGITAL_OR_PHYSICAL_POST kræver Datafordeler-adressen og tillader fysisk post.
DELIVERY_METHOD = "DIGITAL_OR_PHYSICAL_POST"

# Sikkerhedsstop. Ret til False, når konfigurationen er kontrolleret.
DRY_RUN = True


# ============================================================================
# FEEDER-LOGIK - SKAL NORMALT IKKE RETTES
# ============================================================================


def _normalize_cpr(value: object) -> str:
    if value is None:
        return ""
    # openpyxl kan returnere tal. zfill bevarer CPR med foranstillet nul.
    if isinstance(value, int):
        text = str(value).zfill(10)
    else:
        text = str(value).strip()
    return re.sub(r"\D", "", text)


def _validate_configuration() -> None:
    if WORKQUEUE_ID <= 0:
        raise ValueError("WORKQUEUE_ID skal udfyldes i configuration.py.")
    if not EXCEL_FILE.is_file():
        raise FileNotFoundError(f"Excel-filen blev ikke fundet: {EXCEL_FILE}")
    if not BATCH_ID.strip() or BATCH_ID == "RET-MIG-FOR-HVER-KOERSEL":
        raise ValueError("Vælg et unikt BATCH_ID øverst i queue_feeder.py.")
    if not SUBJECT.strip() or SUBJECT == "RET MIG":
        raise ValueError("Udfyld SUBJECT øverst i queue_feeder.py.")
    if DELIVERY_METHOD not in {
        "ONLY_DIGITAL_POST",
        "DIGITAL_OR_PHYSICAL_POST",
    }:
        raise ValueError("DELIVERY_METHOD er ugyldig.")
    if DELIVERY_METHOD == "DIGITAL_OR_PHYSICAL_POST" and not FILL_NAME_AND_ADDRESS_BOOKMARKS:
        raise ValueError(
            "DIGITAL_OR_PHYSICAL_POST kræver navn og adresse fra Datafordeleren."
        )
    if not isinstance(BOOKMARK_VALUES, dict):
        raise TypeError("BOOKMARK_VALUES skal være en dictionary.")


def _read_cpr_numbers() -> list[str]:
    workbook = load_workbook(EXCEL_FILE, read_only=True, data_only=True)
    worksheet = (
        workbook[EXCEL_SHEET_NAME]
        if EXCEL_SHEET_NAME
        else workbook.active
    )
    cpr_numbers: list[str] = []
    errors: list[str] = []
    seen: set[str] = set()
    try:
        for row_number in range(FIRST_DATA_ROW, worksheet.max_row + 1):
            cpr = _normalize_cpr(worksheet[f"{CPR_COLUMN}{row_number}"].value)
            if not cpr:
                continue
            if len(cpr) != 10:
                errors.append(f"Række {row_number}: CPR skal have 10 cifre.")
                continue
            if cpr in seen:
                errors.append(f"Række {row_number}: CPR forekommer flere gange.")
                continue
            seen.add(cpr)
            cpr_numbers.append(cpr)
    finally:
        workbook.close()
    if errors:
        raise ValueError("Excel-filen indeholder fejl:\n" + "\n".join(errors))
    if not cpr_numbers:
        raise ValueError("Excel-filen indeholder ingen CPR-numre.")
    return cpr_numbers


def main() -> None:
    _validate_configuration()
    cpr_numbers = _read_cpr_numbers()

    print(f"Batch-ID: {BATCH_ID}")
    print(f"Antal unikke CPR-numre: {len(cpr_numbers)}")
    print(f"Emne: {SUBJECT}")
    print(f"Leveringsmetode: {DELIVERY_METHOD}")

    if DRY_RUN:
        print("DRY_RUN=True. Der er ikke oprettet work items.")
        print("Ret DRY_RUN til False efter kontrollen.")
        return

    automation_server = AutomationServer.from_environment()
    workqueue = automation_server.workqueue(WORKQUEUE_ID)

    created = 0
    for sequence, cpr in enumerate(cpr_numbers, start=1):
        source_item_id = f"{BATCH_ID}:{sequence:06d}:{cpr}"
        item_data = {
            "box": {
                "cpr": cpr,
                "batch_id": BATCH_ID,
                "batch_sequence": sequence,
                "source_item_id": source_item_id,
                "subject": SUBJECT,
                "output_file_name": OUTPUT_FILE_NAME,
                "bookmark_values": dict(BOOKMARK_VALUES),
                "fill_name_and_address": FILL_NAME_AND_ADDRESS_BOOKMARKS,
                "delivery_method": DELIVERY_METHOD,
                "processing_state": "READY",
                "digitalpost_result": None,
                "last_error": None,
            },
            "defer": None,
            "state": ["Oprettet af queue_feeder.py"],
            "status": {
                "status": "New",
                "status_code": "Klar til dokumentgenerering",
            },
        }
        workqueue.add_item(
            data=item_data,
            reference=source_item_id,
        )
        created += 1
        if created % 100 == 0:
            print(f"Oprettet: {created}/{len(cpr_numbers)}")

    print(f"Færdig. Oprettede {created} work items.")


if __name__ == "__main__":
    main()
