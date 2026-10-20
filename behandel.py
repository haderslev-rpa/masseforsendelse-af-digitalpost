"""Forretningslogik for ét brev-item."""

from __future__ import annotations

import asyncio
import logging
from typing import Any

from automation_server_client import WorkItemError
from q_datafordeleren_api.functionality.datafordeler_use import (
    get_aktuel_navn_og_adresse,
)
from q_digitalpost.digital_post import DigitalPost
from q_digitalpost.models import DeliveryMethod, DigitalPostAddress
from q_gotenberg.fil_til_pdf import konverter_fil_til_pdf
from q_haderslev_vbo.automation_server.ats_find_items_by_box import (
    get_one_workitem_by_box_value,
)
from q_haderslev_vbo.automation_server.ats_update_item_data import (
    update_item_data,
)
from q_sharepoint_api.functionality.sp_documents import download_document
from q_word.udfyld_skabelon import udfyld_word_skabelon

from configuration import (
    DIGITAL_POST_QUEUE_ID,
    SOURCE_PROCESS,
    TEMPLATE_DRIVE_ITEM_ID,
    TEMPLATE_FILE_NAME,
    TEMPLATE_LIBRARY_NAME,
    TEMPLATE_SITE_NAME,
)
from models import ProcessingState

logger = logging.getLogger(__name__)

_TEMPLATE_CACHE: bytes | None = None


class PermanentItemError(WorkItemError):
    """Inputfejl, som kræver manuel rettelse."""


def _box(item: Any) -> dict[str, Any]:
    if not isinstance(getattr(item, "data", None), dict):
        raise PermanentItemError("Itemets data skal være en dictionary.")
    box = item.data.get("box")
    if not isinstance(box, dict):
        raise PermanentItemError("Itemet mangler item.data['box'].")
    return box


def _save(item: Any, state: str | None = None, **updates: Any) -> None:
    box = _box(item)
    box.update(updates)
    update_item_data(item.data, item=item, state=state)
    item.update(item.data)


def _validate_box(box: dict[str, Any]) -> None:
    required = (
        "cpr",
        "batch_id",
        "source_item_id",
        "subject",
        "bookmark_values",
        "fill_name_and_address",
        "delivery_method",
    )
    missing = [name for name in required if box.get(name) is None]
    if missing:
        raise PermanentItemError(
            "Itemets box mangler felt(er): " + ", ".join(missing)
        )
    cpr = "".join(character for character in str(box["cpr"]) if character.isdigit())
    if len(cpr) != 10:
        raise PermanentItemError("CPR skal indeholde præcis 10 cifre.")
    box["cpr"] = cpr
    if not isinstance(box["bookmark_values"], dict):
        raise PermanentItemError("bookmark_values skal være en dictionary.")
    if not isinstance(box["fill_name_and_address"], bool):
        raise PermanentItemError("fill_name_and_address skal være True eller False.")
    if box["delivery_method"] not in {
        DeliveryMethod.ONLY_DIGITAL_POST.value,
        DeliveryMethod.DIGITAL_OR_PHYSICAL_POST.value,
    }:
        raise PermanentItemError("Ukendt delivery_method.")


def _load_person_data(box: dict[str, Any]) -> dict[str, Any] | None:
    needs_person_data = (
        box["fill_name_and_address"]
        or box["delivery_method"]
        == DeliveryMethod.DIGITAL_OR_PHYSICAL_POST.value
    )
    if not needs_person_data:
        return None
    person_data = get_aktuel_navn_og_adresse(box["cpr"])
    if not isinstance(person_data, dict):
        raise RuntimeError("Datafordeleren returnerede ikke en dictionary.")
    if not person_data.get("findes"):
        raise PermanentItemError("CPR-nummeret blev ikke fundet i Datafordeleren.")
    if not person_data.get("kan_sendes_brev"):
        reason = person_data.get("kan_sendes_brev_aarsag") or "Ukendt årsag"
        raise PermanentItemError(f"Der kan ikke sendes brev til personen: {reason}")
    return person_data


def _build_bookmark_values(
    box: dict[str, Any],
    person_data: dict[str, Any] | None,
) -> dict[str, Any]:
    values = dict(box["bookmark_values"])
    if box["fill_name_and_address"]:
        if person_data is None:
            raise RuntimeError("Persondata mangler til navn og adresse.")
        values.update(
            {
                "NAVN": person_data.get("navn") or "",
                "ADRESSE": person_data.get("adresse_linje") or "",
                "POSTNUMMER_OG_BY": person_data.get("by_postnr") or "",
            }
        )
    return values


def _build_address(person_data: dict[str, Any] | None) -> DigitalPostAddress | None:
    if person_data is None:
        return None
    return DigitalPostAddress.from_datafordeler(person_data)


async def _get_template_bytes() -> bytes:
    global _TEMPLATE_CACHE
    if _TEMPLATE_CACHE is None:
        if not TEMPLATE_DRIVE_ITEM_ID.strip():
            raise RuntimeError(
                "TEMPLATE_DRIVE_ITEM_ID mangler i configuration.py."
            )
        _TEMPLATE_CACHE = await asyncio.to_thread(
            download_document,
            site_name=TEMPLATE_SITE_NAME,
            library_name=TEMPLATE_LIBRARY_NAME,
            drive_item_id=TEMPLATE_DRIVE_ITEM_ID,
        )
        if not _TEMPLATE_CACHE.startswith(b"PK"):
            raise RuntimeError("SharePoint-skabelonen er ikke en gyldig DOCX-fil.")
    return _TEMPLATE_CACHE


def _find_existing_digitalpost_item(source_item_id: str) -> dict[str, Any] | None:
    return get_one_workitem_by_box_value(
        queue_id=DIGITAL_POST_QUEUE_ID,
        box_path="source_item_id",
        value=source_item_id,
        statuses=None,
    )


async def behandel_item(item: Any) -> None:
    """Generér ét brev og aflever det til den centrale Digital Post-kø."""
    box = _box(item)
    _validate_box(box)

    existing = await asyncio.to_thread(
        _find_existing_digitalpost_item,
        box["source_item_id"],
    )
    if existing is not None:
        existing_box = existing.get("data", {}).get("box", {})
        _save(
            item,
            "5.0 Eksisterende Digital Post-bestilling fundet",
            processing_state=ProcessingState.DIGITALPOST_QUEUED,
            digitalpost_result={
                "status": "QUEUED",
                "forsendelses_id": existing_box.get("forsendelses_id"),
                "queue_item_id": existing.get("id"),
                "recovered_after_retry": True,
            },
        )
        item.complete(
            "Brevet fandtes allerede i den centrale Digital Post-kø."
        )
        return

    person_data = await asyncio.to_thread(_load_person_data, box)
    _save(
        item,
        "1.0 Input og persondata valideret",
        processing_state=ProcessingState.DATAFORDELER_LOADED,
    )

    template_bytes = await _get_template_bytes()
    _save(
        item,
        "2.0 Word-skabelon hentet",
        processing_state=ProcessingState.TEMPLATE_LOADED,
    )

    bookmark_values = _build_bookmark_values(box, person_data)
    filled_docx = await asyncio.to_thread(
        udfyld_word_skabelon,
        skabelon_bytes=template_bytes,
        vaerdier=bookmark_values,
    )
    _save(
        item,
        "3.0 Word-dokument genereret",
        processing_state=ProcessingState.DOCX_GENERATED,
    )

    pdf_bytes = await asyncio.to_thread(
        konverter_fil_til_pdf,
        filnavn=TEMPLATE_FILE_NAME,
        filindhold=filled_docx,
    )
    if not isinstance(pdf_bytes, bytes) or not pdf_bytes.startswith(b"%PDF-"):
        raise RuntimeError("q-gotenberg returnerede ikke en gyldig PDF.")
    _save(
        item,
        "4.0 PDF genereret",
        processing_state=ProcessingState.PDF_GENERATED,
    )

    delivery_method = DeliveryMethod(box["delivery_method"])
    address = (
        _build_address(person_data)
        if delivery_method == DeliveryMethod.DIGITAL_OR_PHYSICAL_POST
        else None
    )
    digital_post = DigitalPost()
    result = await digital_post.send_digital_post(
        pdf_content=pdf_bytes,
        file_name=box.get("output_file_name") or "brev.pdf",
        cpr_or_cvr=box["cpr"],
        subject=box["subject"],
        source_process=SOURCE_PROCESS,
        source_item_id=box["source_item_id"],
        delivery_method=delivery_method,
        address=address,
        attachments=[],
    )

    result_data = {
        "status": result.status,
        "reason": result.reason,
        "message": result.message,
        "forsendelses_id": result.forsendelses_id,
        "queue_id": result.queue_id,
        "document_count": result.document_count,
    }

    if result.status == "QUEUED":
        _save(
            item,
            "5.0 Brev lagt i den centrale Digital Post-kø",
            processing_state=ProcessingState.DIGITALPOST_QUEUED,
            digitalpost_result=result_data,
        )
        item.complete(
            "Brevet er genereret og lagt i den centrale Digital Post-kø. "
            f"Forsendelses-ID: {result.forsendelses_id}"
        )
        return

    if result.status == "NOT_SENT":
        _save(
            item,
            "5.1 Brev ikke sendt",
            processing_state=ProcessingState.NOT_SENT,
            digitalpost_result=result_data,
        )
        item.complete(
            "Brevet blev ikke sendt. "
            f"Årsag: {result.reason or result.message}"
        )
        return

    raise RuntimeError(f"Ukendt status fra q-digitalpost: {result.status!r}")
