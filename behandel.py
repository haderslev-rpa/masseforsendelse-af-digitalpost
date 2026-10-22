"""Forretningslogik for masseforsendelse-af-digitalpost.

Processen behandler ét Automation Server-item ad gangen:

1. Valider item.data["box"].
2. Hent persondata fra Datafordeleren.
3. Kontrollér, om der kan sendes brev.
4. Hent og klargør hoveddokumentet.
5. Hent eventuelle bilag.
6. Aflever dokumenterne til den centrale Digital Post-kø.
7. Afslut source-itemet som QUEUED eller NOT_SENT.

Next item håndteres udelukkende i main.py gennem workqueue-objektet.
Denne fil indeholder derfor ingen opslag efter næste item og ingen import af
ats_find_items_by_box.
"""

from __future__ import annotations

import asyncio
import logging
from typing import Any

from automation_server_client import WorkItemError
from q_datafordeleren_api.functionality.datafordeler_use import (
    get_aktuel_navn_og_adresse,
)
from q_digitalpost.digital_post import DigitalPost
from q_digitalpost.models import (
    DeliveryMethod,
    DigitalPostAddress,
    DigitalPostAttachment,
)
from q_gotenberg.fil_til_pdf import konverter_fil_til_pdf
from q_haderslev_vbo.automation_server.ats_find_state import find_state
from q_haderslev_vbo.automation_server.ats_update_item_data import (
    update_item_data,
)
from q_word.udfyld_skabelon import udfyld_word_skabelon

from configuration import SOURCE_PROCESS
from sharepoint_template_service import download_file_by_drive_ids


logger = logging.getLogger(__name__)


# ==========================================================
# STATES
# ==========================================================


class States:
    """Processtates, som gemmes på source-itemet."""

    INPUT_VALIDATED = "1.0 Input valideret"
    PERSON_DATA_LOADED = "2.0 Persondata hentet"
    MAIN_DOCUMENT_READY = "3.0 Hoveddokument klargjort"
    ATTACHMENTS_READY = "4.0 Bilag klargjort"
    DIGITALPOST_QUEUED = "5.0 Lagt i Digital Post-kø"
    NOT_SENT = "5.1 Ikke sendt"


# ==========================================================
# FEJLTYPER
# ==========================================================


class PermanentItemError(WorkItemError):
    """Inputfejl, som kun kan løses ved at rette itemet eller opsætningen."""


# ==========================================================
# CACHE
# ==========================================================


# Filerne er fælles for mange items.
# Cache pr. drive + item betyder, at samme Word-skabelon,
# PDF eller bilag kun hentes én gang i samme proceskørsel.
_FILE_CACHE: dict[tuple[str, str], bytes] = {}


# ==========================================================
# GENERELLE HELPERS
# ==========================================================


def _get_box(item: Any) -> dict[str, Any]:
    """Hent item.data['box'] og returnér den som dictionary."""
    if not isinstance(getattr(item, "data", None), dict):
        raise PermanentItemError(
            "Itemets data skal være en dictionary."
        )

    box = item.data.get("box")

    if not isinstance(box, dict):
        raise PermanentItemError(
            "Itemet mangler item.data['box']."
        )

    return box


def _mangler_state(
    data: dict[str, Any],
    state: str,
) -> bool:
    """Returnér True, når den ønskede state ikke allerede findes."""
    return not find_state(data, state)


def _set_state(
    item: Any,
    state: str,
    *,
    box_updates: dict[str, Any] | None = None,
    status: str | None = None,
    status_code: str | None = None,
) -> None:
    """Gem box-opdateringer, status og state på work itemet."""
    update_item_data(
        item.data,
        item=item,
        box_updates=box_updates,
        status=status,
        status_code=status_code,
        state=state,
    )

    item.update(item.data)


def _log_step(
    step: str,
    message: str,
) -> None:
    """Skriv en ensartet logbesked for et procestrin."""
    logger.info(
        "[%s] %s",
        step,
        message,
    )


def _require_text(
    value: Any,
    field_name: str,
) -> str:
    """Kontrollér og returnér obligatorisk tekst."""
    if not isinstance(value, str) or not value.strip():
        raise PermanentItemError(
            f"{field_name} skal være udfyldt med tekst."
        )

    return value.strip()


# ==========================================================
# INPUTVALIDERING
# ==========================================================


def _validate_file_reference(
    value: Any,
    field_name: str,
    expected_extension: str,
) -> dict[str, str]:
    """Validér og returnér en minimal SharePoint-filreference."""
    if not isinstance(value, dict):
        raise PermanentItemError(
            f"{field_name} skal være en dictionary."
        )

    drive_id = _require_text(
        value.get("drive_id"),
        f"{field_name}.drive_id",
    )
    drive_item_id = _require_text(
        value.get("drive_item_id"),
        f"{field_name}.drive_item_id",
    )
    file_name = _require_text(
        value.get("file_name"),
        f"{field_name}.file_name",
    )

    if not file_name.lower().endswith(expected_extension):
        raise PermanentItemError(
            f"{field_name}.file_name skal ende på "
            f"{expected_extension}."
        )

    return {
        "drive_id": drive_id,
        "drive_item_id": drive_item_id,
        "file_name": file_name,
    }


def _validate_box(
    item: Any,
    box: dict[str, Any],
) -> None:
    """Validér alle arbejdsdata, før eksterne systemer kaldes."""
    item_reference = _require_text(
        getattr(item, "reference", None),
        "item.reference",
    )

    if ":" not in item_reference:
        raise PermanentItemError(
            "item.reference skal være den unikke reference fra feederen."
        )

    cpr = "".join(
        character
        for character in str(box.get("cpr", ""))
        if character.isdigit()
    )

    if len(cpr) != 10:
        raise PermanentItemError(
            "box.cpr skal indeholde præcis 10 cifre."
        )

    box["cpr"] = cpr
    box["subject"] = _require_text(
        box.get("subject"),
        "box.subject",
    )
    box["shipment_configuration_file"] = _require_text(
        box.get("shipment_configuration_file"),
        "box.shipment_configuration_file",
    )
    box["output_file_name"] = _require_text(
        box.get("output_file_name"),
        "box.output_file_name",
    )

    if not box["output_file_name"].lower().endswith(".pdf"):
        raise PermanentItemError(
            "box.output_file_name skal ende på .pdf."
        )

    document_mode = _require_text(
        box.get("document_mode"),
        "box.document_mode",
    ).upper()

    if document_mode not in {"WORD", "PDF"}:
        raise PermanentItemError(
            "box.document_mode skal være WORD eller PDF."
        )

    box["document_mode"] = document_mode

    expected_extension = (
        ".docx"
        if document_mode == "WORD"
        else ".pdf"
    )

    box["main_document"] = _validate_file_reference(
        box.get("main_document"),
        "box.main_document",
        expected_extension,
    )

    bookmark_values = box.get("bookmark_values")

    if not isinstance(bookmark_values, dict):
        raise PermanentItemError(
            "box.bookmark_values skal være en dictionary."
        )

    fill_name_and_address = box.get(
        "fill_name_and_address"
    )

    if not isinstance(fill_name_and_address, bool):
        raise PermanentItemError(
            "box.fill_name_and_address skal være true eller false."
        )

    delivery_method = _require_text(
        box.get("delivery_method"),
        "box.delivery_method",
    ).upper()

    allowed_delivery_methods = {
        DeliveryMethod.ONLY_DIGITAL_POST.value,
        DeliveryMethod.DIGITAL_OR_PHYSICAL_POST.value,
    }

    if delivery_method not in allowed_delivery_methods:
        raise PermanentItemError(
            "box.delivery_method har en ukendt værdi."
        )

    if (
        document_mode == "PDF"
        and delivery_method
        == DeliveryMethod.DIGITAL_OR_PHYSICAL_POST.value
    ):
        raise PermanentItemError(
            "PDF-mode må kun bruges med ONLY_DIGITAL_POST."
        )

    if (
        document_mode == "WORD"
        and delivery_method
        == DeliveryMethod.DIGITAL_OR_PHYSICAL_POST.value
        and not fill_name_and_address
    ):
        raise PermanentItemError(
            "DIGITAL_OR_PHYSICAL_POST kræver, at navn og adresse "
            "hentes fra Datafordeleren."
        )

    box["delivery_method"] = delivery_method

    attachments = box.get("attachments")

    if not isinstance(attachments, list):
        raise PermanentItemError(
            "box.attachments skal være en liste."
        )

    validated_attachments: list[dict[str, str]] = []

    for index, attachment in enumerate(
        attachments,
        start=1,
    ):
        validated_attachments.append(
            _validate_file_reference(
                attachment,
                f"box.attachments[{index}]",
                ".pdf",
            )
        )

    box["attachments"] = validated_attachments


# ==========================================================
# DATAFORDELER
# ==========================================================


def _load_person_data(
    box: dict[str, Any],
) -> dict[str, Any]:
    """Hent og returnér borgerens persondata fra Datafordeleren.

    Persondata hentes altid, fordi processen skal kontrollere,
    om der kan sendes brev, også når forsendelsen bruger en
    færdig PDF.

    Output:
        En dictionary med blandt andet status, navn, adresse,
        kan_sendes_brev og kan_sendes_brev_aarsag.
    """
    person_data = get_aktuel_navn_og_adresse(
        box["cpr"]
    )

    if not isinstance(person_data, dict):
        raise RuntimeError(
            "Datafordeleren returnerede ikke en dictionary."
        )

    if not person_data.get("findes"):
        raise PermanentItemError(
            "CPR-nummeret blev ikke fundet i Datafordeleren."
        )

    return person_data


def _build_not_sent_message(
    person_data: dict[str, Any],
) -> str:
    """Byg afslutningsbeskeden, når der ikke kan sendes brev."""
    if person_data.get("er_doed"):
        return "Brev ikke sendt: Borger er død"

    reason = (
        person_data.get("kan_sendes_brev_aarsag")
        or "Ukendt årsag"
    )

    return f"Brev ikke sendt: {reason}"


# ==========================================================
# WORD-VÆRDIER
# ==========================================================


def _build_bookmark_values(
    box: dict[str, Any],
    person_data: dict[str, Any],
) -> dict[str, Any]:
    """Byg og returnér alle værdier til q-word."""
    values = dict(box["bookmark_values"])

    if box["fill_name_and_address"]:
        values.update(
            {
                "NAVN": person_data.get("navn") or "",
                "ADRESSE": (
                    person_data.get("adresse_linje")
                    or ""
                ),
                "POSTNUMMER_OG_BY": (
                    person_data.get("by_postnr")
                    or ""
                ),
            }
        )

    return values


# ==========================================================
# SHAREPOINT-FILER
# ==========================================================


async def _download_sharepoint_file(
    file_reference: dict[str, str],
) -> bytes:
    """Download en SharePoint-fil med lokal procescache."""
    cache_key = (
        file_reference["drive_id"],
        file_reference["drive_item_id"],
    )

    if cache_key not in _FILE_CACHE:
        content = await asyncio.to_thread(
            download_file_by_drive_ids,
            drive_id=file_reference["drive_id"],
            drive_item_id=file_reference["drive_item_id"],
        )

        if not isinstance(content, bytes) or not content:
            raise RuntimeError(
                "SharePoint returnerede en tom fil."
            )

        _FILE_CACHE[cache_key] = content

    return _FILE_CACHE[cache_key]


# ==========================================================
# HOVEDDOKUMENT
# ==========================================================


async def _build_main_pdf(
    box: dict[str, Any],
    person_data: dict[str, Any],
) -> bytes:
    """Klargør hoveddokumentet og returnér PDF-indhold som bytes."""
    main_document = box["main_document"]

    content = await _download_sharepoint_file(
        main_document
    )

    if box["document_mode"] == "PDF":
        if not content.startswith(b"%PDF-"):
            raise PermanentItemError(
                "Hoveddokumentet er ikke en gyldig PDF-fil."
            )

        return content

    if not content.startswith(b"PK"):
        raise PermanentItemError(
            "Word-skabelonen er ikke en gyldig DOCX-fil."
        )

    bookmark_values = _build_bookmark_values(
        box,
        person_data,
    )

    filled_docx = await asyncio.to_thread(
        udfyld_word_skabelon,
        skabelon_bytes=content,
        vaerdier=bookmark_values,
    )

    pdf_bytes = await asyncio.to_thread(
        konverter_fil_til_pdf,
        filnavn=main_document["file_name"],
        filindhold=filled_docx,
    )

    if (
        not isinstance(pdf_bytes, bytes)
        or not pdf_bytes.startswith(b"%PDF-")
    ):
        raise RuntimeError(
            "q-gotenberg returnerede ikke en gyldig PDF."
        )

    return pdf_bytes


# ==========================================================
# BILAG
# ==========================================================


async def _build_attachments(
    box: dict[str, Any],
) -> list[DigitalPostAttachment]:
    """Download fælles bilag og returnér q-digitalpost-modeller."""
    result: list[DigitalPostAttachment] = []

    for attachment in box["attachments"]:
        content = await _download_sharepoint_file(
            attachment
        )

        if not content.startswith(b"%PDF-"):
            raise PermanentItemError(
                "Bilaget er ikke en gyldig PDF: "
                f"{attachment['file_name']}"
            )

        result.append(
            DigitalPostAttachment(
                content=content,
                file_name=attachment["file_name"],
            )
        )

    return result


# ==========================================================
# FYSISK ADRESSE
# ==========================================================


def _build_address(
    delivery_method: DeliveryMethod,
    person_data: dict[str, Any],
) -> DigitalPostAddress | None:
    """Byg en fysisk adresse, når leveringsmetoden kræver det."""
    if (
        delivery_method
        != DeliveryMethod.DIGITAL_OR_PHYSICAL_POST
    ):
        return None

    return DigitalPostAddress.from_datafordeler(
        person_data
    )


# ==========================================================
# AFSLUTNINGSHJÆLPERE
# ==========================================================


def _complete_not_sent(
    item: Any,
    *,
    message: str,
    reason: str,
) -> None:
    """Afslut itemet som completed uden at sende brevet."""
    result_data = {
        "status": "NOT_SENT",
        "reason": reason,
        "message": message,
        "forsendelses_id": None,
        "queue_id": None,
        "document_count": 0,
    }

    _set_state(
        item,
        States.NOT_SENT,
        box_updates={
            "processing_state": "NOT_SENT",
            "digitalpost_result": result_data,
            "last_error": None,
        },
        status="Completed",
        status_code=message,
    )

    item.complete(message)


# ==========================================================
# PROCESSEN FOR ÉT WORK ITEM
# ==========================================================


async def behandel_item(item: Any) -> None:
    """Behandl ét source-item og aflever brevet til Digital Post-køen.

    Output:
        Funktionen returnerer None.

        Itemet afsluttes som completed, når:
        - Datafordeleren oplyser, at der ikke kan sendes brev,
        - brevet er lagt i Digital Post-kø,
        - eller Digital Post returnerer NOT_SENT.

        Fejl bobler op til main.py, som ejer procesfejlhåndteringen.
    """
    data = item.data
    box = _get_box(item)

    # ==========================================================
    # TRIN 1: INPUT_VALIDATED
    # ==========================================================

    step = "INPUT_VALIDATED"
    state = getattr(States, step)

    if _mangler_state(data, state):
        _log_step(step, "Start")

        _validate_box(
            item,
            box,
        )

        _set_state(
            item,
            state,
            box_updates={
                "processing_state": "INPUT_VALIDATED",
                "last_error": None,
            },
            status="In Progress",
            status_code="Input valideret",
        )

        _log_step(step, "Færdig")
    else:
        # Input valideres igen ved genstart, så processen
        # aldrig fortsætter med ugyldige arbejdsdata.
        _validate_box(
            item,
            box,
        )

        _log_step(
            step,
            "Skip state, input genvalideret",
        )

    # ==========================================================
    # TRIN 2: PERSON_DATA_LOADED
    # ==========================================================

    step = "PERSON_DATA_LOADED"
    state = getattr(States, step)

    # Der laves ét opslag i Datafordeleren.
    # Samme person_data genbruges i resten af processen.
    person_data = await asyncio.to_thread(
        _load_person_data,
        box,
    )

    # kan_sendes_brev styrer, om processen må fortsætte.
    # False er et forventet forretningsresultat og skal
    # derfor afsluttes som completed, ikke failed.
    if not person_data.get("kan_sendes_brev"):
        message = _build_not_sent_message(
            person_data
        )

        reason = (
            "BORGER_ER_DOED"
            if person_data.get("er_doed")
            else (
                person_data.get(
                    "kan_sendes_brev_aarsag"
                )
                or "UKENDT_AARSAG"
            )
        )

        _log_step(
            step,
            message,
        )

        _complete_not_sent(
            item,
            message=message,
            reason=reason,
        )

        return

    if _mangler_state(data, state):
        _log_step(step, "Start")

        _set_state(
            item,
            state,
            box_updates={
                "processing_state": "PERSON_DATA_LOADED",
            },
            status="In Progress",
            status_code="Persondata hentet og godkendt",
        )

        _log_step(step, "Færdig")
    else:
        _log_step(
            step,
            "Skip state, persondata genindlæst",
        )

    # ==========================================================
    # TRIN 3: MAIN_DOCUMENT_READY
    # ==========================================================

    step = "MAIN_DOCUMENT_READY"
    state = getattr(States, step)

    # Dokumentbytes gemmes ikke på itemet.
    # Dokumentet gendannes derfor ved genstart.
    main_pdf = await _build_main_pdf(
        box,
        person_data,
    )

    if _mangler_state(data, state):
        _log_step(step, "Start")

        _set_state(
            item,
            state,
            box_updates={
                "processing_state": "MAIN_DOCUMENT_READY",
            },
            status="In Progress",
            status_code="Hoveddokument klargjort",
        )

        _log_step(step, "Færdig")
    else:
        _log_step(
            step,
            "Skip state, dokument gendannet",
        )

    # ==========================================================
    # TRIN 4: ATTACHMENTS_READY
    # ==========================================================

    step = "ATTACHMENTS_READY"
    state = getattr(States, step)

    # Bilagsbytes gemmes ikke på itemet.
    # Bilagene genindlæses derfor ved genstart.
    attachments = await _build_attachments(
        box
    )

    if _mangler_state(data, state):
        _log_step(step, "Start")

        _set_state(
            item,
            state,
            box_updates={
                "processing_state": "ATTACHMENTS_READY",
            },
            status="In Progress",
            status_code="Bilag klargjort",
        )

        _log_step(step, "Færdig")
    else:
        _log_step(
            step,
            "Skip state, bilag genindlæst",
        )

    # ==========================================================
    # TRIN 5: SEND TIL DIGITAL POST-KØ
    # ==========================================================

    step = "SEND_DIGITAL_POST"
    _log_step(step, "Start")

    delivery_method = DeliveryMethod(
        box["delivery_method"]
    )

    address = _build_address(
        delivery_method,
        person_data,
    )

    result = await DigitalPost().send_digital_post(
        pdf_content=main_pdf,
        file_name=box["output_file_name"],
        cpr_or_cvr=box["cpr"],
        subject=box["subject"],
        source_process=SOURCE_PROCESS,
        source_item_id=item.reference,
        delivery_method=delivery_method,
        address=address,
        attachments=attachments,
    )

    result_data = {
        "status": result.status,
        "reason": result.reason,
        "message": result.message,
        "forsendelses_id": result.forsendelses_id,
        "queue_id": result.queue_id,
        "document_count": result.document_count,
    }

    # ==========================================================
    # TRIN 5A: DIGITALPOST_QUEUED
    # ==========================================================

    if result.status == "QUEUED":
        _set_state(
            item,
            States.DIGITALPOST_QUEUED,
            box_updates={
                "processing_state": "DIGITALPOST_QUEUED",
                "digitalpost_result": result_data,
                "last_error": None,
            },
            status="Completed",
            status_code="Lagt i Digital Post-kø",
        )

        item.complete(
            "Brevet er lagt i Digital Post-kø. "
            f"Forsendelses-ID: {result.forsendelses_id}"
        )

        _log_step(
            step,
            "Færdig, brevet er lagt i Digital Post-kø",
        )

        return

    # ==========================================================
    # TRIN 5B: NOT_SENT
    # ==========================================================

    if result.status == "NOT_SENT":
        message = (
            result.message
            or result.reason
            or "Brevet blev ikke sendt"
        )

        _set_state(
            item,
            States.NOT_SENT,
            box_updates={
                "processing_state": "NOT_SENT",
                "digitalpost_result": result_data,
                "last_error": None,
            },
            status="Completed",
            status_code=message,
        )

        item.complete(
            f"Brev ikke sendt: {message}"
        )

        _log_step(
            step,
            f"Færdig uden afsendelse: {message}",
        )

        return

    raise RuntimeError(
        "q-digitalpost returnerede en ukendt status: "
        f"{result.status!r}"
    )