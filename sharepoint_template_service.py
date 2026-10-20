"""SharePoint-hjælpere til masseforsendelsesprocessen.

Feederen finder filer ud fra site, bibliotek og sti. Når filen er fundet,
gemmes kun drive_id og drive_item_id på work itemet. Worker-processen kan
derefter hente filen direkte gennem Microsoft Graph.
"""

from __future__ import annotations

from typing import Any
from urllib.parse import quote

import requests

from q_sharepoint_api.sp_api import get_client


def build_relative_path(
    *,
    base_folder: str,
    template_folder: str,
    file_name: str,
) -> str:
    """Byg en sikker relativ SharePoint-sti og returnér den som tekst."""
    parts = [base_folder, template_folder, file_name]
    cleaned_parts: list[str] = []

    for part in parts:
        if not isinstance(part, str) or not part.strip():
            raise ValueError("Alle dele af SharePoint-stien skal udfyldes.")
        cleaned_parts.extend(
            value.strip()
            for value in part.replace("\\", "/").split("/")
            if value.strip()
        )

    return "/".join(cleaned_parts)


def _get_document_library(
    *,
    client: Any,
    site_id: str,
    library_name: str,
) -> dict[str, str]:
    """Find dokumentbibliotekets list-id og drive-id."""
    response = requests.get(
        f"{client.base}/sites/{site_id}/lists",
        headers=client.auth.graph_headers(),
        timeout=30,
    )
    response.raise_for_status()

    matches = [
        item
        for item in response.json().get("value", [])
        if item.get("displayName", "").casefold()
        == library_name.casefold()
    ]

    if len(matches) != 1:
        raise RuntimeError(
            f"Forventede præcis ét dokumentbibliotek med navnet "
            f"{library_name!r}, men fandt {len(matches)}."
        )

    list_id = matches[0]["id"]
    drive_response = requests.get(
        f"{client.base}/sites/{site_id}/lists/{list_id}/drive",
        headers=client.auth.graph_headers(),
        timeout=30,
    )
    drive_response.raise_for_status()
    drive_id = drive_response.json().get("id")

    if not drive_id:
        raise RuntimeError(
            f"Dokumentbiblioteket {library_name!r} returnerede intet drive-id."
        )

    return {"list_id": list_id, "drive_id": drive_id}


def get_file_metadata_by_path(
    *,
    site_name: str,
    library_name: str,
    relative_path: str,
) -> dict[str, Any]:
    """Find én SharePoint-fil og returnér de nødvendige tekniske metadata."""
    client = get_client()
    site_id = client.get_site_id(site_name)
    library = _get_document_library(
        client=client,
        site_id=site_id,
        library_name=library_name,
    )
    drive_id = library["drive_id"]
    encoded_path = quote(relative_path.strip("/"), safe="/")

    response = requests.get(
        f"{client.base}/drives/{drive_id}/root:/{encoded_path}",
        headers=client.auth.graph_headers(),
        timeout=30,
    )

    if response.status_code == 404:
        raise FileNotFoundError(
            f"SharePoint-filen blev ikke fundet: {relative_path}"
        )

    response.raise_for_status()
    item = response.json()

    if "file" not in item:
        raise ValueError(
            f"SharePoint-stien peger ikke på en fil: {relative_path}"
        )

    drive_item_id = item.get("id")
    if not drive_item_id:
        raise RuntimeError("SharePoint returnerede intet drive-item-id.")

    return {
        "drive_id": drive_id,
        "drive_item_id": drive_item_id,
        "file_name": item.get("name"),
    }


def download_file_by_drive_ids(
    *,
    drive_id: str,
    drive_item_id: str,
) -> bytes:
    """Download én fil direkte med drive-id og drive-item-id.

    Output:
        Filens binære indhold som bytes.
    """
    if not isinstance(drive_id, str) or not drive_id.strip():
        raise ValueError("drive_id mangler.")
    if not isinstance(drive_item_id, str) or not drive_item_id.strip():
        raise ValueError("drive_item_id mangler.")

    client = get_client()
    response = requests.get(
        f"{client.base}/drives/{drive_id.strip()}"
        f"/items/{drive_item_id.strip()}/content",
        headers=client.auth.graph_headers(),
        timeout=120,
        allow_redirects=True,
    )
    response.raise_for_status()

    if not response.content:
        raise RuntimeError("SharePoint returnerede en tom fil.")

    return response.content
