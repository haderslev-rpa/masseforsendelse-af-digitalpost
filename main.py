"""Indgangspunkt for processen masseforsendelse-af-digitalpost."""

from __future__ import annotations

import argparse
import asyncio
import logging
import os
import sys

from automation_server_client import AutomationServer, Workqueue
from dotenv import load_dotenv

from behandel import PermanentItemError, behandel_item
from configuration import MAX_ITEMS_PER_RUN
from models import ProcessingState

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    stream=sys.stdout,
    force=True,
)
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("automation_server_client").setLevel(logging.WARNING)
logger = logging.getLogger(__name__)


def _vaelg_items_til_behandling(
    workqueue: Workqueue,
    *,
    debug: bool,
):
    """Vælg køens items eller ét bestemt item ved lokal debug."""
    if debug:
        debug_reference = os.getenv("DEBUG_ITEM_REFERENCE", "").strip()
        if not debug_reference:
            raise ValueError(
                "DEBUG_ITEM_REFERENCE skal udfyldes i .env ved --debug."
            )
        return [
            item
            for item in workqueue
            if item.reference == debug_reference
        ][:1]

    if MAX_ITEMS_PER_RUN is None:
        return workqueue
    return list(workqueue)[:MAX_ITEMS_PER_RUN]


async def process_workqueue(
    workqueue: Workqueue,
    *,
    debug: bool,
) -> int:
    """Behandl work items ét ad gangen."""
    processed = 0
    for item in _vaelg_items_til_behandling(workqueue, debug=debug):
        with item:
            processed += 1
            logger.info("Behandler brev-item reference=%s", item.reference)
            try:
                await behandel_item(item)
            except PermanentItemError as error:
                logger.error("Permanent itemfejl: %s", error)
                box = item.data.setdefault("box", {})
                box["processing_state"] = ProcessingState.FAILED
                box["last_error"] = str(error)
                item.update(item.data)
                item.fail(str(error))
            except Exception as error:
                logger.exception(
                    "Teknisk fejl for reference=%s",
                    item.reference,
                )
                box = item.data.setdefault("box", {})
                box["last_error"] = f"{type(error).__name__}: {error}"
                item.update(item.data)
                raise
    return processed


def _parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Behandl kun DEBUG_ITEM_REFERENCE ved lokal fejlsøgning.",
    )
    return parser.parse_args()


async def main() -> None:
    arguments = _parse_arguments()
    automation_server = AutomationServer.from_environment()

    # Køen kommer fra Automation Server-processens miljø.
    # Der bruges derfor ikke et hardkodet WORKQUEUE_ID.
    workqueue = automation_server.workqueue()

    count = await process_workqueue(
        workqueue,
        debug=arguments.debug,
    )
    logger.info("Behandlede %s item(s).", count)


if __name__ == "__main__":
    asyncio.run(main())
