"""Indgangspunkt for processen send-breve-via-digitalpost."""

from __future__ import annotations

import asyncio
import logging
import os
import sys

from automation_server_client import AutomationServer, Workqueue
from dotenv import load_dotenv

from behandel import PermanentItemError, behandel_item
from configuration import MAX_ITEMS_PER_RUN, WORKQUEUE_ID
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


def _select_items(workqueue: Workqueue):
    items = workqueue
    debug_reference = os.getenv("DEBUG_ITEM_REFERENCE", "").strip()
    if debug_reference:
        return [item for item in items if item.reference == debug_reference][:1]
    if MAX_ITEMS_PER_RUN is None:
        return items
    return list(items)[:MAX_ITEMS_PER_RUN]


async def process_workqueue(workqueue: Workqueue, debug: bool = False) -> int:
    processed = 0
    for item in _select_items(workqueue):
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
                logger.exception("Teknisk fejl for reference=%s", item.reference)
                box = item.data.setdefault("box", {})
                box["last_error"] = f"{type(error).__name__}: {error}"
                item.update(item.data)
                raise
    return processed


async def main() -> None:
    if WORKQUEUE_ID <= 0:
        raise ValueError("WORKQUEUE_ID skal udfyldes i configuration.py.")
    automation_server = AutomationServer.from_environment()
    workqueue = automation_server.workqueue(WORKQUEUE_ID)
    count = await process_workqueue(workqueue)
    logger.info("Behandlede %s item(s).", count)


if __name__ == "__main__":
    asyncio.run(main())
