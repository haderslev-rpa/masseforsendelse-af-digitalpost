# masseforsendelse-af-digitalpost

Automation Server-proces, som genererer ét brev pr. work item og afleverer PDF'en til den centrale Digital Post-kø.

## Procesflow

1. `queue_feeder.py` læser CPR-numre fra `tests_local/cpr.xlsx` og opretter ét source-item pr. CPR.
2. Worker-processen henter navn og adresse fra Datafordeleren, når funktionen er aktiveret.
3. Den fælles DOCX-skabelon hentes fra SharePoint.
4. `q-word` udfylder `[[PLADSHOLDERE]]`.
5. `q-gotenberg` konverterer DOCX-bytes til PDF-bytes.
6. `q-digitalpost` uploader PDF'en og opretter et item i den centrale Digital Post-kø.
7. Source-itemet completes, når resultatet er `QUEUED` eller `NOT_SENT`.

## Første opsætning

1. Opret køen `masseforsendelse-af-digitalpost` i Automation Server.
2. Sæt køens tekniske ID i `configuration.py` som `WORKQUEUE_ID`.
3. Upload Word-skabelonen til SharePoint.
4. Sæt `TEMPLATE_SITE_NAME`, `TEMPLATE_LIBRARY_NAME` og `TEMPLATE_DRIVE_ITEM_ID` i `configuration.py`.
5. Kør `uv lock` og `uv sync`.
6. Kontrollér `.env`.

## Word-skabelon

`q-word` bruger pladsholdere med dobbelte kantede parenteser, eksempelvis:

- `[[NAVN]]`
- `[[ADRESSE]]`
- `[[POSTNUMMER_OG_BY]]`
- `[[DATO]]`

Når `FILL_NAME_AND_ADDRESS_BOOKMARKS = True`, udfyldes de tre første automatisk via Datafordeleren. Øvrige pladsholdere angives i `BOOKMARK_VALUES` øverst i `queue_feeder.py`.

## Manuel feeder

Læg Excel-filen i:

`tests_local/cpr.xlsx`

Standard er CPR i kolonne A fra række 2. Ret derefter den manuelle konfiguration øverst i `queue_feeder.py`, blandt andet:

- `BATCH_ID`
- `SUBJECT`
- `BOOKMARK_VALUES`
- `FILL_NAME_AND_ADDRESS_BOOKMARKS`
- `DELIVERY_METHOD`

Kør først med `DRY_RUN = True`:

`uv run python queue_feeder.py`

Når optællingen er kontrolleret, ret til `DRY_RUN = False` og kør igen.

## Worker

Kør processen:

`uv run python main.py`

## Dubletbeskyttelse

Referencen består af:

`<BATCH_ID>:<løbenummer>:<CPR>`

Inden et brev oprettes igen, søger worker-processen efter `source_item_id` i den centrale Digital Post-kø. Hvis itemet allerede findes, bliver source-itemet completed uden at oprette en ny bestilling.

Vælg altid et nyt og meningsfuldt `BATCH_ID` for hver udsendelse.
