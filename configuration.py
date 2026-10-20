"""Fast konfiguration til processen send-breve-via-digitalpost."""

# Det tekniske ID på processens egen Automation Server-kø.
# Ret værdien efter køen er oprettet i Automation Server.
WORKQUEUE_ID = 23

# Det tekniske ID på den centrale Digital Post-kø.
DIGITAL_POST_QUEUE_ID = 13

# SharePoint-placering for den fælles Word-skabelon.
TEMPLATE_SITE_NAME = "Automatisering"
TEMPLATE_LIBRARY_NAME = "RPA - Processer/masseforsendelse-af-digitalpost"

# Indsæt SharePoints drive-item-id for DOCX-skabelonen.
# Drive-item-id bruges, så alle items i et batch anvender samme filreference.
TEMPLATE_DRIVE_ITEM_ID = ""
TEMPLATE_FILE_NAME = "brevskabelon.docx"

# Worker-capability i Automation Server.
WORKER_CAPABILITY = "masseforsendelse-af-digitalpost"

# Procesnavnet gemmes i Digital Post-itemets source_process.
SOURCE_PROCESS = "masseforsendelse-af-digitalpost"

# Maksimalt antal NEW-items, der behandles pr. proceskørsel.
# None betyder alle items, som Automation Server leverer til workqueue-objektet.
MAX_ITEMS_PER_RUN = None
