"""Fast teknisk konfiguration for masseforsendelse-af-digitalpost.

Denne fil ændres normalt ikke fra udsendelse til udsendelse.
Alt, som gælder det konkrete brev, Excel-ark og Word-skabelon,
konfigureres i queue_feeder.py og gemmes i item.data["box"].
"""

# Den centrale Digital Post-kø.
# q-digitalpost bruger denne kø til den videre afsendelse.
DIGITAL_POST_QUEUE_ID = 13

# SharePoint-site.
SHAREPOINT_SITE_NAME = "Automatisering"

# Det rigtige dokumentbibliotek.
SHAREPOINT_LIBRARY_NAME = "Dokumenter"

# Mapperne under dokumentbiblioteket.
SHAREPOINT_BASE_FOLDER = (
    "RPA - Processer/"
    "masseforsendelse-af-digitalpost"
)

# Procesnavnet gemmes som source_process i Digital Post-bestillingen.
SOURCE_PROCESS = "masseforsendelse-af-digitalpost"

# None betyder, at alle items, som Automation Server leverer,
# kan behandles i den aktuelle proceskørsel.
MAX_ITEMS_PER_RUN = None
