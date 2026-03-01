"""
Configuratie-template — kopieer dit bestand naar config.py.

    copy config.example.py config.py

config.py wordt NIET naar Git gepusht (staat in .gitignore).
De mappen data/ en screenshots/ ook niet.
"""
from pathlib import Path

# Basis-map van het project (automatisch bepaald)
BASE_DIR = Path(__file__).parent

# ---------------------------------------------------------------
# Pas aan indien gewenst (standaard werkt alles in de projectmap)
# ---------------------------------------------------------------

# Map waar de per-maand Excel-bestanden worden opgeslagen.
# Elk maandbestand (bijv. Onkosten_Januari_2026.xlsx) wordt automatisch aangemaakt.
EXCEL_DIR = BASE_DIR / "data"

# Map waar de screenshots per maand worden opgeslagen.
# Subfolders (bijv. "Februari 2026") worden automatisch aangemaakt.
SCREENSHOTS_DIR = BASE_DIR / "screenshots"

# Map waar HTML-rapporten worden opgeslagen (optioneel).
# Verwijder de regel of zet op None om HTML-rapporten uit te schakelen.
REPORTS_DIR = BASE_DIR / "reports"

# Thuisstation en kantoorstation — gebruikt om heen/terug-richting te bepalen
# bij enkelvoudige tickets (Enkel). Gebruik dezelfde schrijfwijze als NMBS
# (title-case, koppelteken waar van toepassing, bijv. "Antwerpen-Zuid").
HOME_STATION = "Zottegem"
OFFICE_STATION = "Antwerpen-Zuid"

# ---------------------------------------------------------------
# Niet aanpassen — automatisch ingesteld
# ---------------------------------------------------------------

CREDENTIALS_DIR = BASE_DIR / "credentials"
TOKEN_PATH = CREDENTIALS_DIR / "token.json"
CLIENT_SECRET_PATH = CREDENTIALS_DIR / "client_secret.json"
STATE_FILE = BASE_DIR / "processed.json"
