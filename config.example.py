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

# Pad naar het Excel onkostennota-bestand
EXCEL_PATH = BASE_DIR / "data" / "Onkosten Nota.xlsx"

# Map waar de screenshots per maand worden opgeslagen.
# Subfolders (bijv. "Februari 2026") worden automatisch aangemaakt.
SCREENSHOTS_DIR = BASE_DIR / "screenshots"

# ---------------------------------------------------------------
# Niet aanpassen — automatisch ingesteld
# ---------------------------------------------------------------

CREDENTIALS_DIR = BASE_DIR / "credentials"
TOKEN_PATH = CREDENTIALS_DIR / "token.json"
CLIENT_SECRET_PATH = CREDENTIALS_DIR / "client_secret.json"
STATE_FILE = BASE_DIR / "processed.json"
