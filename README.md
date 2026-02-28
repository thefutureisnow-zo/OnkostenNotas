# NMBS Onkostennota

Dit hulpprogramma controleert automatisch je Gmail op NMBS-ticketbevestigingen,
slaat een screenshot van elk ticket op en vult de juiste rij in jouw
onkostennota-Excelbestand in.

---

## Wat doet dit programma?

Elke keer dat je het uitvoert:

1. **Haalt NMBS-ticketmails op** uit je Gmail (de laatste 2 jaar).
2. **Toont elk nieuw ticket** op het scherm en vraagt of je het wilt toevoegen.
3. **Slaat een PNG-screenshot** op in de juiste maandmap (bijv. `screenshots/Februari 2026/`).
4. **Vult de juiste rij in** het Excel-onkostennota-bestand in voor die maand.
5. **Onthoudt welke tickets al verwerkt zijn**, zodat je nooit dubbele rijen krijgt.

Tickets op weekend- of feestdagen worden apart geflagd: het programma vraagt dan
of je ze toch wilt opnemen.

---

## Stap 1 — Python installeren

1. Ga naar [python.org/downloads](https://www.python.org/downloads/) en download
   de nieuwste versie voor Windows (bijv. **Python 3.12**).
2. Start het installatieprogramma.
   > ⚠️ **Belangrijk:** vink **"Add Python to PATH"** aan voor je op *Install Now* klikt.
   > Dit is de meest gemaakte fout — vergeet dit niet!
3. Controleer na de installatie of het gelukt is:
   - Open **Opdrachtprompt** (typ `cmd` in het Startmenu).
   - Typ: `python --version`
   - Je zou zoiets als `Python 3.12.x` moeten zien.

---

## Stap 2 — Dit project downloaden

**Als je Git hebt:**
```
git clone https://github.com/thefutureisnow-zo/OnkostenNotas.git
cd OnkostenNotas
```

**Als je geen Git hebt:**
1. Ga naar de GitHub-pagina van het project.
2. Klik op de groene knop **"Code"** → **"Download ZIP"**.
3. Pak het ZIP-bestand uit naar een map die je makkelijk terugvindt,
   bijv. `C:\repos\OnkostenNotas`.

---

## Stap 3 — Configuratie instellen

1. Ga naar de projectmap (bijv. `C:\repos\OnkostenNotas`).
2. Kopieer `config.example.py` naar `config.py`:
   ```
   copy config.example.py config.py
   ```
3. Open `config.py` met **Kladblok** (rechtsklik → *Openen met* → *Kladblok*).
4. Pas de twee paden aan:
   - `EXCEL_PATH`: het volledige pad naar jouw `Onkosten Nota.xlsx`.
   - `SCREENSHOTS_DIR`: de map waar screenshots per maand worden opgeslagen.

   Standaard staan ze ingesteld op mappen *binnen* de projectmap (`data/` en `screenshots/`),
   wat prima werkt. Als je het bestand elders wilt hebben, pas dan het pad aan.

   > 💡 **Tip:** in Python-paden kun je gewone schuine strepen `/` gebruiken in
   > plaats van dubbele backslashes `\\`.

5. Maak de map `data/` aan en zet je `Onkosten Nota.xlsx` daarin:
   ```
   mkdir data
   copy "C:\pad\naar\Onkosten Nota.xlsx" data\
   ```

---

## Stap 4 — Google Gmail-toegang instellen

Dit is de meest uitgebreide stap, maar je hoeft het maar **één keer** te doen.

### 4a. Google Cloud-project aanmaken

1. Ga naar [console.cloud.google.com](https://console.cloud.google.com).
2. Log in met het **Google-account waarop je de NMBS-mails ontvangt**.
3. Klik linksboven op de projecten-dropdown → **"Nieuw project"**.
4. Geef het een naam (bijv. `NMBS Tickets`) en klik op **Maken**.

### 4b. Gmail API inschakelen

1. Gebruik de zoekbalk bovenaan en zoek naar **"Gmail API"**.
2. Klik op het resultaat → klik op de blauwe knop **"Inschakelen"**.

### 4c. OAuth-inloggegevens aanmaken

1. Ga in het linkermenu naar **"API's en services"** → **"Inloggegevens"**.
2. Klik op **"+ Inloggegevens maken"** → **"OAuth-client-ID"**.
3. Als je gevraagd wordt een *toestemmingsscherm* in te stellen:
   - Kies **"Extern"** → **Maken**.
   - Vul een app-naam in (bijv. `NMBS Onkostennota`) en jouw e-mailadres.
   - Klik op **Opslaan en doorgaan** (je kunt de rest overslaan).
4. Terug bij het aanmaken van de client-ID:
   - **Type toepassing:** kies **"Desktopapp"**.
   - Geef het een naam (bijv. `Onkostennota app`) → klik op **Maken**.
5. Er verschijnt een venster met je inloggegevens. Klik op **"JSON downloaden"**.
6. Hernoem het gedownloade bestand naar **`client_secret.json`**.
7. Maak de map `credentials` aan in de projectmap en zet het bestand daarin:
   ```
   mkdir credentials
   copy C:\Users\JOUWNAAM\Downloads\client_secret.json credentials\
   ```

### 4d. Testgebruiker toevoegen (eenmalig)

Zolang je app niet geverifieerd is door Google, moet je jezelf als testgebruiker toevoegen:

1. Ga naar **"OAuth-toestemmingsscherm"** in het linkermenu.
2. Scroll naar **"Testgebruikers"** → klik op **"+ Gebruikers toevoegen"**.
3. Vul je Gmail-adres in en sla op.

---

## Stap 5 — Afhankelijkheden installeren

1. Open **Opdrachtprompt** in de projectmap.
   > 💡 Tip: navigeer in Verkenner naar de projectmap, klik in de adresbalk
   > en typ `cmd`, dan druk op Enter.
2. Voer uit:
   ```
   pip install -r requirements.txt
   ```
3. Wacht tot alles klaar is (veel tekst die scrolt — dat is normaal).

---

## Stap 6 — Eerste keer uitvoeren

1. Zorg dat je `credentials\client_secret.json` aanwezig is (zie Stap 4).
2. Voer het programma uit:
   ```
   python main.py
   ```
3. **Eerste keer:** er opent een browservenster waar je wordt gevraagd in te loggen
   bij Google en het programma toestemming te geven om je Gmail te lezen.
   Klik op **Toestaan**.
4. Daarna slaat het programma een `token.json` op in de `credentials/`-map,
   zodat je de volgende keer niet opnieuw hoeft in te loggen.
5. Het programma toont elk gevonden ticket en vraagt of je het wilt toevoegen.

---

## Dagelijks gebruik

Voer gewoon uit:

```
python main.py
```

Het programma toont alleen **nieuwe** tickets die nog niet verwerkt zijn.
Als er niets nieuws is, meldt het dat meteen.

### Bij een weekend- of feestdagticket

```
⚠  Dit ticket is gekocht op een zaterdag.
      Toch opnemen in de onkostennota? [j/N]:
```

- **n** (of Enter): permanent overgeslagen, verschijnt nooit meer.
- **j**: wordt toch toegevoegd.

### Bij een normaal ticket

```
      Toevoegen aan de onkostennota? [J/n]:
```

- **J** (of Enter): screenshot opslaan + rij toevoegen aan Excel.
- **n**: overgeslagen voor nu, verschijnt de volgende keer opnieuw.

---

## Problemen oplossen

| Foutmelding | Oplossing |
|---|---|
| `'python' is not recognized` | Python staat niet in PATH. Herinstalleer Python en vink "Add Python to PATH" aan. |
| `ModuleNotFoundError` | Voer `pip install -r requirements.txt` opnieuw uit vanuit de projectmap. |
| `client_secret.json niet gevonden` | Zorg dat het bestand in `credentials\client_secret.json` staat (zie Stap 4). |
| Browser opent niet bij eerste login | Verwijder `credentials\token.json` en voer opnieuw uit. |
| Excel-bestand vergrendeld | Sluit het bestand eerst in Excel voordat je het programma uitvoert. |
| Screenshot mislukt | Controleer of **Google Chrome** geïnstalleerd is. |
| `config.py niet gevonden` | Voer `copy config.example.py config.py` uit en pas de paden aan. |

---

## Bestanden die niet naar GitHub worden gepusht

De volgende bestanden/mappen staan in `.gitignore` en worden nooit gedeeld:

- `config.py` — jouw lokale paden
- `credentials/` — Google-loginbestanden
- `processed.json` — lijst van verwerkte tickets
- `data/` — jouw Excel-bestand
- `screenshots/` — de opgeslagen ticketscreenshots
