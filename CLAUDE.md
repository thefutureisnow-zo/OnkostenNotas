# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this project does

Python CLI that reads NMBS (Belgian railway) ticket confirmation emails from Gmail via OAuth2,
generates PNG screenshots, and writes rows to an Excel "onkostennota" expense file.
Runs interactively — prompts the user to confirm each ticket before writing.

## Commands

```bash
# Install dependencies
pip install -r requirements.txt

# First-time setup: copy config template and fill in paths
copy config.example.py config.py

# Run the tool
python main.py

# Run tests
pytest
pytest tests/test_email_parser.py   # single module
pytest -v                           # verbose
```

## Architecture

```
main.py              Orchestrator: fetch → parse → prompt → screenshot → Excel
gmail_client.py      Gmail API OAuth2; returns (msg_id, order_number, html) tuples
email_parser.py      BeautifulSoup HTML parser → TicketData dataclass
excel_updater.py     openpyxl: find/create Dutch month sheet, append row
screenshot_gen.py    html2image (headless Chrome) → PNG per ticket
holidays_be.py       Belgian holiday + weekend detection (holidays library)
state.py             processed.json — dedup by NMBS order number
config.py            Local paths (gitignored); copy from config.example.py
```

## Key conventions

**Excel sheet names** use Dutch month names: `"Januari 2026"`, `"Februari 2026"`, etc.
The `DUTCH_MONTHS` dict is defined in both `excel_updater.py` and `screenshot_gen.py`.

**Excel structure** (same every sheet):
- Row 7: headers; rows 8–15: data; row 16+: SUM summary formulas
- Col A=Datum (Excel serial int), B=Nr, C=Omschrijving, D=Curr, F=Vervoer, L=Tot.EUR (formula)
- Overflow beyond 8 rows: `_insert_overflow_row()` inserts before row 16 and patches SUM ranges

**Email direction logic** (one email = one Excel row):
- `"Heen en terug"` → direction `"heen/terug"`, uses Heen date
- `"Enkel"` + only Heen date → direction `"heen"`
- `"Enkel"` + only Terug date → direction `"terug"`

**State/dedup**: `state.py` tracks two sets in `processed.json`:
- `processed`: added to Excel (never shown again)
- `skipped_weekend`: user declined a weekend/holiday ticket (never shown again)
- Tickets where user answers "n" to the normal prompt are NOT persisted (shown again next run)

**Screenshot filenames**: `trein_{DDMMYY}_{direction_slug}.png`
where `heen/terug` → `heenenterug` (slash removed).

## Test fixtures

Tests in `tests/conftest.py` use embedded HTML strings (no external file dependencies).
`temp_excel` fixture builds a minimal Excel file programmatically.
Gmail API calls are mocked in `test_main.py` using `unittest.mock.patch`.

## Config

`config.py` is gitignored. Template is `config.example.py`.
Default paths point inside the repo: `data/Onkosten Nota.xlsx` and `screenshots/`.
`data/` and `screenshots/` are also gitignored — never commit expense data.
