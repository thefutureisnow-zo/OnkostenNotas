# CLAUDE.md

## What this project does

Python CLI that reads NMBS (Belgian railway) ticket confirmation emails from Gmail via OAuth2,
generates PNG screenshots, and writes rows to an Excel "onkostennota" expense file.

## Commands

```bash
pip install -r requirements.txt       # install dependencies
copy config.example.py config.py      # first-time setup
python main.py                        # run the tool
python main.py --reset                # reset processed tickets
pytest -v                             # run tests
```

## Architecture

```
main.py              Orchestrator: fetch -> parse -> prompt -> screenshot -> Excel
gmail_client.py      Gmail API OAuth2; returns (msg_id, order_number, html) tuples
email_parser.py      BeautifulSoup HTML parser -> TicketData dataclass
excel_updater.py     openpyxl: find/create Dutch month sheet, append row
screenshot_gen.py    html2image (headless Chrome) -> PNG per ticket
holidays_be.py       Belgian holiday + weekend detection (holidays library)
state.py             processed.json -- dedup by NMBS order number
constants.py         Shared constants (DUTCH_MONTHS)
config.py            Local paths (gitignored); copy from config.example.py
```

## Key conventions

- **Direction logic**: "Heen en terug" -> "heen/terug" (uses Heen date); "Enkel" -> "heen" or "terug" based on which date is present. Station-based inference via `home_station`/`office_station` params refines single-trip direction.
- **Screenshot filenames**: `trein_{DDMMYY}_{direction_slug}_{order_number}.png` (slash removed from heen/terug)
- **ASCII print only**: No unicode symbols in `print()` calls (prevents cp1252 errors on Windows). Use `->`, `OK`, `(!!)` instead.
- **State/dedup**: `processed.json` tracks `processed` (added to Excel) and `skipped_weekend` (declined). Tickets declined via normal prompt are NOT persisted.
- **Config**: `config.py` is gitignored. Template: `config.example.py`. `data/` and `screenshots/` are also gitignored.

## Test fixtures

Tests use embedded HTML strings in `tests/conftest.py` (no external files).
The sample HTML is a simplified mockup, NOT a faithful copy of real NMBS emails.
When debugging real parse errors, inspect the live HTML directly (see `docs/DEBUGGING.md`).

## Documentation

- [docs/DEBUGGING.md](docs/DEBUGGING.md) -- debugging parse failures step by step
- [docs/EXCEL-STRUCTURE.md](docs/EXCEL-STRUCTURE.md) -- Excel row/column layout and overflow logic
- [docs/GIT-WORKFLOW.md](docs/GIT-WORKFLOW.md) -- branch naming, PR workflow, merge strategy
- [docs/ISSUE-LOGGING.md](docs/ISSUE-LOGGING.md) -- issue log format and file path convention
