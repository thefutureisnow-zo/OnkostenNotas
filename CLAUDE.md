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
where `heen/terug` -> `heenenterug` (slash removed).

**Print statements**: use only plain ASCII characters in all `print()` calls.
No arrows (`→`), checkmarks (`✓`), warning signs (`⚠`), or other non-ASCII symbols.
Use `->`, `OK`, `(!!)` etc. instead. This prevents `UnicodeEncodeError` on Windows terminals
that use cp1252 encoding.

## Test fixtures

Tests in `tests/conftest.py` use embedded HTML strings (no external file dependencies).
`temp_excel` fixture builds a minimal Excel file programmatically.
Gmail API calls are mocked in `test_main.py` using `unittest.mock.patch`.

**Important**: The sample HTML in `conftest.py` is a simplified mockup, NOT a faithful copy of the
real NMBS email HTML. When real emails produce parse errors, always inspect the live HTML directly
(see Debugging section below) rather than assuming the test fixtures reflect reality.

## Config

`config.py` is gitignored. Template is `config.example.py`.
Default paths point inside the repo: `data/Onkosten Nota.xlsx` and `screenshots/`.
`data/` and `screenshots/` are also gitignored — never commit expense data.

---

## Debugging parse failures

When `python main.py` shows `Waarschuwing: [ORDER] ... niet gevonden — overgeslagen`, follow
these steps to find and fix the parser:

### 1. Inspect the real email HTML

If you have an `.eml` file, extract and inspect the relevant section:

```python
import email
from bs4 import BeautifulSoup

with open('path/to/EMAIL.eml', 'rb') as f:
    msg = email.message_from_bytes(f.read())

html = next(
    part.get_payload(decode=True).decode(part.get_content_charset() or 'utf-8', errors='replace')
    for part in msg.walk()
    if part.get_content_type() == 'text/html'
)
soup = BeautifulSoup(html, 'lxml')
```

Then probe what the parser is targeting. Examples:

```python
# Check what TDs contain a label (e.g. Totaalbedrag)
for td in soup.find_all('td'):
    t = td.get_text(strip=True)
    if 'Totaalbedrag' in t and len(t) < 200:
        print(len(t), repr(t), '| sib:', td.find_next_sibling('td'))

# Check what spans contain "Van :"
for s in soup.find_all('span'):
    if 'Van' in s.get_text():
        print(repr(s))
```

### 2. Find where the real structure differs

Key lessons learned from past debugging:

- **Wrapper TDs**: NMBS emails often have outer `<td>` containers whose `get_text()` includes ALL
  nested content. A `find_all("td")` loop can match these FIRST before reaching the specific row.
  `find_next_sibling("td")` then returns `None` because the outer wrapper has no siblings.
- **Prefer `full_text` regex over DOM navigation** for values that appear as plain text (prices,
  dates, order numbers). `full_text = soup.get_text(" ", strip=True)` is already computed in
  `parse_nmbs_email` and is more resilient to HTML restructuring.
- **Greedy regex order**: if `re.search` finds the first match and there are multiple numbers
  (e.g. individual ticket prices before the total), anchor the regex to the label:
  `re.search(r"Totaalbedrag\s*:?[^\d]*([\d]+[,.][\d]+)", full_text)`.
- **Station spans**: `soup.find("span", string=re.compile(r"Van\s*:"))` requires the span's
  full text to match. If NMBS changes the label spacing or language, this breaks.

### 3. Verify the fix against a real email before running tests

```bash
python3 -c "
import email
from email_parser import parse_nmbs_email
with open('path/to/EMAIL.eml', 'rb') as f:
    msg = email.message_from_bytes(f.read())
html = next(p.get_payload(decode=True).decode(p.get_content_charset() or 'utf-8', errors='replace')
            for p in msg.walk() if p.get_content_type() == 'text/html')
t = parse_nmbs_email(html)
print(t.order_number, t.from_station, t.to_station, t.direction, t.travel_date, t.price)
"
```

### 4. Run the full test suite

```bash
python3 -m pytest -v
```

All 47 tests must pass before committing. If a test expectation is wrong (not the code), fix the
test — but verify against the real email first to be sure.

---

## Issue logging

Every bug fix or non-trivial feature must have a log file at:

```
issues/<issue-name>/<issue-name>.md
```

The file must cover four sections:

1. **The problem** — what the user observed and what the actual error was
2. **How we spotted it** — the quickest path to finding the root cause (commands run, tracebacks seen)
3. **Root cause** — the underlying reason, not just the symptom
4. **Fix** — what was changed and why, with the PR link

Keep it concise. The goal is that you or the user can re-read it months later and immediately understand
what happened and why the fix works.

---

## Git workflow: branches + pull requests

**All fixes and new features must be developed on a separate branch and merged via pull request.**
Never commit directly to `main`.

### Creating a branch

```bash
git checkout -b fix/description-of-fix   # bug fix
git checkout -b feat/description          # new feature
```

### After making changes

1. Run `python3 -m pytest -v` — all tests must pass.
2. Commit on the branch.
3. Push and open a PR:

```bash
git push -u origin <branch-name>
gh pr create --title "..." --body "..."
```

### PR validation (sub-agent)

After opening a PR, always launch a validation sub-agent with the following instructions:

> "Review the open PR at [PR URL]. Check out the branch, run `python3 -m pytest -v`, confirm all
> tests pass, read the changed files, and report: (1) test results, (2) any logic issues or edge
> cases not covered by the tests, (3) whether the PR is safe to merge."

Only merge after the sub-agent confirms tests pass and raises no blocking issues.

### Merge

```bash
gh pr merge <PR-number> --squash --delete-branch
```
