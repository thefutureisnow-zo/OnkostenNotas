# Excel Structure

## Sheet naming

Sheet names use Dutch month names: `"Januari 2026"`, `"Februari 2026"`, etc.
The `DUTCH_MONTHS` dict is defined in `constants.py` and imported by `excel_updater.py`
and `screenshot_gen.py`.

## Row/column layout (same every sheet)

- **Row 7**: headers
- **Rows 8-15**: data (standard 8 rows)
- **Row 16+**: SUM summary formulas

### Column mapping (1-based)

| Column | Letter | Content |
|--------|--------|---------|
| 1 | A | Datum (Excel serial int) |
| 2 | B | Nr (sequence number) |
| 3 | C | Omschrijving (e.g. "Trein Zottegem - Antwerpen-Zuid heen") |
| 4 | D | Curr ("EUR") |
| 6 | F | Vervoer (ticket price) |
| 12 | L | Tot.EUR (`=SUM(E:K)` formula) |

## Overflow handling

When all 8 standard data rows are full, `_insert_overflow_row()` inserts a new row
before the SUM summary block and patches the SUM ranges to include the new row.

`_find_next_data_row()` scans beyond the standard range to handle sheets that already
have overflow rows, stopping at the first empty row or SUM formula boundary.

## New sheet creation

`_create_month_sheet()` copies the most recent sheet as a template, clears data rows
(columns A-K), updates the month name in B5, and sets the date range in K4/K5.

## Date conversion

`date_to_excel_serial()` converts Python `date` to Excel serial number, accounting for
the Excel 1900 leap year bug (epoch = 1899-12-30).
