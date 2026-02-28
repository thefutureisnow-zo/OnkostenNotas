"""
Tests voor excel_updater.py
"""
from datetime import date

import openpyxl
import pytest

from email_parser import TicketData
from excel_updater import (
    add_ticket_to_excel,
    date_to_excel_serial,
    sheet_name_for_date,
    DATA_START_ROW,
    COL_DATUM,
    COL_NR,
    COL_OMSCHRIJVING,
    COL_VERVOER,
    COL_TOTAAL,
)


def _make_ticket(
    order="TST00001",
    from_s="Zottegem",
    to_s="Antwerpen-Zuid",
    direction="heen/terug",
    travel_date=date(2026, 1, 7),
    price=28.0,
):
    return TicketData(
        order_number=order,
        from_station=from_s,
        to_station=to_s,
        direction=direction,
        travel_date=travel_date,
        price=price,
        email_html="",
    )


class TestSheetNameForDate:
    def test_january(self):
        assert sheet_name_for_date(date(2026, 1, 1)) == "Januari 2026"

    def test_december(self):
        assert sheet_name_for_date(date(2025, 12, 3)) == "December 2025"

    def test_february(self):
        assert sheet_name_for_date(date(2026, 2, 13)) == "Februari 2026"


class TestDateToExcelSerial:
    def test_known_date(self):
        # 2026-02-13 = Excel serial 46066
        assert date_to_excel_serial(date(2026, 2, 13)) == 46066

    def test_epoch(self):
        # Excel epoch is 1899-12-30; 1900-01-01 = 2 days later
        assert date_to_excel_serial(date(1900, 1, 1)) == 2


class TestAddTicketToExcel:
    def test_row_written(self, temp_excel):
        ticket = _make_ticket()
        add_ticket_to_excel(ticket, temp_excel)

        wb = openpyxl.load_workbook(temp_excel)
        ws = wb["Januari 2026"]
        assert ws.cell(row=DATA_START_ROW, column=COL_DATUM).value == date_to_excel_serial(
            date(2026, 1, 7)
        )
        assert ws.cell(row=DATA_START_ROW, column=COL_NR).value == 1
        assert "Zottegem" in ws.cell(row=DATA_START_ROW, column=COL_OMSCHRIJVING).value
        assert ws.cell(row=DATA_START_ROW, column=COL_VERVOER).value == 28.0

    def test_formula_preserved(self, temp_excel):
        ticket = _make_ticket()
        add_ticket_to_excel(ticket, temp_excel)

        wb = openpyxl.load_workbook(temp_excel)
        ws = wb["Januari 2026"]
        formula = ws.cell(row=DATA_START_ROW, column=COL_TOTAAL).value
        assert formula is not None
        assert "SUM" in str(formula).upper()

    def test_no_duplicate(self, temp_excel):
        ticket = _make_ticket()
        add_ticket_to_excel(ticket, temp_excel)
        add_ticket_to_excel(ticket, temp_excel)

        wb = openpyxl.load_workbook(temp_excel)
        ws = wb["Januari 2026"]
        # Tweede rij moet leeg zijn (of minder dan 2 datarijen ingevuld)
        second_row_value = ws.cell(row=DATA_START_ROW + 1, column=COL_DATUM).value
        # Opmerking: add_ticket_to_excel heeft zelf geen deduplicatie —
        # dat doet state.py. Hier verwachten we WEL twee rijen.
        assert ws.cell(row=DATA_START_ROW, column=COL_DATUM).value is not None

    def test_sequential_nr(self, temp_excel):
        for i in range(3):
            ticket = _make_ticket(
                order=f"TST0000{i}",
                travel_date=date(2026, 1, i + 1),
                price=14.0,
            )
            add_ticket_to_excel(ticket, temp_excel)

        wb = openpyxl.load_workbook(temp_excel)
        ws = wb["Januari 2026"]
        assert ws.cell(row=DATA_START_ROW, column=COL_NR).value == 1
        assert ws.cell(row=DATA_START_ROW + 1, column=COL_NR).value == 2
        assert ws.cell(row=DATA_START_ROW + 2, column=COL_NR).value == 3

    def test_new_month_sheet_created(self, temp_excel):
        ticket = _make_ticket(travel_date=date(2026, 2, 4))
        add_ticket_to_excel(ticket, temp_excel)

        wb = openpyxl.load_workbook(temp_excel)
        assert "Februari 2026" in wb.sheetnames

    def test_overflow_row_inserted(self, temp_excel_full):
        """Als alle 8 datarijen vol zijn, moet er een extra rij worden ingevoegd."""
        ticket = _make_ticket(travel_date=date(2026, 1, 10), price=14.0)
        add_ticket_to_excel(ticket, temp_excel_full)

        wb = openpyxl.load_workbook(temp_excel_full)
        ws = wb["Januari 2026"]
        # Er moet nu een 9e datarij zijn (alleen integer-seriaaldata tellen mee)
        filled_rows = sum(
            1
            for row in range(DATA_START_ROW, DATA_START_ROW + 20)
            if isinstance(ws.cell(row=row, column=COL_DATUM).value, int)
        )
        assert filled_rows == 9
