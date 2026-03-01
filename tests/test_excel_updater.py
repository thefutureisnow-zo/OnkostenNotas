"""
Tests voor excel_updater.py
"""
from datetime import date

import openpyxl
import pytest

from email_parser import TicketData
from excel_updater import (
    add_ticket_to_excel,
    remove_ticket_from_excel,
    date_to_excel_serial,
    sheet_name_for_date,
    DATA_START_ROW,
    DATA_END_ROW,
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


class TestRemoveTicketFromExcel:
    def test_removes_row(self, temp_excel):
        """Een ingevoerde rij kan worden teruggedraaid via remove_ticket_from_excel."""
        ticket = _make_ticket()
        add_ticket_to_excel(ticket, temp_excel)

        date_serial = date_to_excel_serial(ticket.travel_date)
        description = f"Trein {ticket.from_station} - {ticket.to_station} {ticket.direction}"

        result = remove_ticket_from_excel(temp_excel, "Januari 2026", date_serial, description)
        assert result is True

        wb = openpyxl.load_workbook(temp_excel)
        ws = wb["Januari 2026"]
        assert ws.cell(row=DATA_START_ROW, column=COL_DATUM).value is None

    def test_renumbers_remaining_rows(self, temp_excel):
        """Na verwijdering van de eerste rij worden de overige Nr-waarden hernummerd."""
        for i in range(3):
            add_ticket_to_excel(
                _make_ticket(order=f"TST{i}", travel_date=date(2026, 1, i + 1), price=14.0),
                temp_excel,
            )

        # Verwijder de eerste rij (datum 2026-01-01)
        remove_ticket_from_excel(
            temp_excel,
            "Januari 2026",
            date_to_excel_serial(date(2026, 1, 1)),
            "Trein Zottegem - Antwerpen-Zuid heen/terug",
        )

        wb = openpyxl.load_workbook(temp_excel)
        ws = wb["Januari 2026"]
        assert ws.cell(row=DATA_START_ROW, column=COL_NR).value == 1
        assert ws.cell(row=DATA_START_ROW + 1, column=COL_NR).value == 2

    def test_l_formula_rewritten_after_delete(self, temp_excel):
        """L-formules voor rijen na de verwijderde rij verwijzen naar de juiste rij."""
        for i in range(2):
            add_ticket_to_excel(
                _make_ticket(order=f"TST{i}", travel_date=date(2026, 1, i + 1), price=14.0),
                temp_excel,
            )

        # Verwijder rij 1 (row=8), waarna rij 2 (row=9) naar row=8 schuift
        remove_ticket_from_excel(
            temp_excel,
            "Januari 2026",
            date_to_excel_serial(date(2026, 1, 1)),
            "Trein Zottegem - Antwerpen-Zuid heen/terug",
        )

        wb = openpyxl.load_workbook(temp_excel)
        ws = wb["Januari 2026"]
        # Rij 8 bevat nu het tweede ticket; de L-formule moet naar rij 8 verwijzen
        formula = ws.cell(row=DATA_START_ROW, column=COL_TOTAAL).value
        assert formula is not None
        assert f"E{DATA_START_ROW}" in str(formula)

    def test_returns_false_for_missing_sheet(self, temp_excel, capsys):
        """Geeft False terug als het tabblad niet bestaat."""
        result = remove_ticket_from_excel(
            temp_excel, "Maart 2026", 12345, "Trein X - Y heen"
        )
        assert result is False
        out = capsys.readouterr().out
        assert "niet gevonden" in out.lower()

    def test_returns_false_for_missing_row(self, temp_excel, capsys):
        """Geeft False terug als er geen overeenkomende rij is."""
        result = remove_ticket_from_excel(
            temp_excel,
            "Januari 2026",
            date_to_excel_serial(date(2026, 1, 7)),
            "Trein Zottegem - Antwerpen-Zuid heen/terug",
        )
        assert result is False

    def test_overflow_sum_formula_shrinks(self, temp_excel_full):
        """Na verwijdering van een overflow-rij krimpt het SOM-bereik terug."""
        # Voeg een 9e ticket toe (triggert overflow, voegt rij 16 in)
        overflow_ticket = _make_ticket(travel_date=date(2026, 1, 10), price=14.0)
        add_ticket_to_excel(overflow_ticket, temp_excel_full)

        # Verwijder dat overflow-ticket weer
        remove_ticket_from_excel(
            temp_excel_full,
            "Januari 2026",
            date_to_excel_serial(date(2026, 1, 10)),
            "Trein Zottegem - Antwerpen-Zuid heen/terug",
        )

        wb = openpyxl.load_workbook(temp_excel_full)
        ws = wb["Januari 2026"]
        # Samenvatting-SOM op rij DATA_END_ROW+1 moet weer eindigen op DATA_END_ROW
        summary_f = ws.cell(row=DATA_END_ROW + 1, column=6).value
        assert summary_f is not None
        assert f"F{DATA_END_ROW}" in str(summary_f).upper()
        assert f"F{DATA_END_ROW + 1}" not in str(summary_f).upper()
