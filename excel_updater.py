"""
Voegt een verwerkt ticket toe aan de juiste maandtab in de onkostennota.
"""
import calendar
import re
from datetime import date
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

from email_parser import TicketData

DUTCH_MONTHS = {
    1: "Januari", 2: "Februari", 3: "Maart", 4: "April",
    5: "Mei", 6: "Juni", 7: "Juli", 8: "Augustus",
    9: "September", 10: "Oktober", 11: "November", 12: "December",
}

# Rijen waar de data in staan (inclusief)
DATA_START_ROW = 8
DATA_END_ROW = 15  # standaard 8 rijen; bij overflow wordt er een rij ingevoegd

# Kolomnummers (1-gebaseerd)
COL_DATUM = 1        # A
COL_NR = 2           # B
COL_OMSCHRIJVING = 3 # C
COL_CURR = 4         # D
COL_VERVOER = 6      # F
COL_TOTAAL = 12      # L


def sheet_name_for_date(d: date) -> str:
    return f"{DUTCH_MONTHS[d.month]} {d.year}"


def date_to_excel_serial(d: date) -> int:
    """Zet een Python-datum om naar een Excel-serieel getal."""
    # Excel telt dagen vanaf 1900-01-00 (met de bekende 1900-schrikkeljaarfout)
    delta = d - date(1899, 12, 30)
    return delta.days


def _find_next_data_row(ws) -> int:
    """Geeft het rijnummer van de eerste lege rij in het datablok."""
    for row in range(DATA_START_ROW, DATA_END_ROW + 1):
        if ws.cell(row=row, column=COL_DATUM).value is None:
            return row
    return DATA_END_ROW + 1  # overflow: één rij voorbij het standaard bereik


def _insert_overflow_row(ws, insert_at: int) -> None:
    """
    Voegt een lege rij in vóór `insert_at` en werkt de SOM-formules bij
    in het samenvattingsgedeelte zodat de nieuwe rij meegeteld wordt.
    """
    ws.insert_rows(insert_at)

    # Stel de L-formule in voor de nieuw ingevoegde rij
    ws.cell(row=insert_at, column=COL_TOTAAL).value = (
        f"=SUM(E{insert_at}:K{insert_at})"
    )

    new_last_data_row = insert_at  # dit is nu de laaste datarij

    # Werk SOM-formules bij in het samenvattingsdeel (rijen na de data)
    for summary_row in range(new_last_data_row + 1, new_last_data_row + 20):
        for col_idx in range(1, COL_TOTAAL + 1):
            cell = ws.cell(row=summary_row, column=col_idx)
            if cell.value and isinstance(cell.value, str) and "SUM(" in cell.value.upper():
                col_letter = get_column_letter(col_idx)
                # Vergroot het bereik: SUM(X8:X15) → SUM(X8:X<new_last>)
                cell.value = re.sub(
                    rf"SUM\({col_letter}(\d+):{col_letter}(\d+)\)",
                    lambda m, c=col_letter, nl=new_last_data_row: (
                        f"SUM({c}{m.group(1)}:{c}{nl})"
                    ),
                    cell.value,
                    flags=re.IGNORECASE,
                )


def _create_month_sheet(wb, sheet_name: str):
    """
    Maakt een nieuw maandblad door het meest recente blad te kopiëren,
    de datarijen te wissen en de maandnaam bij te werken.
    """
    template = wb.worksheets[-1]
    new_ws = wb.copy_worksheet(template)
    new_ws.title = sheet_name

    # Wis datarijen (kolommen A t/m K; L bevat formules die we bewaren)
    for row in range(DATA_START_ROW, DATA_END_ROW + 1):
        for col in range(1, COL_TOTAAL):  # A t/m K
            new_ws.cell(row=row, column=col).value = None

    # Maandnaam (cel B5 heeft een spatie vóór de naam, conform het origineel)
    new_ws["B5"] = f" {sheet_name}"

    # Datumbereik (K4 = eerste dag, K5 = laatste dag van de maand)
    parts = sheet_name.split()
    month_num = next(k for k, v in DUTCH_MONTHS.items() if v == parts[0])
    year_num = int(parts[1])
    first_day = date(year_num, month_num, 1)
    last_day = date(year_num, month_num, calendar.monthrange(year_num, month_num)[1])
    new_ws["K4"] = date_to_excel_serial(first_day)
    new_ws["K5"] = date_to_excel_serial(last_day)

    return new_ws


def remove_ticket_from_excel(
    excel_path: Path, sheet_name: str, travel_date_serial: int, description: str
) -> bool:
    """
    Verwijdert een ticketrij uit de opgegeven maandtab.
    Geeft True terug als de rij gevonden en verwijderd is, anders False.
    Gooit een OSError als het bestand vergrendeld is (bijv. open in Excel).
    """
    try:
        wb = openpyxl.load_workbook(excel_path)
    except PermissionError:
        raise OSError(
            f"Het Excel-bestand is vergrendeld. Sluit het eerst in Excel: {excel_path}"
        )

    if sheet_name not in wb.sheetnames:
        print(f"  Waarschuwing: tabblad '{sheet_name}' niet gevonden in Excel.")
        return False

    ws = wb[sheet_name]

    # Bepaal de laatste gevulde datarij (voor overflow-detectie)
    last_data_row = DATA_START_ROW - 1
    for row in range(DATA_START_ROW, DATA_START_ROW + 50):
        if isinstance(ws.cell(row=row, column=COL_DATUM).value, int):
            last_data_row = row

    # Zoek de overeenkomende rij op datum EN omschrijving
    matches = []
    for row in range(DATA_START_ROW, last_data_row + 1):
        if (
            ws.cell(row=row, column=COL_DATUM).value == travel_date_serial
            and ws.cell(row=row, column=COL_OMSCHRIJVING).value == description
        ):
            matches.append(row)

    if not matches:
        print(
            f"  Waarschuwing: rij voor '{description}' niet gevonden in '{sheet_name}'."
            " Al handmatig verwijderd?"
        )
        return False

    if len(matches) > 1:
        print(
            f"  Waarschuwing: meerdere overeenkomende rijen in '{sheet_name}'."
            " Eerste rij wordt verwijderd."
        )
    target_row = matches[0]

    ws.delete_rows(target_row, 1)
    new_last_data_row = last_data_row - 1

    # Herschrijf per-rij L-formule voor alle datarijen op en na de verwijderde positie
    # (openpyxl past rijverwijzingen in formules NIET automatisch aan)
    for row in range(target_row, new_last_data_row + 1):
        if isinstance(ws.cell(row=row, column=COL_DATUM).value, int):
            ws.cell(row=row, column=COL_TOTAAL).value = f"=SUM(E{row}:K{row})"

    # Hernummer kolom B (Nr) voor alle overblijvende datarijen
    nr = 1
    for row in range(DATA_START_ROW, new_last_data_row + 1):
        if isinstance(ws.cell(row=row, column=COL_DATUM).value, int):
            ws.cell(row=row, column=COL_NR).value = nr
            nr += 1

    # Bij overflow: verklein SOM-bereiken in de samenvattingsrijen (alle kolommen)
    if last_data_row > DATA_END_ROW:
        for summary_row in range(new_last_data_row + 1, new_last_data_row + 20):
            for col_idx in range(1, COL_TOTAAL + 1):
                cell = ws.cell(row=summary_row, column=col_idx)
                if cell.value and isinstance(cell.value, str) and "SUM(" in cell.value.upper():
                    col_letter = get_column_letter(col_idx)
                    cell.value = re.sub(
                        rf"SUM\({col_letter}(\d+):{col_letter}(\d+)\)",
                        lambda m, c=col_letter, nl=new_last_data_row: (
                            f"SUM({c}{m.group(1)}:{c}{nl})"
                        ),
                        cell.value,
                        flags=re.IGNORECASE,
                    )

    wb.save(excel_path)
    return True


def add_ticket_to_excel(ticket: TicketData, excel_path: Path) -> None:
    """
    Voegt het ticket als nieuwe rij toe aan de juiste maandtab.
    Als de tab nog niet bestaat, wordt ze aangemaakt op basis van de vorige maand.
    Gooit een OSError als het bestand vergrendeld is (bijv. open in Excel).
    """
    sheet_name = sheet_name_for_date(ticket.travel_date)

    try:
        wb = openpyxl.load_workbook(excel_path)
    except PermissionError:
        raise OSError(
            f"Het Excel-bestand is vergrendeld. Sluit het eerst in Excel: {excel_path}"
        )

    if sheet_name not in wb.sheetnames:
        ws = _create_month_sheet(wb, sheet_name)
    else:
        ws = wb[sheet_name]

    next_row = _find_next_data_row(ws)

    if next_row > DATA_END_ROW:
        # Alle standaardrijen zijn vol: voeg een extra rij in
        _insert_overflow_row(ws, next_row)

    # Schrijf de ticketgegevens
    nr = next_row - DATA_START_ROW + 1
    description = (
        f"Trein {ticket.from_station} - {ticket.to_station} {ticket.direction}"
    )

    ws.cell(row=next_row, column=COL_DATUM).value = date_to_excel_serial(ticket.travel_date)
    ws.cell(row=next_row, column=COL_NR).value = nr
    ws.cell(row=next_row, column=COL_OMSCHRIJVING).value = description
    ws.cell(row=next_row, column=COL_CURR).value = "EUR"
    ws.cell(row=next_row, column=COL_VERVOER).value = ticket.price

    # Zorg dat de L-formule aanwezig is (normaal al zo bij nieuw blad, maar voor zekerheid)
    if ws.cell(row=next_row, column=COL_TOTAAL).value is None:
        ws.cell(row=next_row, column=COL_TOTAAL).value = (
            f"=SUM(E{next_row}:K{next_row})"
        )

    wb.save(excel_path)
