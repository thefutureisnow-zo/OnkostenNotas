"""
Hoofdscript — voer dagelijks uit om nieuwe NMBS-tickets te verwerken.

Gebruik:
    python main.py                      # alle tickets
    python main.py --month januari      # alleen januari (huidig jaar)
    python main.py --month "maart 2025" # alleen maart 2025
    python main.py --reset              # wis de verwerkte-ticketslijst
"""
import argparse
import sys
from datetime import date
from pathlib import Path

try:
    import config
except ModuleNotFoundError:
    print(
        "Fout: config.py niet gevonden.\n"
        "Kopieer config.example.py naar config.py en pas de paden aan:\n\n"
        "    copy config.example.py config.py\n"
    )
    sys.exit(1)

from constants import DUTCH_MONTHS_REVERSE
from email_parser import TicketData, parse_nmbs_email, ParseError
from excel_updater import (
    add_ticket_to_excel,
    excel_path_for_date,
    remove_ticket_from_excel,
    sheet_name_for_date,
    date_to_excel_serial,
)
from gmail_client import fetch_nmbs_emails
from holidays_be import is_work_day, day_type_label
from report_gen import format_summary_table, generate_html_report
from screenshot_gen import save_screenshot
from state import (
    load_state,
    save_state,
    is_processed,
    is_skipped,
    mark_processed,
    mark_skipped_weekend,
    get_metadata,
)


def parse_month_arg(arg: str) -> tuple[int, int]:
    """Parseer een maandargument zoals 'januari' of 'februari 2025'.

    Geeft (maand, jaar) terug. Bij alleen een maandnaam wordt het huidige jaar gebruikt.
    Stopt het programma bij een ongeldige maandnaam.
    """
    parts = arg.strip().lower().split()
    month_name = parts[0]

    if month_name not in DUTCH_MONTHS_REVERSE:
        print(f"Fout: onbekende maand '{month_name}'.")
        print(f"Geldige maanden: {', '.join(DUTCH_MONTHS_REVERSE.keys())}")
        sys.exit(1)

    month = DUTCH_MONTHS_REVERSE[month_name]
    year = int(parts[1]) if len(parts) > 1 else date.today().year
    return month, year


def _prompt(question: str, default_yes: bool = True) -> bool:
    hint = "[J/n]" if default_yes else "[j/N]"
    answer = input(f"      {question} {hint}: ").strip().lower()
    if not answer:
        return default_yes
    return answer in ("j", "y", "ja", "yes")


def _print_ticket(ticket: TicketData, index: int, total: int) -> None:
    print(
        f"\n[{index}/{total}] {ticket.from_station} -> {ticket.to_station}"
        f" ({ticket.direction})"
    )
    print(
        f"      Datum : {ticket.travel_date.strftime('%d/%m/%Y')}  |"
        f"  Prijs : EUR {ticket.price:.2f}"
    )
    print(f"      Bestelnummer : {ticket.order_number}")


def main(month_filter: tuple[int, int] | None = None) -> None:
    excel_dir = config.EXCEL_DIR
    excel_dir.mkdir(parents=True, exist_ok=True)
    config.SCREENSHOTS_DIR.mkdir(parents=True, exist_ok=True)

    print("NMBS Onkostennota -- nieuwe tickets verwerken\n")
    if month_filter:
        from constants import DUTCH_MONTHS
        print(f"Maandfilter: {DUTCH_MONTHS[month_filter[0]]} {month_filter[1]}\n")
    print("Mails ophalen uit Gmail...")

    try:
        raw_emails = fetch_nmbs_emails(config.CLIENT_SECRET_PATH, config.TOKEN_PATH)
    except FileNotFoundError as exc:
        print(f"\nFout: {exc}")
        sys.exit(1)

    state = load_state(config.STATE_FILE)

    # Parseer alle nog niet-verwerkte mails
    tickets: list[TicketData] = []
    for _msg_id, order_number, html_body in raw_emails:
        if is_processed(order_number, state) or is_skipped(order_number, state):
            continue
        try:
            ticket = parse_nmbs_email(
                html_body,
                home_station=getattr(config, "HOME_STATION", None),
                office_station=getattr(config, "OFFICE_STATION", None),
            )
            tickets.append(ticket)
        except ParseError as exc:
            print(f"  Waarschuwing: {exc} -- overgeslagen.")

    # Filter op maand indien opgegeven
    if month_filter:
        target_month, target_year = month_filter
        tickets = [
            t for t in tickets
            if t.travel_date.month == target_month and t.travel_date.year == target_year
        ]

    if not tickets:
        print("Geen nieuwe tickets gevonden.")
        return

    tickets.sort(key=lambda t: t.travel_date)
    total = len(tickets)
    print(f"{total} nieuw(e) ticket(s) gevonden.\n")

    added = 0
    skipped_weekend = 0
    added_tickets: list[TicketData] = []
    screenshot_paths: list[Path] = []

    for i, ticket in enumerate(tickets, 1):
        _print_ticket(ticket, i, total)

        # Weekend / feestdag controle
        if not is_work_day(ticket.travel_date):
            label = day_type_label(ticket.travel_date)
            print(f"\n  (!!)  Dit ticket is gekocht op een {label}.")
            if not _prompt("Toch opnemen in de onkostennota?", default_yes=False):
                mark_skipped_weekend(ticket.order_number, state)
                save_state(state, config.STATE_FILE)
                print("      Permanent overgeslagen (wordt niet meer getoond).")
                skipped_weekend += 1
                continue

        if not _prompt("Toevoegen aan de onkostennota?", default_yes=True):
            print("      Overgeslagen (wordt volgende keer opnieuw getoond).")
            continue

        # Screenshot opslaan
        scr_path = None
        try:
            scr_path = save_screenshot(ticket, config.SCREENSHOTS_DIR)
            print(f"      Screenshot opgeslagen: {scr_path}")
        except RuntimeError as exc:
            print(f"      Waarschuwing: {exc}")
            print("      Verdergaan zonder screenshot...")

        # Excel bijwerken
        try:
            result_path = add_ticket_to_excel(ticket, excel_dir)
        except OSError as exc:
            print(f"\n  Fout: {exc}")
            print("  Dit ticket wordt NIET als verwerkt gemarkeerd. Probeer opnieuw.")
            continue

        excel_metadata = {
            "filename": result_path.name,
            "travel_date_serial": date_to_excel_serial(ticket.travel_date),
            "description": (
                f"Trein {ticket.from_station} - {ticket.to_station} {ticket.direction}"
            ),
        }
        mark_processed(ticket.order_number, state, metadata=excel_metadata)
        save_state(state, config.STATE_FILE)
        added += 1
        added_tickets.append(ticket)
        if scr_path:
            screenshot_paths.append(scr_path)
        print(f"      OK  Toegevoegd aan {result_path.name}")

    print(f"\nKlaar: {added} ticket(s) toegevoegd", end="")
    if skipped_weekend:
        print(f", {skipped_weekend} weekend/feestdag(en) overgeslagen", end="")
    print(".")

    # Samenvattingstabel
    if added_tickets:
        print(f"\n{format_summary_table(added_tickets)}")

        # HTML-rapport
        reports_dir = getattr(config, "REPORTS_DIR", None)
        if reports_dir:
            try:
                report_path = generate_html_report(
                    added_tickets, reports_dir, screenshot_paths=screenshot_paths or None
                )
                print(f"\nHTML-rapport: {report_path}")
                import webbrowser
                webbrowser.open(str(report_path))
            except OSError as exc:
                print(f"\nWaarschuwing: HTML-rapport kon niet worden aangemaakt: {exc}")


def reset_state() -> None:
    """Wis processed.json en verwijder de bijbehorende Excel-rijen na bevestiging."""
    state = load_state(config.STATE_FILE)
    processed_orders = state.get("processed", [])
    n_processed = len(processed_orders)
    n_skipped = len(state.get("skipped_weekend", []))

    print("NMBS Onkostennota -- verwerkte tickets wissen\n")
    print(f"  Verwerkte tickets   : {n_processed}")
    print(f"  Weekend-overgeslagen: {n_skipped}")
    print()

    if n_processed == 0 and n_skipped == 0:
        print("Niets te wissen -- de lijst is al leeg.")
        return

    orders_with_meta = [o for o in processed_orders if get_metadata(o, state) is not None]
    orders_without_meta = [o for o in processed_orders if get_metadata(o, state) is None]

    if orders_with_meta:
        print(f"  {len(orders_with_meta)} ticket(s) worden ook automatisch uit Excel verwijderd.")
    if orders_without_meta:
        print(
            f"  {len(orders_without_meta)} ticket(s) zonder Excel-info:"
            " verwijder deze rijen eventueel handmatig."
        )
    print()

    answer = input("  Wil je de lijst echt wissen? [j/N]: ").strip().lower()
    if answer not in ("j", "y", "ja", "yes"):
        print("Geannuleerd.")
        return

    # Verwijder Excel-rijen voor tickets met metadata.
    excel_dir = config.EXCEL_DIR
    if orders_with_meta:
        for order in orders_with_meta:
            meta = get_metadata(order, state)
            excel_path = excel_dir / meta["filename"]
            try:
                removed = remove_ticket_from_excel(
                    excel_path,
                    meta["travel_date_serial"],
                    meta["description"],
                )
                if removed:
                    print(
                        f"  Verwijderd uit Excel: {meta['description']}"
                        f" ({meta['filename']})"
                    )
            except OSError as exc:
                print(f"\n  Fout: {exc}")
                print("  Reset afgebroken. Los het probleem op en probeer opnieuw.")
                return

    if config.STATE_FILE.exists():
        config.STATE_FILE.unlink()
    print(f"OK  {config.STATE_FILE.name} gewist. Alle tickets worden opnieuw aangeboden.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="NMBS Onkostennota verwerker")
    parser.add_argument(
        "--reset",
        action="store_true",
        help="Wis de lijst van verwerkte tickets (processed.json)",
    )
    parser.add_argument(
        "--month",
        type=str,
        default=None,
        help="Verwerk alleen tickets van deze maand (bijv. 'januari' of 'maart 2025')",
    )
    args = parser.parse_args()

    if args.reset:
        reset_state()
    else:
        month_filter = parse_month_arg(args.month) if args.month else None
        main(month_filter=month_filter)
