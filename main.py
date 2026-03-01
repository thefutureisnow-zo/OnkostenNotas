"""
Hoofdscript — voer dagelijks uit om nieuwe NMBS-tickets te verwerken.

Gebruik:
    python main.py           # normaal gebruik
    python main.py --reset   # wis de verwerkte-ticketslijst (processed.json)
"""
import argparse
import sys
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

from email_parser import TicketData, parse_nmbs_email, ParseError
from excel_updater import (
    add_ticket_to_excel,
    remove_ticket_from_excel,
    sheet_name_for_date,
    date_to_excel_serial,
)
from gmail_client import fetch_nmbs_emails
from holidays_be import is_work_day, day_type_label
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
        f"  Prijs : € {ticket.price:.2f}"
    )
    print(f"      Bestelnummer : {ticket.order_number}")


def main() -> None:
    # Zorg dat de data- en screenshotsmappen bestaan
    config.EXCEL_PATH.parent.mkdir(parents=True, exist_ok=True)
    config.SCREENSHOTS_DIR.mkdir(parents=True, exist_ok=True)

    if not config.EXCEL_PATH.exists():
        print(
            f"Fout: Excel-bestand niet gevonden op {config.EXCEL_PATH}\n"
            "Controleer het pad in config.py en zorg dat het bestand aanwezig is."
        )
        sys.exit(1)

    print("NMBS Onkostennota — nieuwe tickets verwerken\n")
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
            print(f"  Waarschuwing: {exc} — overgeslagen.")

    if not tickets:
        print("Geen nieuwe tickets gevonden.")
        return

    tickets.sort(key=lambda t: t.travel_date)
    total = len(tickets)
    print(f"{total} nieuw(e) ticket(s) gevonden.\n")

    added = 0
    skipped_weekend = 0

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
        try:
            screenshot_path = save_screenshot(ticket, config.SCREENSHOTS_DIR)
            print(f"      Screenshot opgeslagen: {screenshot_path}")
        except RuntimeError as exc:
            print(f"      Waarschuwing: {exc}")
            print("      Verdergaan zonder screenshot...")

        # Excel bijwerken
        try:
            add_ticket_to_excel(ticket, config.EXCEL_PATH)
        except OSError as exc:
            print(f"\n  Fout: {exc}")
            print("  Dit ticket wordt NIET als verwerkt gemarkeerd. Probeer opnieuw.")
            continue

        excel_metadata = {
            "sheet_name": sheet_name_for_date(ticket.travel_date),
            "travel_date_serial": date_to_excel_serial(ticket.travel_date),
            "description": (
                f"Trein {ticket.from_station} - {ticket.to_station} {ticket.direction}"
            ),
        }
        mark_processed(ticket.order_number, state, metadata=excel_metadata)
        save_state(state, config.STATE_FILE)
        added += 1
        print(f"      OK  Toegevoegd aan {config.EXCEL_PATH.name}")

    print(f"\nKlaar: {added} ticket(s) toegevoegd", end="")
    if skipped_weekend:
        print(f", {skipped_weekend} weekend/feestdag(en) overgeslagen", end="")
    print(".")


def reset_state() -> None:
    """Wis processed.json en verwijder de bijbehorende Excel-rijen na bevestiging."""
    state = load_state(config.STATE_FILE)
    processed_orders = state.get("processed", [])
    n_processed = len(processed_orders)
    n_skipped = len(state.get("skipped_weekend", []))

    print("NMBS Onkostennota — verwerkte tickets wissen\n")
    print(f"  Verwerkte tickets   : {n_processed}")
    print(f"  Weekend-overgeslagen: {n_skipped}")
    print()

    if n_processed == 0 and n_skipped == 0:
        print("Niets te wissen — de lijst is al leeg.")
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
    # Bij een OSError (bestand vergrendeld) wordt de reset afgebroken.
    if orders_with_meta and config.EXCEL_PATH.exists():
        for order in orders_with_meta:
            meta = get_metadata(order, state)
            try:
                removed = remove_ticket_from_excel(
                    config.EXCEL_PATH,
                    meta["sheet_name"],
                    meta["travel_date_serial"],
                    meta["description"],
                )
                if removed:
                    print(
                        f"  Verwijderd uit Excel: {meta['description']}"
                        f" ({meta['sheet_name']})"
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
    args = parser.parse_args()

    if args.reset:
        reset_state()
    else:
        main()
