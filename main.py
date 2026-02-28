"""
Hoofdscript — voer dagelijks uit om nieuwe NMBS-tickets te verwerken.

Gebruik:
    python main.py
"""
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
from excel_updater import add_ticket_to_excel
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
)


def _prompt(question: str, default_yes: bool = True) -> bool:
    hint = "[J/n]" if default_yes else "[j/N]"
    answer = input(f"      {question} {hint}: ").strip().lower()
    if not answer:
        return default_yes
    return answer in ("j", "y", "ja", "yes")


def _print_ticket(ticket: TicketData, index: int, total: int) -> None:
    print(
        f"\n[{index}/{total}] {ticket.from_station} → {ticket.to_station}"
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
            tickets.append(parse_nmbs_email(html_body))
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
            print(f"\n  ⚠  Dit ticket is gekocht op een {label}.")
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

        mark_processed(ticket.order_number, state)
        save_state(state, config.STATE_FILE)
        added += 1
        print(f"      ✓ Toegevoegd aan {config.EXCEL_PATH.name}")

    print(f"\nKlaar: {added} ticket(s) toegevoegd", end="")
    if skipped_weekend:
        print(f", {skipped_weekend} weekend/feestdag(en) overgeslagen", end="")
    print(".")


if __name__ == "__main__":
    main()
