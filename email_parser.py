"""
Parseert de HTML-inhoud van een NMBS-bevestigingsmail en geeft een
gestructureerd TicketData-object terug.
"""
import re
from dataclasses import dataclass
from datetime import date

from bs4 import BeautifulSoup


@dataclass
class TicketData:
    order_number: str
    from_station: str   # title-case, bijv. "Zottegem"
    to_station: str     # title-case, bijv. "Antwerpen-Zuid"
    direction: str      # "heen" | "terug" | "heen/terug"
    travel_date: date
    price: float
    email_html: str     # originele HTML, bewaard voor de screenshot


class ParseError(Exception):
    """Wordt gegeven als de e-mail niet de verwachte structuur heeft."""


def infer_direction(
    from_station: str,
    to_station: str,
    home_station: str,
    office_station: str,
) -> str | None:
    """
    Bepaal de richting op basis van het stationpaar.
    Geeft "heen" als je VAN thuis NAAR kantoor reist,
    "terug" als je VAN kantoor NAAR thuis reist,
    of None als de stations niet overeenkomen met het geconfigureerde paar.
    Vergelijking is hoofdletterongevoelig.
    """
    from_lower = from_station.lower()
    to_lower = to_station.lower()
    home_lower = home_station.lower()
    office_lower = office_station.lower()

    if from_lower == home_lower and to_lower == office_lower:
        return "heen"
    if from_lower == office_lower and to_lower == home_lower:
        return "terug"
    return None


def _title_station(name: str) -> str:
    """Zet een stationsnaam in title-case: ANTWERPEN-ZUID → Antwerpen-Zuid."""
    return name.strip().title()


def parse_nmbs_email(html: str) -> TicketData:
    """
    Parseer een NMBS-bevestigingsmail (HTML) en geef een TicketData terug.

    Ondersteunde gevallen:
      - "2e klas, Heen en terug" → direction="heen/terug", datum = Heen-datum
      - "2e klas, Enkel"         → direction="heen" of "terug" afhankelijk van
                                   welke datum aanwezig is
    """
    soup = BeautifulSoup(html, "lxml")
    full_text = soup.get_text(" ", strip=True)

    # --- Bestelnummer ---
    order_match = re.search(r"Bestelnummer:\s*([A-Z0-9]+)", full_text)
    if not order_match:
        raise ParseError("Bestelnummer niet gevonden in de e-mail.")
    order_number = order_match.group(1)

    # --- Van / Naar ---
    van_span = soup.find("span", string=re.compile(r"Van\s*:"))
    naar_span = soup.find("span", string=re.compile(r"Naar\s*:"))
    if not van_span or not naar_span:
        raise ParseError(f"[{order_number}] Van/Naar-stations niet gevonden.")

    from_station = _title_station(van_span.find_next_sibling("span").get_text())
    to_station = _title_station(naar_span.find_next_sibling("span").get_text())

    # --- Klasse & richting ---
    if "Heen en terug" in full_text:
        trip_type = "heen/terug"
    elif "Enkel" in full_text:
        trip_type = "enkel"
    else:
        raise ParseError(f"[{order_number}] Reistype (Enkel/Heen en terug) niet gevonden.")

    # --- Reisdatum(s) ---
    heen_match = re.search(r"Heen:\s*(\d{2}/\d{2}/\d{4})", full_text)
    terug_match = re.search(r"Terug:\s*(\d{2}/\d{2}/\d{4})", full_text)

    def parse_date(s: str) -> date:
        day, month, year = s.split("/")
        return date(int(year), int(month), int(day))

    if trip_type == "heen/terug":
        if not heen_match:
            raise ParseError(f"[{order_number}] Geen Heen-datum gevonden bij heen/terug.")
        direction = "heen/terug"
        travel_date = parse_date(heen_match.group(1))
    else:
        # Enkel: bepaal richting op basis van welke datum aanwezig is
        if heen_match and not terug_match:
            direction = "heen"
            travel_date = parse_date(heen_match.group(1))
        elif terug_match and not heen_match:
            direction = "terug"
            travel_date = parse_date(terug_match.group(1))
        elif heen_match and terug_match:
            # Beide aanwezig bij Enkel: gebruik Heen-datum, markeer als heen
            direction = "heen"
            travel_date = parse_date(heen_match.group(1))
        else:
            raise ParseError(f"[{order_number}] Geen reisdatum gevonden.")

    # --- Totaalbedrag ---
    # Gebruik full_text: "Totaalbedrag : € 28,00" → regex pakt het getal direct erna
    price_match = re.search(r"Totaalbedrag\s*:?[^\d]*([\d]+[,.][\d]+)", full_text)
    if not price_match:
        raise ParseError(f"[{order_number}] Totaalbedrag niet gevonden in e-mail.")
    price = float(price_match.group(1).replace(",", "."))


    return TicketData(
        order_number=order_number,
        from_station=from_station,
        to_station=to_station,
        direction=direction,
        travel_date=travel_date,
        price=price,
        email_html=html,
    )
