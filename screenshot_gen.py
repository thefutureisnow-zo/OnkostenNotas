"""
Maakt een PNG-screenshot van de NMBS-bevestigingsmail.
Vereist dat Google Chrome geïnstalleerd is.
"""
from datetime import date
from pathlib import Path

from email_parser import TicketData

DUTCH_MONTHS = {
    1: "Januari", 2: "Februari", 3: "Maart", 4: "April",
    5: "Mei", 6: "Juni", 7: "Juli", 8: "Augustus",
    9: "September", 10: "Oktober", 11: "November", 12: "December",
}


def _screenshot_filename(ticket: TicketData) -> str:
    """Bijv. trein_130226_heenenterug.png"""
    date_str = ticket.travel_date.strftime("%d%m%y")
    direction_slug = ticket.direction.replace("/", "en")  # heen/terug → heenenterug
    return f"trein_{date_str}_{direction_slug}.png"


def _month_folder(d: date, screenshots_dir: Path) -> Path:
    folder_name = f"{DUTCH_MONTHS[d.month]} {d.year}"
    folder = screenshots_dir / folder_name
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def save_screenshot(ticket: TicketData, screenshots_dir: Path) -> Path:
    """
    Render de e-mail HTML naar een PNG en sla op in de juiste maandmap.
    Geeft het pad naar de opgeslagen PNG terug.
    Gooit een RuntimeError als Chrome niet beschikbaar is.
    """
    try:
        from html2image import HtmlToImage
    except ImportError:
        raise RuntimeError(
            "html2image is niet geïnstalleerd. Voer uit: pip install html2image"
        )

    output_dir = _month_folder(ticket.travel_date, screenshots_dir)
    filename = _screenshot_filename(ticket)
    output_path = output_dir / filename

    if output_path.exists():
        # Al aanwezig (bijv. bij herverwerking); niet overschrijven
        return output_path

    try:
        hti = HtmlToImage(
            output_path=str(output_dir),
            custom_flags=["--no-sandbox", "--disable-gpu"],
        )
        hti.screenshot(
            html_str=ticket.email_html,
            save_as=filename,
            size=(800, 1400),
        )
    except Exception as exc:
        raise RuntimeError(
            f"Screenshot mislukt voor {ticket.order_number}. "
            "Controleer of Google Chrome geïnstalleerd is.\n"
            f"Technische details: {exc}"
        ) from exc

    return output_path
