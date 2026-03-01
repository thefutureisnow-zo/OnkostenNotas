"""
Genereer een samenvattingstabel (terminal) en een HTML-rapport na verwerking.
"""
from datetime import datetime
from pathlib import Path

from email_parser import TicketData


def format_summary_table(tickets: list[TicketData]) -> str:
    """Geeft een ASCII-tabel terug met alle verwerkte tickets en een totaalrij."""
    if not tickets:
        return "Geen tickets verwerkt deze sessie."

    # Kolombreedtes
    w_nr = 4
    w_datum = 12
    w_omschr = 40
    w_bedrag = 10

    header = (
        f"| {'Nr':>{w_nr}} | {'Datum':<{w_datum}} | {'Omschrijving':<{w_omschr}} | {'Bedrag':>{w_bedrag}} |"
    )
    sep = f"+{'-' * (w_nr + 2)}+{'-' * (w_datum + 2)}+{'-' * (w_omschr + 2)}+{'-' * (w_bedrag + 2)}+"

    lines = [sep, header, sep]

    total = 0.0
    for i, t in enumerate(tickets, 1):
        desc = f"Trein {t.from_station} - {t.to_station} {t.direction}"
        if len(desc) > w_omschr:
            desc = desc[: w_omschr - 3] + "..."
        datum = t.travel_date.strftime("%d/%m/%Y")
        bedrag = f"{t.price:.2f}"
        total += t.price
        lines.append(
            f"| {i:>{w_nr}} | {datum:<{w_datum}} | {desc:<{w_omschr}} | {bedrag:>{w_bedrag}} |"
        )

    lines.append(sep)
    totaal_label = "Totaal:"
    padding = w_nr + 2 + w_datum + 2 + w_omschr + 2 + 3  # cols before bedrag
    lines.append(f"|{' ' * padding}{totaal_label:>{w_bedrag}} |")
    lines.append(f"|{' ' * padding}{total:>{w_bedrag}.2f} |")
    lines.append(sep)

    return "\n".join(lines)


def generate_html_report(
    tickets: list[TicketData],
    reports_dir: Path,
    screenshot_paths: list[Path] | None = None,
) -> Path:
    """Genereer een HTML-rapport en sla op in reports_dir. Geeft het pad terug."""
    reports_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_path = reports_dir / f"run_{timestamp}.html"

    screenshots = screenshot_paths or []

    total = sum(t.price for t in tickets)
    rows_html = ""

    if not tickets:
        rows_html = (
            '<tr><td colspan="5" style="text-align:center;padding:12px;">'
            "Geen tickets verwerkt deze sessie.</td></tr>"
        )
    else:
        for i, t in enumerate(tickets, 1):
            desc = f"Trein {t.from_station} - {t.to_station} {t.direction}"
            datum = t.travel_date.strftime("%d/%m/%Y")
            screenshot_cell = ""
            if i - 1 < len(screenshots):
                fname = screenshots[i - 1].name
                screenshot_cell = f'<a href="{screenshots[i - 1]}">{fname}</a>'
            rows_html += (
                f"<tr>"
                f"<td>{i}</td>"
                f"<td>{datum}</td>"
                f"<td>{desc}</td>"
                f"<td style='text-align:right'>{t.price:.2f}</td>"
                f"<td>{screenshot_cell}</td>"
                f"</tr>\n"
            )

    html = f"""<!DOCTYPE html>
<html lang="nl">
<head>
<meta charset="utf-8">
<title>NMBS Onkostennota - {timestamp}</title>
<style>
  body {{ font-family: Arial, sans-serif; margin: 20px; }}
  table {{ border-collapse: collapse; width: 100%; max-width: 900px; }}
  th, td {{ border: 1px solid #ccc; padding: 8px; }}
  th {{ background: #f0f0f0; text-align: left; }}
  tr:nth-child(even) {{ background: #fafafa; }}
  .total {{ font-weight: bold; text-align: right; padding-top: 12px; }}
</style>
</head>
<body>
<h1>NMBS Onkostennota</h1>
<p>Gegenereerd: {datetime.now().strftime("%d/%m/%Y %H:%M:%S")}</p>
<p>Aantal tickets: {len(tickets)}</p>
<table>
<thead>
<tr><th>Nr</th><th>Datum</th><th>Omschrijving</th><th>Bedrag (EUR)</th><th>Screenshot</th></tr>
</thead>
<tbody>
{rows_html}
</tbody>
</table>
<p class="total">Totaal: EUR {total:.2f}</p>
</body>
</html>
"""
    report_path.write_text(html, encoding="utf-8")
    return report_path
