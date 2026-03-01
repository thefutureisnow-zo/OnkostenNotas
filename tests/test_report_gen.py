"""
Tests voor report_gen.py — terminal tabel en HTML-rapport.
"""
from datetime import date
from pathlib import Path

import pytest

from email_parser import TicketData
from report_gen import format_summary_table, generate_html_report


def _make_ticket(
    order="TST001",
    from_s="Zottegem",
    to_s="Antwerpen-Zuid",
    direction="heen",
    travel_date=date(2026, 1, 7),
    price=14.0,
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


class TestFormatSummaryTable:
    def test_empty_list_returns_no_tickets_message(self):
        result = format_summary_table([])
        assert "geen tickets" in result.lower()

    def test_single_ticket_shows_all_fields(self):
        ticket = _make_ticket()
        result = format_summary_table([ticket])
        assert "Zottegem" in result
        assert "Antwerpen-Zuid" in result
        assert "heen" in result
        assert "14.00" in result
        assert "07/01/2026" in result

    def test_shows_nr_column(self):
        ticket = _make_ticket()
        result = format_summary_table([ticket])
        # Should have a Nr column with value 1
        assert "Nr" in result
        lines = result.strip().split("\n")
        # Data line should contain "1"
        data_lines = [l for l in lines if "Zottegem" in l]
        assert len(data_lines) == 1
        assert " 1 " in data_lines[0] or "| 1 |" in data_lines[0]

    def test_multiple_tickets_numbered_sequentially(self):
        tickets = [
            _make_ticket(order="T1", direction="heen", price=14.0),
            _make_ticket(order="T2", direction="terug", price=14.0),
        ]
        result = format_summary_table(tickets)
        lines = result.strip().split("\n")
        data_lines = [l for l in lines if "Zottegem" in l or "Antwerpen" in l]
        assert len(data_lines) == 2

    def test_total_row_shows_sum(self):
        tickets = [
            _make_ticket(order="T1", price=14.0),
            _make_ticket(order="T2", price=14.0),
        ]
        result = format_summary_table(tickets)
        assert "28.00" in result
        assert "Totaal" in result or "totaal" in result

    def test_round_trip_ticket(self):
        ticket = _make_ticket(direction="heen/terug", price=28.0)
        result = format_summary_table([ticket])
        # Description may be truncated, but direction should be partially visible
        assert "heen/" in result
        assert "28.00" in result

    def test_output_is_ascii_safe(self):
        """Alle output moet ASCII-veilig zijn (geen Unicode-symbolen)."""
        ticket = _make_ticket()
        result = format_summary_table([ticket])
        result.encode("ascii")  # Moet niet crashen

    def test_long_station_name_truncated(self):
        """Lange stationsnamen worden afgekapt zodat de tabel leesbaar blijft."""
        ticket = _make_ticket(
            from_s="Bruxelles-Midi / Brussel-Zuid",
            to_s="Antwerpen-Centraal",
        )
        result = format_summary_table([ticket])
        # Should not be excessively wide — description column capped
        lines = result.strip().split("\n")
        for line in lines:
            assert len(line) <= 90


class TestGenerateHtmlReport:
    def test_creates_file(self, tmp_path):
        tickets = [_make_ticket()]
        report_path = generate_html_report(tickets, tmp_path)
        assert report_path.exists()
        assert report_path.suffix == ".html"

    def test_contains_ticket_data(self, tmp_path):
        tickets = [_make_ticket()]
        report_path = generate_html_report(tickets, tmp_path)
        html = report_path.read_text(encoding="utf-8")
        assert "Zottegem" in html
        assert "Antwerpen-Zuid" in html
        assert "14.00" in html

    def test_contains_total(self, tmp_path):
        tickets = [
            _make_ticket(order="T1", price=14.0),
            _make_ticket(order="T2", price=14.0),
        ]
        report_path = generate_html_report(tickets, tmp_path)
        html = report_path.read_text(encoding="utf-8")
        assert "28.00" in html

    def test_filename_contains_timestamp(self, tmp_path):
        tickets = [_make_ticket()]
        report_path = generate_html_report(tickets, tmp_path)
        assert report_path.name.startswith("run_")
        assert report_path.name.endswith(".html")

    def test_empty_list_still_creates_report(self, tmp_path):
        report_path = generate_html_report([], tmp_path)
        assert report_path.exists()
        html = report_path.read_text(encoding="utf-8")
        assert "geen tickets" in html.lower() or "0" in html

    def test_screenshot_links_included(self, tmp_path):
        """Als screenshot_paths meegegeven worden, staan er links in het rapport."""
        tickets = [_make_ticket()]
        screenshots = [Path("screenshots/Januari 2026/trein_070126_heen_TST001.png")]
        report_path = generate_html_report(tickets, tmp_path, screenshot_paths=screenshots)
        html = report_path.read_text(encoding="utf-8")
        assert "trein_070126_heen_TST001.png" in html

    def test_valid_html_structure(self, tmp_path):
        tickets = [_make_ticket()]
        report_path = generate_html_report(tickets, tmp_path)
        html = report_path.read_text(encoding="utf-8")
        assert "<html" in html
        assert "</html>" in html
        assert "<table" in html
