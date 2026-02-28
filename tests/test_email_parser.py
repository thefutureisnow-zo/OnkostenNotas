"""
Tests voor email_parser.py
"""
from datetime import date

import pytest

from email_parser import parse_nmbs_email, ParseError


class TestRoundTrip:
    def test_order_number(self, sample_html_round_trip):
        ticket = parse_nmbs_email(sample_html_round_trip)
        assert ticket.order_number == "UPL1IGGK"

    def test_stations(self, sample_html_round_trip):
        ticket = parse_nmbs_email(sample_html_round_trip)
        assert ticket.from_station == "Zottegem"
        assert ticket.to_station == "Antwerpen-Zuid"

    def test_direction(self, sample_html_round_trip):
        ticket = parse_nmbs_email(sample_html_round_trip)
        assert ticket.direction == "heen/terug"

    def test_date(self, sample_html_round_trip):
        ticket = parse_nmbs_email(sample_html_round_trip)
        assert ticket.travel_date == date(2026, 2, 13)

    def test_price(self, sample_html_round_trip):
        ticket = parse_nmbs_email(sample_html_round_trip)
        assert ticket.price == 28.0

    def test_html_preserved(self, sample_html_round_trip):
        ticket = parse_nmbs_email(sample_html_round_trip)
        assert "UPL1IGGK" in ticket.email_html


class TestSingleHeen:
    def test_direction(self, sample_html_single_heen):
        ticket = parse_nmbs_email(sample_html_single_heen)
        assert ticket.direction == "heen"

    def test_date(self, sample_html_single_heen):
        ticket = parse_nmbs_email(sample_html_single_heen)
        assert ticket.travel_date == date(2026, 1, 7)

    def test_price(self, sample_html_single_heen):
        ticket = parse_nmbs_email(sample_html_single_heen)
        assert ticket.price == 14.0


class TestSingleTerug:
    def test_direction(self, sample_html_single_terug):
        ticket = parse_nmbs_email(sample_html_single_terug)
        assert ticket.direction == "terug"

    def test_date(self, sample_html_single_terug):
        ticket = parse_nmbs_email(sample_html_single_terug)
        assert ticket.travel_date == date(2026, 1, 7)


class TestParseError:
    def test_missing_order_number(self):
        with pytest.raises(ParseError, match="Bestelnummer"):
            parse_nmbs_email("<html><body>geen bestelnummer</body></html>")

    def test_missing_stations(self):
        html = """<html><body>
            <td>Bestelnummer: TEST001</td>
            <div>2e klas, Enkel</div>
            <div>Heen: 01/01/2026</div>
            <td>Totaalbedrag :</td><td>€ 14,00</td>
        </body></html>"""
        with pytest.raises(ParseError, match="Van/Naar"):
            parse_nmbs_email(html)
