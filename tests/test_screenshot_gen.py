"""
Tests voor screenshot_gen.py — bestandsnaamgeneratie.
"""
from datetime import date

from email_parser import TicketData
from screenshot_gen import _screenshot_filename


def _make_ticket(order_number: str, direction: str = "heen") -> TicketData:
    return TicketData(
        order_number=order_number,
        from_station="Zottegem",
        to_station="Antwerpen-Zuid",
        direction=direction,
        travel_date=date(2026, 2, 4),
        price=14.0,
        email_html="",
    )


class TestScreenshotFilename:
    def test_filename_includes_order_number(self):
        ticket = _make_ticket("ABC12345")
        assert "ABC12345" in _screenshot_filename(ticket)

    def test_filename_format(self):
        ticket = _make_ticket("ABC12345", direction="heen")
        assert _screenshot_filename(ticket) == "trein_040226_heen_ABC12345.png"

    def test_round_trip_direction_slug(self):
        ticket = _make_ticket("XYZ00001", direction="heen/terug")
        assert _screenshot_filename(ticket) == "trein_040226_heenenterug_XYZ00001.png"

    def test_same_day_different_orders_produce_different_filenames(self):
        ticket_a = _make_ticket("ORDER111", direction="heen")
        ticket_b = _make_ticket("ORDER222", direction="heen")
        assert _screenshot_filename(ticket_a) != _screenshot_filename(ticket_b)

    def test_terug_direction(self):
        ticket = _make_ticket("ZZZ99999", direction="terug")
        assert _screenshot_filename(ticket) == "trein_040226_terug_ZZZ99999.png"
