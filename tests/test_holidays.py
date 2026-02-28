"""
Tests voor holidays_be.py
"""
from datetime import date

from holidays_be import is_work_day, day_type_label


class TestIsWorkDay:
    def test_regular_monday(self):
        assert is_work_day(date(2026, 2, 2)) is True  # maandag

    def test_saturday(self):
        assert is_work_day(date(2026, 2, 7)) is False  # zaterdag

    def test_sunday(self):
        assert is_work_day(date(2026, 2, 8)) is False  # zondag

    def test_new_year(self):
        assert is_work_day(date(2026, 1, 1)) is False  # Nieuwjaar

    def test_armistice_day(self):
        assert is_work_day(date(2026, 11, 11)) is False  # Wapenstilstand

    def test_christmas(self):
        assert is_work_day(date(2026, 12, 25)) is False  # Kerstmis

    def test_national_day(self):
        assert is_work_day(date(2026, 7, 21)) is False  # Belgische Nationale Feestdag

    def test_friday_workday(self):
        assert is_work_day(date(2026, 1, 9)) is True  # vrijdag, geen feestdag


class TestDayTypeLabel:
    def test_saturday_label(self):
        assert day_type_label(date(2026, 2, 7)) == "zaterdag"

    def test_sunday_label(self):
        assert day_type_label(date(2026, 2, 8)) == "zondag"

    def test_holiday_label(self):
        label = day_type_label(date(2026, 1, 1))
        assert "feestdag" in label

    def test_workday_label(self):
        assert day_type_label(date(2026, 2, 2)) == "werkdag"
