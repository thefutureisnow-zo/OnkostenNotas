"""
Controleert of een datum een werkdag is (geen weekend of Belgische feestdag).
"""
from datetime import date
import holidays


def is_work_day(d: date) -> bool:
    """Geeft True terug als de datum een gewone werkdag is."""
    if d.weekday() >= 5:  # 5 = zaterdag, 6 = zondag
        return False
    be_holidays = holidays.Belgium(years=d.year)
    return d not in be_holidays


def day_type_label(d: date) -> str:
    """Geeft een beschrijving van waarom de dag geen werkdag is."""
    if d.weekday() == 5:
        return "zaterdag"
    if d.weekday() == 6:
        return "zondag"
    be_holidays = holidays.Belgium(years=d.year)
    if d in be_holidays:
        return f"feestdag ({be_holidays[d]})"
    return "werkdag"
