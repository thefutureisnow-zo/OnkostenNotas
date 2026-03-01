"""Gedeelde constanten voor het OnkostenNotas-project."""

DUTCH_MONTHS = {
    1: "Januari", 2: "Februari", 3: "Maart", 4: "April",
    5: "Mei", 6: "Juni", 7: "Juli", 8: "Augustus",
    9: "September", 10: "Oktober", 11: "November", 12: "December",
}

# Omgekeerde lookup: "januari" → 1 (kleineletters)
DUTCH_MONTHS_REVERSE = {v.lower(): k for k, v in DUTCH_MONTHS.items()}
