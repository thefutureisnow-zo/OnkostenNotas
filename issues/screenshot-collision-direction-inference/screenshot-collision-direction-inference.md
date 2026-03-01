# screenshot-collision-direction-inference

## 1. The problem

When two separate "Enkel" tickets fell on the same date (e.g. 04/02/2026), both screenshots
were saved under the same filename (`trein_040226_heen.png`), so the second overwrote the first.

Additionally, a ticket for Antwerpen-Zuid → Zottegem (a return-home trip) was labelled "heen"
because NMBS happened to put its date under the "Heen:" field in the confirmation email.
The direction shown and written to Excel was therefore wrong.

Observed output:

```
[36/39] Antwerpen-Zuid -> Zottegem (heen)   ← should be "terug"
      Screenshot opgeslagen: ...trein_040226_heen.png

[37/39] Zottegem -> Antwerpen-Zuid (heen)
      Screenshot opgeslagen: ...trein_040226_heen.png   ← same filename, overwrites previous
```

## 2. How we spotted it

User ran `python main.py` and noticed the second ticket's screenshot path matched the first,
and that "Antwerpen-Zuid -> Zottegem" was marked "heen" instead of "terug".

## 3. Root cause

**Filename collision**: `_screenshot_filename()` in `screenshot_gen.py` built filenames from
only `{date}_{direction}`, with no ticket-unique component. Two tickets on the same day with
the same direction produced identical filenames.

**Wrong direction**: `parse_nmbs_email()` in `email_parser.py` inferred direction for "Enkel"
tickets solely from which date label ("Heen:" / "Terug:") NMBS used in the email. NMBS does not
guarantee that label matches the travel direction — it is an artifact of how the order was
structured, not a reliable indicator of the actual journey direction.

## 4. Fix

**Filename**: Append `ticket.order_number` to the filename.
New format: `trein_{DDMMYY}_{direction_slug}_{order_number}.png` (e.g. `trein_040226_heen_WR826GNF.png`).
Order numbers are unique per ticket, so collisions are impossible.

**Direction**: Added `infer_direction(from_station, to_station, home_station, office_station)`
to `email_parser.py`. In `main.py`, after parsing, for "Enkel" tickets (direction "heen" or
"terug") the station pair is compared case-insensitively against `config.HOME_STATION` and
`config.OFFICE_STATION`. If they match, the inferred direction overrides the NMBS label.
Falls back to the label if stations are not configured (`getattr` guard for backward compat).

PR: fix/screenshot-collision-direction-inference
