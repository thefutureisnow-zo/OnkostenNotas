# Issue: UnicodeEncodeError crashes main.py on Windows before any ticket is processed

## The problem

Running `python main.py` on Windows crashed immediately after fetching emails, before
displaying any ticket or reaching the screenshot step. The user reported it as a
screenshot failure, because the script never made it that far.

```
UnicodeEncodeError: 'charmap' codec can't encode character '\u2192' in position 18:
character maps to <undefined>
```

## How we spotted it

Ran the script directly instead of inspecting code:

```bash
echo "j" | python main.py
```

The traceback pointed straight to `_print_ticket()` in `main.py`, line 47 — an f-string
containing `→` (U+2192).

## Root cause

Windows terminals default to `cp1252` encoding. Three characters in `main.py` fall outside
that encoding:

| Character | Unicode | Location         |
|-----------|---------|------------------|
| `→`       | U+2192  | `_print_ticket`  |
| `⚠`       | U+26A0  | weekend warning  |
| `✓`       | U+2713  | success message  |

## Fix

Replaced all three with plain ASCII equivalents in `main.py`:

- `→` -> `->`
- `⚠` -> `(!!)`
- `✓` -> `OK`

No imports, no encoding hacks, no runtime reconfiguration needed.

## Convention added

Added to `CLAUDE.md` under Key conventions:

> **Print statements**: use only plain ASCII characters in all `print()` calls.
> No arrows, checkmarks, warning signs, or other non-ASCII symbols.

## PR

[#4 fix: replace non-ASCII print chars to prevent UnicodeEncodeError on Windows](https://github.com/thefutureisnow-zo/OnkostenNotas/pull/4)
