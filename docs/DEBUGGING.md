# Debugging Parse Failures

When `python main.py` shows `Waarschuwing: [ORDER] ... niet gevonden -- overgeslagen`, follow
these steps to find and fix the parser.

## 1. Inspect the real email HTML

If you have an `.eml` file, extract and inspect the relevant section:

```python
import email
from bs4 import BeautifulSoup

with open('path/to/EMAIL.eml', 'rb') as f:
    msg = email.message_from_bytes(f.read())

html = next(
    part.get_payload(decode=True).decode(part.get_content_charset() or 'utf-8', errors='replace')
    for part in msg.walk()
    if part.get_content_type() == 'text/html'
)
soup = BeautifulSoup(html, 'lxml')
```

Then probe what the parser is targeting. Examples:

```python
# Check what TDs contain a label (e.g. Totaalbedrag)
for td in soup.find_all('td'):
    t = td.get_text(strip=True)
    if 'Totaalbedrag' in t and len(t) < 200:
        print(len(t), repr(t), '| sib:', td.find_next_sibling('td'))

# Check what spans contain "Van :"
for s in soup.find_all('span'):
    if 'Van' in s.get_text():
        print(repr(s))
```

## 2. Find where the real structure differs

Key lessons learned from past debugging:

- **Wrapper TDs**: NMBS emails often have outer `<td>` containers whose `get_text()` includes ALL
  nested content. A `find_all("td")` loop can match these FIRST before reaching the specific row.
  `find_next_sibling("td")` then returns `None` because the outer wrapper has no siblings.
- **Prefer `full_text` regex over DOM navigation** for values that appear as plain text (prices,
  dates, order numbers). `full_text = soup.get_text(" ", strip=True)` is already computed in
  `parse_nmbs_email` and is more resilient to HTML restructuring.
- **Greedy regex order**: if `re.search` finds the first match and there are multiple numbers
  (e.g. individual ticket prices before the total), anchor the regex to the label:
  `re.search(r"Totaalbedrag\s*:?[^\d]*([\d]+[,.][\d]+)", full_text)`.
- **Station spans**: `soup.find("span", string=re.compile(r"Van\s*:"))` requires the span's
  full text to match. If NMBS changes the label spacing or language, this breaks.

## 3. Verify the fix against a real email before running tests

```bash
python3 -c "
import email
from email_parser import parse_nmbs_email
with open('path/to/EMAIL.eml', 'rb') as f:
    msg = email.message_from_bytes(f.read())
html = next(p.get_payload(decode=True).decode(p.get_content_charset() or 'utf-8', errors='replace')
            for p in msg.walk() if p.get_content_type() == 'text/html')
t = parse_nmbs_email(html)
print(t.order_number, t.from_station, t.to_station, t.direction, t.travel_date, t.price)
"
```

## 4. Run the full test suite

```bash
python3 -m pytest -v
```

All tests must pass before committing. If a test expectation is wrong (not the code), fix the
test -- but verify against the real email first to be sure.
