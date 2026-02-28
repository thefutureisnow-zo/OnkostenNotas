"""
Verbinding met de Gmail API via OAuth2.
Haalt NMBS-bevestigingsmails op en geeft de HTML-inhoud terug.
"""
import base64
import email as email_lib
import re
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
GMAIL_QUERY = (
    'from:no-reply@sales.belgiantrain.be subject:"NMBS Mobile Ticket" newer_than:2y'
)


def get_gmail_service(client_secret_path: Path, token_path: Path):
    """
    Bouw een geauthenticeerde Gmail API-service.
    Bij de eerste keer opent er een browservenster voor de OAuth-toestemming.
    """
    creds = None

    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not client_secret_path.exists():
                raise FileNotFoundError(
                    f"client_secret.json niet gevonden op {client_secret_path}\n"
                    "Zie de README voor instructies om dit bestand aan te maken."
                )
            flow = InstalledAppFlow.from_client_secrets_file(
                str(client_secret_path), SCOPES
            )
            creds = flow.run_local_server(port=0)

        token_path.parent.mkdir(parents=True, exist_ok=True)
        token_path.write_text(creds.to_json(), encoding="utf-8")

    return build("gmail", "v1", credentials=creds)


def _extract_html_from_raw(raw_bytes: bytes) -> str | None:
    """Haal de HTML-body op uit een ruwe e-mailbericht (bytes)."""
    parsed = email_lib.message_from_bytes(raw_bytes)
    if parsed.is_multipart():
        for part in parsed.walk():
            if part.get_content_type() == "text/html":
                payload = part.get_payload(decode=True)
                charset = part.get_content_charset() or "utf-8"
                return payload.decode(charset, errors="replace")
    else:
        if parsed.get_content_type() == "text/html":
            payload = parsed.get_payload(decode=True)
            charset = parsed.get_content_charset() or "utf-8"
            return payload.decode(charset, errors="replace")
    return None


def fetch_nmbs_emails(
    client_secret_path: Path, token_path: Path
) -> list[tuple[str, str, str]]:
    """
    Haal alle NMBS-ticketmails op uit Gmail.

    Geeft een lijst terug van (message_id, order_number, html_body).
    Mails zonder leesbare HTML worden overgeslagen met een waarschuwing.
    """
    service = get_gmail_service(client_secret_path, token_path)

    messages = []
    page_token = None

    while True:
        kwargs = {"userId": "me", "q": GMAIL_QUERY, "maxResults": 500}
        if page_token:
            kwargs["pageToken"] = page_token
        result = service.users().messages().list(**kwargs).execute()
        messages.extend(result.get("messages", []))
        page_token = result.get("nextPageToken")
        if not page_token:
            break

    emails = []
    for msg in messages:
        msg_data = (
            service.users()
            .messages()
            .get(userId="me", id=msg["id"], format="raw")
            .execute()
        )
        raw = base64.urlsafe_b64decode(msg_data["raw"] + "==")
        html_body = _extract_html_from_raw(raw)

        if not html_body:
            print(f"  Waarschuwing: geen HTML gevonden in bericht {msg['id']}, overgeslagen.")
            continue

        # Haal bestelnummer op uit het onderwerp
        parsed_msg = email_lib.message_from_bytes(raw)
        subject = parsed_msg.get("Subject", "")
        order_match = re.search(r"([A-Z0-9]+)\s*-\s*NMBS Mobile Ticket", subject)
        order_number = order_match.group(1) if order_match else msg["id"]

        emails.append((msg["id"], order_number, html_body))

    return emails
