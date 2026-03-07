"""
Invoice Automation — invoice_processor.py
Polls IMAP inbox, extracts PDF data via Claude API, creates Notion entries.
"""

import imaplib
import email
import time
import base64
import json
import logging
import os
from datetime import datetime
from email.header import decode_header

import anthropic
import requests

# ── Logging ──────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)s  %(message)s",
)
log = logging.getLogger(__name__)

# ── Config (injected via environment variables) ───────────────────────────────
IMAP_HOST     = os.environ["IMAP_HOST"]          # e.g. "imap.one.com"
IMAP_USER     = os.environ["IMAP_USER"]          # full email address
IMAP_PASSWORD = os.environ["IMAP_PASSWORD"]

ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]
NOTION_API_KEY    = os.environ["NOTION_API_KEY"]
NOTION_DATABASE_ID = os.environ["NOTION_DATABASE_ID"]

POLL_INTERVAL_SECONDS = 15 * 60  # 15 minutes

# ── Claude client ─────────────────────────────────────────────────────────────
claude = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

EXTRACTION_PROMPT = """
You are an invoice data extraction assistant. Analyse the attached PDF invoice and
return ONLY a valid JSON object — no explanation, no markdown, no code fences.

Fields to extract:
{
  "supplier_name":   string | null,
  "amount_dkk":      number | null,    // numeric only, no currency symbols
  "due_date":        string | null,    // ISO 8601 format: YYYY-MM-DD, or null if not found
  "payment_type":    "Auto-debit" | "Manual" | "Unknown",
  "notes":           string            // brief extraction notes or uncertainties
}

Rules:
- payment_type: look for keywords like "betalingsservice", "PBS", "direct debit",
  "automatisk betaling" → Auto-debit; otherwise → Manual; if genuinely unclear → Unknown.
- amount_dkk: convert any decimal commas to decimal points (Danish format: 1.234,56 → 1234.56).
- If due_date is missing or ambiguous, set to null (do NOT guess).
- supplier_name: prefer the legal entity name over trading name if both present.
"""


def extract_invoice_data(pdf_bytes: bytes) -> dict:
    """Send PDF to Claude and return structured invoice data."""
    pdf_b64 = base64.standard_b64encode(pdf_bytes).decode("utf-8")

    message = claude.messages.create(
        model="claude-opus-4-6",
        max_tokens=512,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "document",
                        "source": {
                            "type": "base64",
                            "media_type": "application/pdf",
                            "data": pdf_b64,
                        },
                    },
                    {"type": "text", "text": EXTRACTION_PROMPT},
                ],
            }
        ],
    )

    raw = message.content[0].text.strip()
    # Strip markdown code fences if present
    if raw.startswith('''\\''')[1]
        if raw.startswith('json'):
            raw = raw[4:]
        raw = raw.strip()
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        log.warning("Claude returned non-JSON: %s", raw)
        return {
            "supplier_name": None,
            "amount_dkk": None,
            "due_date": None,
            "payment_type": "Unknown",
            "notes": f"Extraction failed — raw response: {raw[:200]}",
        }


# ── Notion helpers ────────────────────────────────────────────────────────────
NOTION_VERSION = "2022-06-28"
NOTION_HEADERS = {
    "Authorization": f"Bearer {NOTION_API_KEY}",
    "Content-Type": "application/json",
    "Notion-Version": NOTION_VERSION,
}


def upload_pdf_to_notion(pdf_bytes: bytes, filename: str) -> str | None:
    """
    Upload PDF via Notion file upload API and return the file_upload_id.
    Falls back to None if upload fails (page will be created without attachment).
    """
    # Step 1 — create upload
    resp = requests.post(
        "https://api.notion.com/v1/file_uploads",
        headers=NOTION_HEADERS,
        json={"content_type": "application/pdf"},
    )
    if not resp.ok:
        log.warning("Notion file upload init failed: %s", resp.text)
        return None

    upload = resp.json()
    upload_url = upload.get("upload_url")
    file_upload_id = upload.get("id")

    # Step 2 — send bytes
    upload_resp = requests.post(
        upload_url,
        headers={"Authorization": f"Bearer {NOTION_API_KEY}", "Notion-Version": NOTION_VERSION},
        files={"file": (filename, pdf_bytes, "application/pdf")},
    )
    if not upload_resp.ok:
        log.warning("Notion file upload failed: %s", upload_resp.text)
        return None

    return file_upload_id


def create_notion_entry(data: dict, pdf_bytes: bytes | None, filename: str, email_subject: str):
    """Create a Notion database page for one invoice."""
    due_date_missing = data.get("due_date") is None

    # Base properties
    properties = {
        "Name": {
            "title": [{"text": {"content": data.get("supplier_name") or "Unknown Supplier"}}]
        },
        "Status": {
            "status": {"name": "Pending Review"}
        },
        "Payment Type": {
            "select": {"name": data.get("payment_type", "Unknown")}
        },
        "Source Email Subject": {
            "rich_text": [{"text": {"content": email_subject[:200]}}]
        },
        "Notes": {
            "rich_text": [{"text": {"content": _build_notes(data, due_date_missing)}}]
        },
    }

    # Amount (DKK)
    if data.get("amount_dkk") is not None:
        properties["Amount (DKK)"] = {"number": float(data["amount_dkk"])}

    # Due date
    if data.get("due_date"):
        properties["Due Date"] = {"date": {"start": data["due_date"]}}

    # PDF attachment
    if pdf_bytes:
        file_upload_id = upload_pdf_to_notion(pdf_bytes, filename)
        if file_upload_id:
            properties["PDF"] = {
                "files": [
                    {
                        "name": filename,
                        "type": "file_upload",
                        "file_upload": {"id": file_upload_id},
                    }
                ]
            }

    payload = {
        "parent": {"database_id": NOTION_DATABASE_ID},
        "properties": properties,
    }

    resp = requests.post(
        "https://api.notion.com/v1/pages",
        headers=NOTION_HEADERS,
        json=payload,
    )
    if resp.ok:
        log.info("Notion entry created: %s", data.get("supplier_name"))
    else:
        log.error("Notion creation failed: %s", resp.text)


def _build_notes(data: dict, due_date_missing: bool) -> str:
    parts = []
    if due_date_missing:
        parts.append("⚠️ No due date found on invoice — please set manually.")
    if data.get("payment_type") == "Unknown":
        parts.append("⚠️ Payment type unclear — please set Auto-debit or Manual.")
    if data.get("notes"):
        parts.append(data["notes"])
    return "  ".join(parts) if parts else ""


# ── IMAP polling ──────────────────────────────────────────────────────────────

def get_pdf_attachments(msg) -> list[tuple[str, bytes]]:
    """Return list of (filename, bytes) for all PDF parts in an email."""
    pdfs = []
    for part in msg.walk():
        if part.get_content_type() == "application/pdf":
            raw_name = part.get_filename() or "invoice.pdf"
            filename = decode_header(raw_name)[0][0]
            if isinstance(filename, bytes):
                filename = filename.decode(errors="replace")
            pdfs.append((filename, part.get_payload(decode=True)))
    return pdfs


def process_unseen_emails():
    log.info("Connecting to IMAP…")
    with imaplib.IMAP4_SSL(IMAP_HOST) as imap:
        imap.login(IMAP_USER, IMAP_PASSWORD)
        imap.select("INBOX")

        _, message_ids = imap.search(None, "UNSEEN")
        ids = message_ids[0].split()
        log.info("%d unseen message(s) found.", len(ids))

        for msg_id in ids:
            _, msg_data = imap.fetch(msg_id, "(RFC822)")
            raw = msg_data[0][1]
            msg = email.message_from_bytes(raw)

            subject_raw = msg.get("Subject", "No subject")
            subject = decode_header(subject_raw)[0][0]
            if isinstance(subject, bytes):
                subject = subject.decode(errors="replace")

            pdfs = get_pdf_attachments(msg)

            if not pdfs:
                log.info("No PDF in message '%s' — skipping.", subject)
                # Mark as seen so we don't re-check
                imap.store(msg_id, "+FLAGS", "\\Seen")
                continue

            for filename, pdf_bytes in pdfs:
                log.info("Processing PDF: %s  (from: %s)", filename, subject)
                data = extract_invoice_data(pdf_bytes)
                create_notion_entry(data, pdf_bytes, filename, subject)

            imap.store(msg_id, "+FLAGS", "\\Seen")


def main():
    log.info("Invoice processor started. Poll interval: %ds", POLL_INTERVAL_SECONDS)
    while True:
        try:
            process_unseen_emails()
        except Exception as exc:
            log.exception("Error during poll cycle: %s", exc)
        log.info("Sleeping %d minutes…", POLL_INTERVAL_SECONDS // 60)
        time.sleep(POLL_INTERVAL_SECONDS)


if __name__ == "__main__":
    main()
