"""
Invoice Automation - invoice_processor.py
Polls IMAP inbox, extracts PDF data via Claude API, creates Trello cards.
Runs a Flask webhook server to receive Trello card move events and update totals.
"""

import imaplib
import email
import time
import base64
import json
import logging
import os
import threading
from datetime import datetime
from email.header import decode_header

import anthropic
import requests
from flask import Flask, request, jsonify

app = Flask(__name__)

# Logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)s  %(message)s",
)
log = logging.getLogger(__name__)

# Config
IMAP_HOST = os.environ["IMAP_HOST"]
IMAP_USER = os.environ["IMAP_USER"]
IMAP_PASSWORD = os.environ["IMAP_PASSWORD"]
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]
TRELLO_API_KEY = os.environ["TRELLO_API_KEY"]
TRELLO_TOKEN = os.environ["TRELLO_TOKEN"]
TRELLO_BOARD_ID = os.environ["TRELLO_BOARD_ID"]

POLL_INTERVAL_SECONDS = 3 * 60

claude = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

def get_full_board_id():
    """Get the full 24-char board ID from the short ID."""
    resp = requests.get(
        f"https://api.trello.com/1/boards/{TRELLO_BOARD_ID}",
        params={**TRELLO_AUTH, "fields": "id"},
    )
    resp.raise_for_status()
    return resp.json()["id"]

TRELLO_FULL_BOARD_ID = None  # Set at startup

MONTH_NAMES = {
    1: "Januar", 2: "Februar", 3: "Marts", 4: "April",
    5: "Maj", 6: "Juni", 7: "Juli", 8: "August",
    9: "September", 10: "Oktober", 11: "November", 12: "December"
}

EXTRACTION_PROMPT = """
You are an invoice data extraction assistant. Analyse the attached PDF and
return ONLY a valid JSON object with no explanation, no markdown, no code fences.

Fields to extract:
{
  "is_invoice":      true or false,
  "supplier_name":   string or null,
  "amount_dkk":      number or null,
  "due_date":        string or null,
  "payment_type":    "Auto-debit" or "Manual" or "Unknown",
  "is_kreditnota":   true or false,
  "notes":           string
}

Rules:
- is_invoice: set to true ONLY if this is an invoice (faktura), credit note (kreditnota),
  or a refund/payout specification where money is paid TO the recipient (e.g. Dansk Retursystem
  "Specifikation for udbetalt pant" — a pant refund for returned bottles/cans).
  Set to false for account statements (kontoudtog), letters, contracts, reminders without
  an invoice amount, or any other document that is not a direct payment request or refund.
- payment_type: look for keywords like betalingsservice, PBS, direct debit,
  automatisk betaling -> Auto-debit; otherwise -> Manual; if unclear -> Unknown.
- amount_dkk: always extract as a POSITIVE number regardless of how it appears.
  Convert decimal commas to decimal points (1.234,56 -> 1234.56).
  For Dansk Retursystem pant specs, use the absolute value of "I alt ekskl. moms".
- is_kreditnota: set to true if the document is a credit note (kreditnota, kreditering,
  credit note, godtgørelse), OR if it is a payout/refund document where money flows TO
  the recipient (e.g. Dansk Retursystem "Specifikation for udbetalt pant for emballager").
- due_date: use the explicit due date (forfaldsdato) if present.
  If no explicit due date, calculate it from the invoice date + payment terms.
  Examples: "30 dage netto" = invoice date + 30 days, "14 dage netto" = invoice date + 14 days,
  "netto 8 dage" = invoice date + 8 days. For pant refunds, use the bogføringsdato as the date.
  If neither due date nor terms are found, set to null.
- due_date format: YYYY-MM-DD
- supplier_name: prefer legal entity name. For Dansk Retursystem documents, use "Dansk Retursystem".
- Return ONLY the JSON object, nothing else.
"""


def extract_invoice_data(pdf_bytes):
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

    if "```" in raw:
        parts = raw.split("```")
        for part in parts:
            part = part.strip()
            if part.startswith("json"):
                part = part[4:].strip()
            try:
                return json.loads(part)
            except json.JSONDecodeError:
                continue

    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        log.warning("Claude returned non-JSON: %s", raw)
        return {
            "supplier_name": None,
            "amount_dkk": None,
            "due_date": None,
            "payment_type": "Unknown",
            "notes": "Extraction failed: " + raw[:200],
        }


# Trello helpers

TRELLO_AUTH = {"key": TRELLO_API_KEY, "token": TRELLO_TOKEN}


def get_board_lists():
    resp = requests.get(
        f"https://api.trello.com/1/boards/{TRELLO_BOARD_ID}/lists",
        params=TRELLO_AUTH,
    )
    resp.raise_for_status()
    return {lst["name"]: lst["id"] for lst in resp.json()}


def get_board_labels():
    resp = requests.get(
        f"https://api.trello.com/1/boards/{TRELLO_BOARD_ID}/labels",
        params=TRELLO_AUTH,
    )
    resp.raise_for_status()
    return {lbl["name"]: lbl["id"] for lbl in resp.json() if lbl.get("name")}


def ensure_list(lists, name):
    if name in lists:
        return lists[name]
    resp = requests.post(
        f"https://api.trello.com/1/boards/{TRELLO_BOARD_ID}/lists",
        params={**TRELLO_AUTH, "name": name},
    )
    resp.raise_for_status()
    new_id = resp.json()["id"]
    lists[name] = new_id
    log.info("Created Trello list: %s", name)
    return new_id


def ensure_label(labels, name, color):
    if name in labels:
        return labels[name]
    resp = requests.post(
        "https://api.trello.com/1/labels",
        params={**TRELLO_AUTH, "name": name, "color": color, "idBoard": TRELLO_FULL_BOARD_ID},
    )
    if not resp.ok:
        log.warning("Label creation failed: %s — skipping label", resp.text)
        return None
    new_id = resp.json()["id"]
    labels[name] = new_id
    log.info("Created Trello label: %s", name)
    return new_id


def build_card_description(data, email_subject):
    lines = []

    amount = data.get("amount_dkk")
    if amount is not None:
        lines.append(f"**Beloeb:** kr. {amount:,.2f}")

    due = data.get("due_date")
    if due:
        try:
            due_fmt = datetime.strptime(due, "%Y-%m-%d").strftime("%d.%m.%Y")
            lines.append(f"**Forfaldsdato:** {due_fmt}")
        except ValueError:
            lines.append(f"**Forfaldsdato:** {due}")
    else:
        lines.append("**Forfaldsdato:** Ikke fundet - saet manuelt")

    payment = data.get("payment_type", "Unknown")
    lines.append(f"**Betalingstype:** {payment}")
    lines.append(f"**Email:** {email_subject}")

    notes = data.get("notes", "")
    if notes:
        lines.append(f"\nNoter: {notes}")

    if payment == "Unknown":
        lines.append("Betalingstype ukendt - tjek om Auto-debit eller Manuel")

    return "\n".join(lines)


def update_total_card(lists, list_name, list_id):
    try:
        resp = requests.get(
            f"https://api.trello.com/1/lists/{list_id}/cards",
            params=TRELLO_AUTH,
        )
        resp.raise_for_status()
        cards = resp.json()

        total = 0.0
        total_card_id = None

        for card in cards:
            if card["name"].startswith("TOTAL:"):
                total_card_id = card["id"]
            else:
                desc = card.get("desc", "")
                for line in desc.split("\n"):
                    if "**Beloeb:**" in line:
                        try:
                            amount_str = line.split("**Beloeb:**")[1].strip().replace(",", "").replace(" kr.", "").replace("kr. ", "")
                            total += float(amount_str)
                        except (ValueError, IndexError):
                            pass

        total_text = f"TOTAL: kr. {total:,.2f}"

        if total_card_id:
            requests.put(
                f"https://api.trello.com/1/cards/{total_card_id}",
                params={**TRELLO_AUTH, "name": total_text},
            )
        else:
            requests.post(
                "https://api.trello.com/1/cards",
                params={
                    **TRELLO_AUTH,
                    "name": total_text,
                    "idList": list_id,
                    "pos": "top",
                },
            )
        log.info("Updated total for %s: %s", list_name, total_text)
    except Exception as e:
        log.warning("Could not update total card: %s", e)


def create_trello_card(data, pdf_bytes, filename, email_subject):
    if not data.get("is_invoice", True):
        log.info("Skipping non-invoice PDF (%s): %s", filename, email_subject)
        return

    lists = get_board_lists()
    labels = get_board_labels()

    supplier = data.get("supplier_name") or ""
    is_drs = "dansk retursystem" in supplier.lower()

    due_date = None  # FIX: initialise before branching so it's always defined

    if is_drs:
        list_name = "DRS"
        list_id = ensure_list(lists, "DRS")
    else:
        due_date = data.get("due_date")
        if due_date:
            try:
                due = datetime.strptime(due_date, "%Y-%m-%d")
                list_name = MONTH_NAMES[due.month]
            except (ValueError, KeyError):
                list_name = "Ingen dato"
        else:
            list_name = "Ingen dato"

        list_id = ensure_list(lists, list_name)

    auto_debit_label = ensure_label(labels, "Auto-debit", "blue")
    manual_label = ensure_label(labels, "Manuel", "red")
    kreditnota_label = ensure_label(labels, "Kreditnota", "green")
    ensure_label(labels, "Rykker", "yellow")

    is_kreditnota = data.get("is_kreditnota", False)
    payment_type = data.get("payment_type", "Unknown")
    label_ids = []

    if is_kreditnota:
        if kreditnota_label:
            label_ids.append(kreditnota_label)
    else:
        if payment_type == "Auto-debit" and auto_debit_label:
            label_ids.append(auto_debit_label)
        elif payment_type == "Manual" and manual_label:
            label_ids.append(manual_label)

    supplier = data.get("supplier_name") or "Ukendt leverandoer"
    amount = data.get("amount_dkk")

    if is_kreditnota:
        if amount is not None:
            data["amount_dkk"] = -abs(amount)
            card_name = f"{supplier} - KREDIT -kr. {amount:,.2f}"
        else:
            card_name = f"{supplier} - KREDIT"
    else:
        if amount is not None:
            card_name = f"{supplier} - kr. {amount:,.2f}"
        else:
            card_name = supplier

    desc = build_card_description(data, email_subject)

    params = {
        **TRELLO_AUTH,
        "name": card_name,
        "desc": desc,
        "idList": list_id,
        "idLabels": ",".join(label_ids),
    }

    if due_date:
        params["due"] = due_date + "T12:00:00.000Z"

    resp = requests.post("https://api.trello.com/1/cards", params=params)
    if not resp.ok:
        log.error("Trello card creation failed: %s", resp.text)
        return

    card = resp.json()
    card_id = card["id"]
    log.info("Trello card created: %s in list %s", card_name, list_name)

    if pdf_bytes:
        try:
            attach_resp = requests.post(
                f"https://api.trello.com/1/cards/{card_id}/attachments",
                params=TRELLO_AUTH,
                files={"file": (filename, pdf_bytes, "application/pdf")},
            )
            if attach_resp.ok:
                log.info("PDF attached: %s", filename)
            else:
                log.warning("PDF attach failed: %s", attach_resp.text)
        except Exception as e:
            log.warning("PDF attach error: %s", e)

    if not is_drs:
        update_total_card(lists, list_name, list_id)


# Email polling

def get_pdf_attachments(msg):
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
    log.info("Connecting to IMAP...")
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
                log.info("No PDF in message '%s' - skipping.", subject)
                imap.store(msg_id, "+FLAGS", "\\Seen")
                continue

            for filename, pdf_bytes in pdfs:
                log.info("Processing PDF: %s (from: %s)", filename, subject)
                data = extract_invoice_data(pdf_bytes)
                create_trello_card(data, pdf_bytes, filename, subject)

            imap.store(msg_id, "+FLAGS", "\\Seen")


def recalculate_all_totals():
    """Recalculate totals for all lists on the board."""
    try:
        lists = get_board_lists()
        action_lists = {"Betalt", "Bogført"}
        for list_name, list_id in lists.items():
            if list_name not in action_lists and not list_name.startswith("TOTAL"):
                update_total_card(lists, list_name, list_id)
    except Exception as e:
        log.warning("Error recalculating totals: %s", e)


@app.route("/webhook", methods=["GET", "POST"])
def webhook():
    if request.method == "GET":
        return "", 200
    try:
        payload = request.get_json(silent=True) or {}
        action = payload.get("action", {})
        action_type = action.get("type", "")
        if action_type == "updateCard":
            data = action.get("data", {})
            card = data.get("card", {})
            if "listAfter" in data or "listBefore" in data or "closed" in card:
                log.info("Card moved or archived — recalculating totals")
                threading.Thread(target=recalculate_all_totals, daemon=True).start()
    except Exception as e:
        log.warning("Webhook error: %s", e)
    return "", 200


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"}), 200


def polling_loop():
    while True:
        try:
            process_unseen_emails()
        except Exception as exc:
            log.exception("Error during poll cycle: %s", exc)
        log.info("Sleeping %d minutes...", POLL_INTERVAL_SECONDS // 60)
        time.sleep(POLL_INTERVAL_SECONDS)


def register_trello_webhook(public_url):
    """Register Trello webhook if not already registered."""
    try:
        callback_url = f"{public_url}/webhook"
        resp = requests.get(
            "https://api.trello.com/1/tokens/" + TRELLO_TOKEN + "/webhooks",
            params=TRELLO_AUTH,
        )
        if resp.ok:
            for wh in resp.json():
                if wh.get("callbackURL") == callback_url:
                    log.info("Webhook already registered: %s", callback_url)
                    return
        resp = requests.post(
            "https://api.trello.com/1/webhooks",
            params={
                **TRELLO_AUTH,
                "callbackURL": callback_url,
                "idModel": TRELLO_FULL_BOARD_ID,
                "description": "Invoice Totals",
            },
        )
        if resp.ok:
            log.info("Webhook registered: %s", callback_url)
        else:
            log.warning("Webhook registration failed: %s", resp.text)
    except Exception as e:
        log.warning("Webhook registration error: %s", e)


def main():
    global TRELLO_FULL_BOARD_ID
    TRELLO_FULL_BOARD_ID = get_full_board_id()
    log.info("Invoice processor started (Trello). Board ID: %s. Poll interval: %ds", TRELLO_FULL_BOARD_ID, POLL_INTERVAL_SECONDS)

    public_url = os.environ.get("RENDER_EXTERNAL_URL", "")
    if public_url:
        register_trello_webhook(public_url)
    else:
        log.warning("RENDER_EXTERNAL_URL not set - webhook not registered")

    t = threading.Thread(target=polling_loop, daemon=True)
    t.start()

    port = int(os.environ.get("PORT", 10000))
    log.info("Starting webhook server on port %d", port)
    app.run(host="0.0.0.0", port=port)


if __name__ == "__main__":
    main()
