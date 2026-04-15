"""
Outlook email poller via Microsoft Graph API.
Polls a configured mail folder for new deal emails every 15 minutes.
"""

from __future__ import annotations

import json
import os
import sqlite3
import time
import threading
from datetime import datetime, timezone
from typing import Any

import requests

from prompts import CLASSIFY_EMAIL_PROMPT

# ── Config ────────────────────────────────────────────────────────────────

MICROSOFT_CLIENT_ID = os.getenv("MICROSOFT_CLIENT_ID")
MICROSOFT_CLIENT_SECRET = os.getenv("MICROSOFT_CLIENT_SECRET")
MICROSOFT_TENANT_ID = os.getenv("MICROSOFT_TENANT_ID")
OUTLOOK_USER_EMAIL = os.getenv("OUTLOOK_USER_EMAIL", "sgovil@cemtrex.com")
OUTLOOK_FOLDER_NAME = os.getenv("OUTLOOK_FOLDER_NAME", "Acquisitions/2026")
POLL_INTERVAL_SECONDS = int(os.getenv("POLL_INTERVAL_SECONDS", "900"))  # 15 min

DB_PATH = os.getenv("MA_DB_PATH", "./ma_copilot.db")


# ── Graph Auth ────────────────────────────────────────────────────────────

def get_graph_token() -> str:
    url = f"https://login.microsoftonline.com/{MICROSOFT_TENANT_ID}/oauth2/v2.0/token"
    resp = requests.post(url, data={
        "grant_type": "client_credentials",
        "client_id": MICROSOFT_CLIENT_ID,
        "client_secret": MICROSOFT_CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
    }, timeout=30)
    resp.raise_for_status()
    return resp.json()["access_token"]


def _graph_headers() -> dict[str, str]:
    return {"Authorization": f"Bearer {get_graph_token()}"}


# ── Folder resolution ─────────────────────────────────────────────────────

def resolve_folder_id(folder_path: str) -> str:
    """Resolve a nested folder path like 'Acquisitions/2026' under Inbox.

    Returns the Graph folder ID.
    """
    headers = _graph_headers()
    parts = folder_path.strip("/").split("/")

    # Start from Inbox
    url = f"https://graph.microsoft.com/v1.0/users/{OUTLOOK_USER_EMAIL}/mailFolders/Inbox/childFolders"
    current_id = None

    for part in parts:
        resp = requests.get(url, headers=headers, timeout=15)
        resp.raise_for_status()
        folders = resp.json().get("value", [])
        matched = None
        for f in folders:
            if f["displayName"].lower() == part.lower():
                matched = f
                break
        if not matched:
            raise ValueError(f"Mail folder '{part}' not found (resolving '{folder_path}')")
        current_id = matched["id"]
        url = f"https://graph.microsoft.com/v1.0/users/{OUTLOOK_USER_EMAIL}/mailFolders/{current_id}/childFolders"

    return current_id


# ── DB for processed email tracking ──────────────────────────────────────

def init_email_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS processed_emails (
            message_id TEXT PRIMARY KEY,
            subject TEXT,
            sender TEXT,
            processed_at TEXT,
            is_deal INTEGER,
            deal_name TEXT
        )
    """)
    conn.commit()
    conn.close()


def is_email_processed(message_id: str) -> bool:
    conn = sqlite3.connect(DB_PATH)
    row = conn.execute(
        "SELECT 1 FROM processed_emails WHERE message_id = ?", (message_id,)
    ).fetchone()
    conn.close()
    return row is not None


def mark_email_processed(message_id: str, subject: str, sender: str,
                          is_deal: bool, deal_name: str = ""):
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        """INSERT OR IGNORE INTO processed_emails
           (message_id, subject, sender, processed_at, is_deal, deal_name)
           VALUES (?, ?, ?, ?, ?, ?)""",
        (message_id, subject, sender,
         datetime.now(timezone.utc).isoformat(), int(is_deal), deal_name),
    )
    conn.commit()
    conn.close()


# ── Email fetching ────────────────────────────────────────────────────────

def fetch_recent_emails(folder_id: str, top: int = 20) -> list[dict[str, Any]]:
    """Fetch the most recent emails from the folder."""
    headers = _graph_headers()
    url = (
        f"https://graph.microsoft.com/v1.0/users/{OUTLOOK_USER_EMAIL}"
        f"/mailFolders/{folder_id}/messages"
        f"?$top={top}&$orderby=receivedDateTime desc"
        f"&$select=id,subject,from,receivedDateTime,body,hasAttachments"
    )
    resp = requests.get(url, headers=headers, timeout=30)
    resp.raise_for_status()
    return resp.json().get("value", [])


def fetch_attachments(message_id: str) -> list[dict[str, Any]]:
    """Fetch attachments for a specific email."""
    headers = _graph_headers()
    url = (
        f"https://graph.microsoft.com/v1.0/users/{OUTLOOK_USER_EMAIL}"
        f"/messages/{message_id}/attachments"
    )
    resp = requests.get(url, headers=headers, timeout=30)
    resp.raise_for_status()
    return resp.json().get("value", [])


def download_attachment(message_id: str, attachment_id: str) -> bytes:
    """Download attachment content."""
    headers = _graph_headers()
    url = (
        f"https://graph.microsoft.com/v1.0/users/{OUTLOOK_USER_EMAIL}"
        f"/messages/{message_id}/attachments/{attachment_id}/$value"
    )
    resp = requests.get(url, headers=headers, timeout=60)
    resp.raise_for_status()
    return resp.content


def extract_email_data(msg: dict) -> dict[str, Any]:
    """Extract useful fields from a Graph API message object."""
    sender_data = msg.get("from", {}).get("emailAddress", {})
    return {
        "message_id": msg["id"],
        "subject": msg.get("subject", "(no subject)"),
        "sender_name": sender_data.get("name", ""),
        "sender_email": sender_data.get("address", ""),
        "received_at": msg.get("receivedDateTime", ""),
        "body_text": msg.get("body", {}).get("content", ""),
        "has_attachments": msg.get("hasAttachments", False),
    }


def find_pdf_attachments(message_id: str) -> list[dict[str, Any]]:
    """Find PDF attachments on a message. Returns list of {name, id, size}."""
    attachments = fetch_attachments(message_id)
    pdfs = []
    for att in attachments:
        name = att.get("name", "")
        content_type = att.get("contentType", "")
        if name.lower().endswith(".pdf") or "pdf" in content_type.lower():
            pdfs.append({
                "name": name,
                "id": att["id"],
                "size": att.get("size", 0),
                "content_bytes": att.get("contentBytes"),  # base64 for small attachments
            })
    return pdfs


# ── Classify email ────────────────────────────────────────────────────────

def _get_active_deals() -> list[str]:
    """Get list of active deal names from Dropbox folders + processed emails."""
    deals = set()

    # From Dropbox folders
    try:
        import dropbox_client
        folders = dropbox_client.list_deal_folders()
        deals.update(folders)
    except Exception as e:
        print(f"[CLASSIFIER] Could not list Dropbox folders: {e}")

    # From processed emails that were flagged as deals
    try:
        conn = sqlite3.connect(DB_PATH)
        rows = conn.execute(
            "SELECT DISTINCT deal_name FROM processed_emails WHERE is_deal = 1 AND deal_name != ''"
        ).fetchall()
        conn.close()
        for row in rows:
            deals.add(row[0])
    except Exception:
        pass

    return sorted(deals)


def classify_email(openai_client, subject: str, sender: str, body: str,
                   attachment_names: list[str] | None = None) -> dict:
    """Classify an email into: new_cim, existing_deal_update, new_deal_no_cim, noise."""
    body_truncated = body[:2000] if body else ""
    active_deals = _get_active_deals()
    deals_str = "\n".join(f"- {d}" for d in active_deals) if active_deals else "(none tracked yet)"
    attachments_str = "\n".join(f"- {a}" for a in (attachment_names or [])) or "(no attachments)"

    prompt = CLASSIFY_EMAIL_PROMPT.replace(
        "__SUBJECT__", subject
    ).replace(
        "__SENDER__", sender
    ).replace(
        "__BODY__", body_truncated
    ).replace(
        "__ACTIVE_DEALS__", deals_str
    ).replace(
        "__ATTACHMENTS__", attachments_str
    )

    response = openai_client.chat.completions.create(
        model="gpt-4.1-mini",
        messages=[{"role": "user", "content": prompt}],
        temperature=0,
    )
    text = response.choices[0].message.content.strip()

    if text.startswith("```"):
        text = text.split("\n", 1)[1] if "\n" in text else text[3:]
    if text.endswith("```"):
        text = text[:-3]
    text = text.strip()

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return {"category": "noise", "confidence": "low", "reason": "Failed to parse classification"}


# ── Polling loop ──────────────────────────────────────────────────────────

def start_polling(openai_client, pipeline_callback, slack_say_callback):
    """Start the background polling thread.

    Args:
        openai_client: OpenAI client instance
        pipeline_callback: function(email_data, pdf_bytes, pdf_name) -> None
            Called when a NEW deal email with a CIM PDF is found.
        slack_say_callback: function(text) -> None
            Called to send Slack notifications.
    """
    init_email_db()

    def _poll_loop():
        folder_id = None
        while True:
            try:
                if folder_id is None:
                    folder_id = resolve_folder_id(OUTLOOK_FOLDER_NAME)
                    print(f"[POLLER] Resolved folder '{OUTLOOK_FOLDER_NAME}' → {folder_id[:30]}...")

                emails = fetch_recent_emails(folder_id, top=20)
                new_count = 0

                for msg in emails:
                    email_data = extract_email_data(msg)
                    mid = email_data["message_id"]

                    if is_email_processed(mid):
                        continue

                    new_count += 1
                    subject = email_data["subject"]
                    sender = email_data["sender_email"]
                    sender_name = email_data["sender_name"]
                    body = email_data["body_text"]

                    # Get attachment names for classification context
                    attachment_names = []
                    if email_data["has_attachments"]:
                        try:
                            atts = fetch_attachments(mid)
                            attachment_names = [a.get("name", "") for a in atts]
                        except Exception:
                            pass

                    # Classify with active deal awareness
                    classification = classify_email(
                        openai_client, subject, sender, body, attachment_names
                    )
                    category = classification.get("category", "noise")
                    matched_deal = classification.get("matched_deal")
                    reason = classification.get("reason", "")

                    print(f"[POLLER] {category}: {subject} — {reason}")

                    if category == "noise":
                        mark_email_processed(mid, subject, sender, is_deal=False)
                        continue

                    if category == "existing_deal_update":
                        # Notify about update to existing deal, don't re-score
                        deal_name = matched_deal or subject
                        att_str = ", ".join(attachment_names) if attachment_names else "no attachments"
                        slack_say_callback(
                            f"📎 *Update for {deal_name}*\n"
                            f"From: {sender_name or sender}\n"
                            f"Subject: {subject}\n"
                            f"Attachments: {att_str}"
                        )
                        mark_email_processed(mid, subject, sender,
                                              is_deal=True, deal_name=deal_name)
                        continue

                    if category == "new_deal_no_cim":
                        slack_say_callback(
                            f"📬 New deal email from *{sender_name or sender}*: "
                            f"*{subject}*. No CIM attached."
                        )
                        mark_email_processed(mid, subject, sender,
                                              is_deal=True, deal_name="")
                        continue

                    if category == "new_cim":
                        # New deal with CIM — find and process the PDF
                        pdfs = find_pdf_attachments(mid) if email_data["has_attachments"] else []

                        if pdfs:
                            pdf = pdfs[0]
                            if pdf.get("content_bytes"):
                                import base64
                                pdf_bytes = base64.b64decode(pdf["content_bytes"])
                            else:
                                pdf_bytes = download_attachment(mid, pdf["id"])

                            print(f"[POLLER] New CIM: {subject} → {pdf['name']}")
                            try:
                                pipeline_callback(email_data, pdf_bytes, pdf["name"])
                                mark_email_processed(mid, subject, sender,
                                                      is_deal=True, deal_name=subject)
                            except Exception as e:
                                print(f"[POLLER] Pipeline error for {subject}: {e}")
                                slack_say_callback(
                                    f"⚠️ Pipeline error processing *{subject}*: {str(e)}"
                                )
                                mark_email_processed(mid, subject, sender, is_deal=True)
                        else:
                            # Classified as new_cim but no PDF found
                            slack_say_callback(
                                f"📬 New deal from *{sender_name or sender}*: "
                                f"*{subject}*. Classified as CIM but no PDF found."
                            )
                            mark_email_processed(mid, subject, sender,
                                                  is_deal=True, deal_name="")
                        continue

                    # Fallback — unknown category
                    mark_email_processed(mid, subject, sender, is_deal=False)

                if new_count:
                    print(f"[POLLER] Processed {new_count} new emails")
                else:
                    print(f"[POLLER] No new emails")

            except Exception as e:
                print(f"[POLLER] Error: {e}")
                folder_id = None  # Re-resolve on next attempt

            time.sleep(POLL_INTERVAL_SECONDS)

    thread = threading.Thread(target=_poll_loop, daemon=True, name="outlook-poller")
    thread.start()
    print(f"[POLLER] Started — polling '{OUTLOOK_FOLDER_NAME}' every {POLL_INTERVAL_SECONDS}s")
    return thread
