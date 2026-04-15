import io
import json
import os
import re
from datetime import datetime, timezone

import requests
from fastapi import FastAPI, Request
from slack_bolt import App
from slack_bolt.adapter.fastapi import SlackRequestHandler
from openai import OpenAI

from pipeline import run_deal_pipeline
from outlook_poller import start_polling, init_email_db
from dropbox_monitor import start_monitoring

api = FastAPI()

SLACK_BOT_TOKEN = os.getenv("SLACK_BOT_TOKEN")
SLACK_SIGNING_SECRET = os.getenv("SLACK_SIGNING_SECRET")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
SLACK_CHANNEL_ID = os.getenv("SLACK_CHANNEL_ID", "")

openai_client = OpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY else None

# ── Conversation state ────────────────────────────────────────────────────
# Key: (channel_id, user_id) → deal pipeline result dict
active_cim_sessions = {}


# ── Slack helpers ─────────────────────────────────────────────────────────

def slack_post(text: str, channel: str = ""):
    """Post a message to Slack. Used by the Outlook poller."""
    ch = channel or SLACK_CHANNEL_ID
    if not ch or not SLACK_BOT_TOKEN:
        print(f"[SLACK] No channel configured. Message: {text}")
        return
    try:
        requests.post(
            "https://slack.com/api/chat.postMessage",
            headers={"Authorization": f"Bearer {SLACK_BOT_TOKEN}"},
            json={"channel": ch, "text": text},
            timeout=10,
        )
    except Exception as e:
        print(f"[SLACK] Post failed: {e}")


# ── Outlook poller callback ──────────────────────────────────────────────

def _outlook_pipeline_callback(email_data: dict, pdf_bytes: bytes, pdf_name: str):
    """Called by the Outlook poller when a deal email with a CIM PDF is found."""
    if not openai_client:
        slack_post("⚠️ OPENAI_API_KEY not set — cannot process CIM.")
        return

    subject = email_data.get("subject", "")
    sender = email_data.get("sender_email", "")
    slack_post(f"📧 Processing CIM from *{sender}*: *{subject}* ({pdf_name})...")

    run_deal_pipeline(
        openai_client=openai_client,
        pdf_bytes=pdf_bytes,
        pdf_name=pdf_name,
        slack_say=slack_post,
    )


def _outlook_slack_callback(text: str):
    """Called by the Outlook poller for Slack notifications."""
    slack_post(text)


# ── Startup ───────────────────────────────────────────────────────────────

@api.on_event("startup")
def startup_event():
    init_email_db()

    # Start Outlook poller if Graph credentials are configured
    has_graph = all([
        os.getenv("MICROSOFT_CLIENT_ID"),
        os.getenv("MICROSOFT_CLIENT_SECRET"),
        os.getenv("MICROSOFT_TENANT_ID"),
    ])
    if has_graph and openai_client:
        try:
            start_polling(openai_client, _outlook_pipeline_callback, _outlook_slack_callback)
        except Exception as e:
            print(f"[STARTUP] Outlook poller failed to start: {e}")
    else:
        print("[STARTUP] Outlook poller not started (missing Graph credentials or OpenAI key)")

    # Start Dropbox monitor if configured
    if os.getenv("DROPBOX_REFRESH_TOKEN"):
        try:
            start_monitoring(_outlook_slack_callback)
        except Exception as e:
            print(f"[STARTUP] Dropbox monitor failed to start: {e}")
    else:
        print("[STARTUP] Dropbox monitor not started (missing DROPBOX_REFRESH_TOKEN)")


# ── Slack app ─────────────────────────────────────────────────────────────

slack_app = None
handler = None

if SLACK_BOT_TOKEN and SLACK_SIGNING_SECRET:
    slack_app = App(
        token=SLACK_BOT_TOKEN,
        signing_secret=SLACK_SIGNING_SECRET
    )
    handler = SlackRequestHandler(slack_app)

    @slack_app.event("app_mention")
    def handle_mention(body, say):
        raw_text = body["event"].get("text", "")
        text = re.sub(r"<@\w+>", "", raw_text).strip()

        if not openai_client:
            say("OPENAI_API_KEY is missing in Railway.")
            return

        try:
            response = openai_client.chat.completions.create(
                model="gpt-4.1-mini",
                messages=[
                    {"role": "system", "content": "You are a helpful M&A assistant in Slack. Reply briefly and clearly. Use Slack formatting (*bold* not **bold**)."},
                    {"role": "user", "content": text}
                ]
            )
            say(response.choices[0].message.content)
        except Exception as e:
            say(f"OpenAI error: {str(e)}")

    @slack_app.event("file_shared")
    def handle_file_shared(body, client, say):
        """Handle CIM PDF uploads via Slack — runs the full deal pipeline."""
        event = body["event"]
        file_id = event["file_id"]
        channel = event.get("channel_id")
        user_id = event.get("user_id", "")
        file_name = "(unknown file)"

        try:
            if not openai_client:
                say(text="OPENAI_API_KEY is missing in Railway.", channel=channel)
                return

            file_info = client.files_info(file=file_id)
            file_obj = file_info["file"]
            file_name = file_obj["name"]
            mimetype = file_obj.get("mimetype", "")
            file_url = file_obj.get("url_private")

            if mimetype != "application/pdf":
                say(text=f"Got your file: *{file_name}*. Right now I only handle PDFs.", channel=channel)
                return

            if not file_url:
                say(text=f"Could not access private URL for *{file_name}*.", channel=channel)
                return

            say(text=f"Reading *{file_name}* — running full deal pipeline...", channel=channel)

            # Download PDF from Slack
            download_resp = requests.get(
                file_url,
                headers={"Authorization": f"Bearer {SLACK_BOT_TOKEN}"},
                timeout=60,
            )
            download_resp.raise_for_status()
            pdf_bytes = download_resp.content

            # Run the pipeline — Score → Excel → Diligence → LOI → Dropbox
            def channel_say(text):
                say(text=text, channel=channel)

            result = run_deal_pipeline(
                openai_client=openai_client,
                pdf_bytes=pdf_bytes,
                pdf_name=file_name,
                slack_say=channel_say,
            )

            # Store session for follow-up Q&A
            company = result.get("company_name", "Unknown")
            scorecard = result.get("scorecard", {})
            active_cim_sessions[(channel, user_id)] = {
                "file_id": result.get("openai_file_id", ""),
                "scorecard": json.dumps(scorecard, indent=2),
                "company": company,
                "pipeline_result": result,
                "history": [],
            }

        except Exception as e:
            print(f"[FILE_SHARED ERROR] {e}")
            say(text=f"*Something went wrong processing {file_name} — here's the error: {str(e)}*", channel=channel)

    @slack_app.event("message")
    def handle_message(body, say):
        event = body.get("event", {})
        if event.get("bot_id") or event.get("subtype"):
            return

        channel = event.get("channel", "")
        user_id = event.get("user", "")
        text = event.get("text", "").strip()
        key = (channel, user_id)

        if not text:
            return

        try:
            text_lower = text.lower()

            # "done" ends CIM Q&A session
            if text_lower == "done" and key in active_cim_sessions:
                company = active_cim_sessions[key].get("company", "the CIM")
                del active_cim_sessions[key]
                say(text=f"Closed the *{company}* session. Upload a new CIM whenever you're ready.", channel=channel)
                return

            # ── CIM follow-up Q&A ─────────────────────────────────────
            if key in active_cim_sessions and openai_client:
                session = active_cim_sessions[key]
                company = session.get("company", "this company")

                # Detect email drafting requests
                is_email = any(w in text_lower for w in ["draft", "email", "write to", "respond to", "reply to", "reach out"])

                system_msg = (
                    f"You have access to the full CIM for {company}. "
                    f"The scorecard produced was:\n\n{session['scorecard']}\n\n"
                    "Answer follow-up questions about this deal using the CIM content. "
                    "Be specific, cite numbers from the CIM when possible. "
                    "Use Slack formatting (*bold* not **bold**)."
                )
                if is_email:
                    system_msg += (
                        "\n\nThe user wants you to draft an email or message. "
                        "Write in a direct, no-fluff CEO style appropriate for M&A outreach. "
                        "Keep it professional but not corporate-speak. Short sentences. "
                        "Get to the point fast. Sound like a principal, not an advisor."
                    )

                input_content = [
                    {"type": "input_file", "file_id": session["file_id"]},
                ]

                response = openai_client.responses.create(
                    model="gpt-4.1-mini",
                    input=[
                        {"role": "user", "content": input_content + [{"type": "input_text", "text": system_msg}]},
                        *session["history"][-10:],
                        {"role": "user", "content": [{"type": "input_text", "text": text}]},
                    ],
                )

                answer = response.output_text
                say(text=answer, channel=channel)

                session["history"].append({"role": "user", "content": [{"type": "input_text", "text": text}]})
                session["history"].append({"role": "assistant", "content": [{"type": "output_text", "text": answer}]})

        except Exception as e:
            print(f"[MESSAGE ERROR] {e}")
            say(text=f"*Something went wrong — here's the error: {str(e)}*", channel=channel)


# ── FastAPI routes ────────────────────────────────────────────────────────

@api.get("/")
def root():
    return {
        "ok": True,
        "service": "Cemtrex M&A Deal Bot",
        "has_bot_token": bool(SLACK_BOT_TOKEN),
        "has_signing_secret": bool(SLACK_SIGNING_SECRET),
        "has_openai_key": bool(OPENAI_API_KEY),
        "has_dropbox": bool(os.getenv("DROPBOX_REFRESH_TOKEN")),
        "has_graph": all([
            os.getenv("MICROSOFT_CLIENT_ID"),
            os.getenv("MICROSOFT_CLIENT_SECRET"),
            os.getenv("MICROSOFT_TENANT_ID"),
        ]),
        "outlook_folder": os.getenv("OUTLOOK_FOLDER_NAME", "Acquisitions/2026"),
    }


@api.post("/slack/events")
async def slack_events(req: Request):
    if handler is None:
        return {"ok": False, "error": "Slack env vars missing"}
    return await handler.handle(req)
