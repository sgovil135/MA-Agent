import io
import os
import re
import requests
from fastapi import FastAPI, Request
from slack_bolt import App
from slack_bolt.adapter.fastapi import SlackRequestHandler
from openai import OpenAI

api = FastAPI()

SLACK_BOT_TOKEN = os.getenv("SLACK_BOT_TOKEN")
SLACK_SIGNING_SECRET = os.getenv("SLACK_SIGNING_SECRET")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

openai_client = OpenAI(api_key=OPENAI_API_KEY) if OPENAI_API_KEY else None

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
                model="gpt-4o-mini",
                messages=[
                    {
                        "role": "system",
                        "content": "You are a helpful assistant in Slack. Reply briefly and clearly."
                    },
                    {
                        "role": "user",
                        "content": text
                    }
                ]
            )
            answer = response.choices[0].message.content
            say(answer)

        except Exception as e:
            say(f"OpenAI error: {str(e)}")

    @slack_app.event("file_shared")
    def handle_file_shared(body, client, say):
        event = body["event"]
        file_id = event["file_id"]
        channel = event.get("channel_id")

        if not openai_client:
            say(text="OPENAI_API_KEY is missing in Railway.", channel=channel)
            return

        try:
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

            say(text=f"Reading *{file_name}*...", channel=channel)

            download_resp = requests.get(
                file_url,
                headers={"Authorization": f"Bearer {SLACK_BOT_TOKEN}"},
                timeout=60,
            )
            download_resp.raise_for_status()

            pdf_bytes = download_resp.content

            uploaded_file = openai_client.files.create(
                file=(file_name, io.BytesIO(pdf_bytes), "application/pdf"),
                purpose="user_data",
            )

            cim_prompt = (
                "You are an M&A analyst scoring a Confidential Information Memorandum (CIM).\n\n"
                "SCORING RULES:\n"
                "- Scores must be 1, 1.5, 2, 2.5, 3, 3.5, 4, 4.5, or 5 only. No other values.\n\n"
                "SCORE THESE 5 CATEGORIES:\n"
                "1. Autonomy — does real management exist without the seller?\n"
                "2. Cash Flow Quality — are margins durable and healthy?\n"
                "3. Growth Reality — is the growth path clear and practical?\n"
                "4. Downside Risk — how fragile is the revenue base?\n"
                "5. Strategic Fit — does this fit Cemtrex's specific platform? (see acquirer context below)\n\n"
                "ABOUT THE ACQUIRER (for Strategic Fit scoring):\n"
                "- Cemtrex is a public microcap platform company\n"
                "- Core divisions: AIS (industrial contracting/fabrication, roll-up strategy), "
                "Invocon (aerospace & defense engineering), Vicon (security/surveillance), "
                "VDI (simulation, private, not part of Cemtrex)\n"
                "- Acquisition filters: 30%+ gross margins, $1B+ TAM, recurring/repeat revenue, "
                "operational independence (no babysitting), must de-lever cleanly\n"
                "- Strategic priorities: tuck-ins to AIS, aerospace adjacencies via Invocon, "
                "businesses that run without the owner\n"
                "- Hard passes: broken ops, fake margin stories, key-man dependent, "
                "sectors requiring deep new expertise\n"
                "- Long-term direction: space infrastructure (fuel systems, power systems, logistics, life support)\n"
                "Score Strategic Fit against THIS specific platform — not a generic acquirer.\n"
                "Strategic Fit scoring guide:\n"
                "  5 = Direct tuck-in to AIS or Invocon, accretive immediately\n"
                "  3 = Adjacent — could bolt on but needs a thesis\n"
                "  1 = Totally standalone, new sector, would require hands-on management\n\n"
                "VERDICT:\n"
                "- 18+ → \"Go for it\"\n"
                "- 15–17 → \"Proceed Conservatively\"\n"
                "- Under 15 → \"Pass\"\n\n"
                "IMPORTANT: Use Slack formatting, NOT Markdown. Use *text* for bold (single asterisks). "
                "Do NOT use **text** or any Markdown syntax.\n\n"
                "FORMAT THE RESPONSE EXACTLY LIKE THIS:\n\n"
                "*[Company Name] — Deal Scorecard*\n\n"
                "*Business*\n"
                "[2-3 sentence description]\n\n"
                "*Financials*\n"
                "> Revenue: $X / $X / $X\n"
                "> EBITDA: $X / $X / $X\n\n"
                "*Scores*\n"
                "> Autonomy: X/5 — [one sentence]\n"
                "> Cash Flow Quality: X/5 — [one sentence]\n"
                "> Growth Reality: X/5 — [one sentence]\n"
                "> Downside Risk: X/5 — [one sentence]\n"
                "> Strategic Fit: X/5 — [one sentence]\n\n"
                "*Total: X/25 — [VERDICT IN CAPS]*\n\n"
                "*Valuation Range*\n"
                "$XM – $XM ([X–Xx EBITDA])\n\n"
                "*Open Questions*\n"
                "1. [question]\n"
                "2. [question]\n"
                "3. [question]\n\n"
                "Keep it under 400 words. No fluff. If data is missing, say so and flag it as a risk."
            )

            response = openai_client.responses.create(
                model="gpt-4.1-mini",
                input=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "input_file",
                                "file_id": uploaded_file.id,
                            },
                            {
                                "type": "input_text",
                                "text": cim_prompt,
                            },
                        ],
                    }
                ],
            )

            say(text=response.output_text, channel=channel)

        except Exception as e:
            say(text=f"PDF analysis error: {str(e)}", channel=channel)


@api.get("/")
def root():
    return {
        "ok": True,
        "has_bot_token": bool(SLACK_BOT_TOKEN),
        "has_signing_secret": bool(SLACK_SIGNING_SECRET),
        "has_openai_key": bool(OPENAI_API_KEY),
    }


@api.post("/slack/events")
async def slack_events(req: Request):
    if handler is None:
        return {"ok": False, "error": "Slack env vars missing"}
    return await handler.handle(req)
