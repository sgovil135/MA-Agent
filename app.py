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
                                "text": (
                                    "Analyze this CIM and return:\n"
                                    "1. What the business actually does in 2 sentences\n"
                                    "2. Revenue and EBITDA for the last 3 years if available\n"
                                    "3. Top 3 positives\n"
                                    "4. Top 3 negatives\n"
                                    "5. Biggest red flag\n"
                                    "6. 5 diligence questions\n"
                                    "Be concise and skeptical."
                                ),
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
