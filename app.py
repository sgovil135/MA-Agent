import os
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
        text = body["event"]["text"]

        if openai_client is None:
            say("OPENAI_API_KEY is missing in Railway.")
            return

        response = openai_client.responses.create(
            model="gpt-4.1-mini",
            input=f"Summarize this briefly: {text}"
        )

        answer = response.output_text
        say(answer)

    @slack_app.event("file_shared")
    def handle_file_shared(body, client, say):
        file_id = body["event"]["file_id"]
        file_info = client.files_info(file=file_id)
        file_name = file_info["file"]["name"]

        if openai_client is None:
            say(f"Got your file: {file_name}. But OPENAI_API_KEY is missing in Railway.")
            return

        say(f"Got your file: {file_name}. AI processing comes next.")

@api.get("/")
def root():
    return {
        "ok": True,
        "has_bot_token": bool(SLACK_BOT_TOKEN),
        "has_signing_secret": bool(SLACK_SIGNING_SECRET),
        "has_openai_key": bool(OPENAI_API_KEY),
        "openai_key_prefix": OPENAI_API_KEY[:7] if OPENAI_API_KEY else None,
    }

@api.post("/slack/events")
async def slack_events(req: Request):
    if handler is None:
        return {"ok": False, "error": "Slack env vars missing"}
    return await handler.handle(req)
