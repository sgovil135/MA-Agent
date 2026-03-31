import os
from fastapi import FastAPI, Request
from slack_bolt import App
from slack_bolt.adapter.fastapi import SlackRequestHandler

api = FastAPI()

SLACK_BOT_TOKEN = os.getenv("SLACK_BOT_TOKEN")
SLACK_SIGNING_SECRET = os.getenv("SLACK_SIGNING_SECRET")

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
        say(f"I got your message: {text}")

    @slack_app.event("message")
    def handle_message_events(body, say, logger):
        event = body.get("event", {})

        # ignore bot messages
        if event.get("subtype") == "bot_message":
            return

        files = event.get("files", [])
        if not files:
            return

        file_names = [f.get("name", "unknown_file") for f in files]
        say(f"Got your file(s): {', '.join(file_names)}")

@api.get("/")
def root():
    return {
        "ok": True,
        "has_bot_token": bool(SLACK_BOT_TOKEN),
        "has_signing_secret": bool(SLACK_SIGNING_SECRET),
    }

@api.post("/slack/events")
async def slack_events(req: Request):
    if handler is None:
        return {"ok": False, "error": "Slack env vars missing"}
    return await handler.handle(req)
