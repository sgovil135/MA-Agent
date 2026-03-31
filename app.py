import os
from fastapi import FastAPI, Request
from slack_bolt import App
from slack_bolt.adapter.fastapi import SlackRequestHandler

slack_app = App(
    token=os.environ["SLACK_BOT_TOKEN"],
    signing_secret=os.environ["SLACK_SIGNING_SECRET"]
)

handler = SlackRequestHandler(slack_app)
api = FastAPI()

@slack_app.event("app_mention")
def handle_mention(body, say):
    text = body["event"]["text"]
    say(f"I got your message: {text}")

@api.get("/")
def root():
    return {"ok": True}

@api.post("/slack/events")
async def slack_events(req: Request):
    return await handler.handle(req)
