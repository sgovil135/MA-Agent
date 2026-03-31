import os
from fastapi import FastAPI, Request
from slack_bolt import App
from slack_bolt.adapter.fastapi import SlackRequestHandler
from openai import OpenAI

api = FastAPI()

SLACK_BOT_TOKEN = os.getenv("SLACK_BOT_TOKEN")
SLACK_SIGNING_SECRET = os.getenv("SLACK_SIGNING_SECRET")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

client = OpenAI(api_key=OPENAI_API_KEY)

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

        response = client.responses.create(
            model="gpt-4.1-mini",
            input=f"Summarize this briefly: {text}"
        )

        answer = response.output[0].content[0].text
        say(answer)

    @slack_app.event("file_shared")
    def handle_file_shared(body, client, say):
        file_id = body["event"]["file_id"]
        file_info = client.files_info(file=file_id)
        file_name = file_info["file"]["name"]

        say(f"Processing {file_name}...")

        # For now, we just simulate
        response = client.responses.create(
            model="gpt-4.1-mini",
            input=f"Give me a short M&A style summary of a company based on a CIM called {file_name}"
        )

        answer = response.output[0].content[0].text
        say(answer)

@api.get("/")
def root():
    return {"ok": True}

@api.post("/slack/events")
async def slack_events(req: Request):
    if handler is None:
        return {"ok": False}
    return await handler.handle(req)
