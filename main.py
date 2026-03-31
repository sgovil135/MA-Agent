from __future__ import annotations

import json
import os
import re
import sqlite3
import uuid
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Literal

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from pydantic import BaseModel, Field

# ============================================================
# Cemtrex M&A Copilot MVP
# Single-file FastAPI app for deal intake + scoring.
#
# What it does:
# - creates deal records from email text / notes / attachments
# - stores deals in SQLite
# - scores using Saagar's 5-factor rubric
# - supports confidence + missing info + knockout flags
# - designed for Acquisitions/2026 intake, but folder watching can be added later
#
# Run:
#   pip install fastapi uvicorn python-multipart
#   uvicorn cemtrex_ma_copilot_mvp:app --reload
#
# Optional env vars:
#   MA_DB_PATH=./ma_copilot.db
#   OPENAI_API_KEY=...
#   OPENAI_MODEL=gpt-5.4-mini
#
# Notes:
# - This file uses deterministic heuristics by default.
# - If OPENAI_API_KEY is set, you can wire in an LLM extractor in llm_extract().
# - Outlook folder ingestion is intentionally not hardcoded in v1.
#   The right pattern is: Acquisitions/2026 is your trigger queue, not your only data source.
# ============================================================

app = FastAPI(title="Cemtrex M&A Copilot MVP", version="0.1.0")
DB_PATH = os.getenv("MA_DB_PATH", "./ma_copilot.db")


# -----------------------------
# Models
# -----------------------------
class ScoreBand(str):
    GO = "go_for_it"
    CONSERVATIVE = "proceed_conservatively"
    BETTER = "go_find_something_better"


class NextStep(str):
    PASS = "pass"
    REQUEST_INFO = "request_info"
    NDA = "nda"
    SELLER_CALL = "seller_call"
    DEEP_DIVE = "deep_dive"


class SourceType(str):
    EMAIL = "email"
    TEASER = "teaser"
    CIM = "cim"
    NOTES = "notes"
    ATTACHMENT = "attachment"


class CategoryScore(BaseModel):
    score: int = Field(ge=1, le=5)
    confidence: Literal["high", "medium", "low"]
    reason: str


class DealScorecard(BaseModel):
    autonomy: CategoryScore
    cash_flow_quality: CategoryScore
    growth_reality: CategoryScore
    downside_risk: CategoryScore
    strategic_fit: CategoryScore
    total_score: int = Field(ge=5, le=25)
    score_band: Literal["go_for_it", "proceed_conservatively", "go_find_something_better"]
    knockout_flags: list[str] = Field(default_factory=list)
    missing_information: list[str] = Field(default_factory=list)
    red_flags: list[str] = Field(default_factory=list)
    next_step: Literal["pass", "request_info", "nda", "seller_call", "deep_dive"]
    executive_view: str


class DealRecord(BaseModel):
    deal_id: str
    company_name: str | None = None
    broker: str | None = None
    subject: str | None = None
    source_type: Literal["email", "teaser", "cim", "notes", "attachment"]
    raw_text: str
    extracted_facts: dict[str, Any]
    scorecard: DealScorecard
    created_at: str
    updated_at: str


class IngestRequest(BaseModel):
    subject: str | None = None
    sender: str | None = None
    source_type: Literal["email", "teaser", "cim", "notes", "attachment"] = "email"
    raw_text: str


# -----------------------------
# DB
# -----------------------------
def get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS deals (
            deal_id TEXT PRIMARY KEY,
            company_name TEXT,
            broker TEXT,
            subject TEXT,
            source_type TEXT NOT NULL,
            raw_text TEXT NOT NULL,
            extracted_facts TEXT NOT NULL,
            scorecard TEXT NOT NULL,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """
    )
    conn.commit()
    conn.close()


@app.on_event("startup")
def startup_event() -> None:
    init_db()


# -----------------------------
# Utility helpers
# -----------------------------
def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


@dataclass
class Evidence:
    value: Any
    confidence: str
    present: bool


def clean_text(text: str) -> str:
    text = text.replace("\r", "\n")
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def extract_money(text: str, label: str) -> str | None:
    pattern = rf"{label}[^\n:$]*[:$\s]+([\$]?[\d,.]+\s?(?:m|mm|million|b|bn|billion)?)"
    m = re.search(pattern, text, re.IGNORECASE)
    return m.group(1).strip() if m else None


def find_first(text: str, patterns: list[str]) -> str | None:
    for pattern in patterns:
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            if m.groups():
                return m.group(1).strip()
            return m.group(0).strip()
    return None


# -----------------------------
# Extraction
# -----------------------------
def heuristic_extract(raw_text: str, subject: str | None, sender: str | None) -> dict[str, Any]:
    text = clean_text(raw_text)
    company_name = find_first(
        text,
        [
            r"company:\s*([^\n]+)",
            r"project\s+([A-Z][A-Za-z0-9\- ]+)",
            r"subject:\s*acquisition opportunity:\s*([^\n]+)",
        ],
    )

    location = find_first(text, [r"location:\s*([^\n]+)"])
    industry = find_first(
        text,
        [
            r"description:\s*the company is\s*([^\n.]+)",
            r"industry:\s*([^\n]+)",
            r"sector:\s*([^\n]+)",
        ],
    )

    revenue = extract_money(text, "revenue")
    ebitda = extract_money(text, "ebitda")
    gross_margin = find_first(text, [r"gross margin[s]?[:\s]+([\d.]+%)"])
    recurring = find_first(
        text,
        [r"recurring revenue[:\s]+([\d.]+%)", r"repeat revenue[:\s]+([\d.]+%)"],
    )

    founder_owned = bool(re.search(r"founder[- ]owned|owner is seeking retirement|seller is the business", text, re.I))
    management_team = bool(re.search(r"management team|second layer|leadership team|president|ceo hire", text, re.I))
    concentration = find_first(
        text,
        [r"customer concentration[:\s]+([^\n]+)", r"top customer[s]?[:\s]+([^\n]+)"],
    )
    addbacks = bool(re.search(r"addback|adjusted ebitda|normaliz(?:ed|ation)", text, re.I))
    backlog = find_first(text, [r"backlog[:\s]+([^\n]+)"])
    links = re.findall(r"https?://\S+", text)

    return {
        "company_name": company_name,
        "subject": subject,
        "sender": sender,
        "broker": sender,
        "location": location,
        "industry": industry,
        "revenue": revenue,
        "ebitda": ebitda,
        "gross_margin": gross_margin,
        "recurring_revenue": recurring,
        "founder_owned": founder_owned,
        "management_team_mentioned": management_team,
        "customer_concentration": concentration,
        "addbacks_mentioned": addbacks,
        "backlog": backlog,
        "links": links,
    }


# -----------------------------
# Scoring logic
# -----------------------------
def score_band(total_score: int) -> ScoreBand:
    if total_score > 18:
        return ScoreBand.GO
    if 15 <= total_score <= 17:
        return ScoreBand.CONSERVATIVE
    return ScoreBand.BETTER


def category(score: int, confidence: str, reason: str) -> CategoryScore:
    return CategoryScore(score=score, confidence=confidence, reason=reason)


def score_autonomy(facts: dict[str, Any], text: str) -> CategoryScore:
    if facts.get("management_team_mentioned") and not facts.get("founder_owned"):
        return category(5, "medium", "Real management appears to exist and the seller does not appear central to daily operations.")
    if facts.get("management_team_mentioned") and facts.get("founder_owned"):
        return category(3, "medium", "Some key-man risk exists, but there are signs of management depth.")
    if facts.get("founder_owned"):
        return category(1, "high", "Owner dependence appears high and transition risk looks material.")
    return category(3, "low", "Management depth is unclear, so autonomy is scored neutral pending org details.")


def score_cash_flow_quality(facts: dict[str, Any], text: str) -> CategoryScore:
    gm = facts.get("gross_margin")
    recurring = facts.get("recurring_revenue")
    if gm:
        try:
            gm_value = float(gm.replace("%", ""))
            if gm_value >= 40 and recurring:
                return category(5, "medium", "Margins look healthy and there is evidence of recurring or repeat revenue.")
            if gm_value >= 25:
                return category(3, "medium", "Economics look decent, but quality is not yet clearly exceptional.")
            return category(1, "medium", "Margin profile appears rough and cash flow quality may be weak.")
        except ValueError:
            pass

    if facts.get("addbacks_mentioned"):
        return category(3, "low", "Adjusted EBITDA is referenced, but quality of earnings is not yet proven.")
    return category(3, "low", "Insufficient margin detail, so cash flow quality is neutral provisional.")


def score_growth_reality(facts: dict[str, Any], text: str) -> CategoryScore:
    growth_markers = len(re.findall(r"growth|expansion|backlog|cross-sell|tailwind|new markets", text, re.I))
    if growth_markers >= 3:
        return category(5, "medium", "There is a practical path to growth supported by multiple indicators.")
    if growth_markers >= 1:
        return category(3, "medium", "Modest growth seems plausible, but proof is limited.")
    return category(1, "low", "No practical growth path is evident from the materials provided.")


def score_downside_risk(facts: dict[str, Any], text: str) -> CategoryScore:
    risk_markers = []
    if facts.get("customer_concentration"):
        risk_markers.append("customer concentration")
    if facts.get("founder_owned"):
        risk_markers.append("owner dependence")
    if re.search(r"cyclical|volatile|large contracts|project-based|timing", text, re.I):
        risk_markers.append("revenue volatility")

    if not risk_markers:
        return category(5, "low", "No major fragility markers are obvious yet, though data is still incomplete.")
    if len(risk_markers) == 1:
        return category(3, "medium", f"Risk appears manageable, but {risk_markers[0]} needs diligence.")
    return category(1, "medium", f"Multiple downside risks are present: {', '.join(risk_markers)}.")


def score_strategic_fit(facts: dict[str, Any], text: str) -> CategoryScore:
    strategic_terms = re.findall(
        r"AIS|industrial|automation|controls|electrical|fire suppression|aerospace|defense|service", text, re.I
    )
    if len(strategic_terms) >= 3:
        return category(5, "medium", "This looks like a direct platform fit or close tuck-in opportunity.")
    if len(strategic_terms) >= 1:
        return category(3, "medium", "This appears adjacent to the current operating footprint.")
    return category(1, "low", "This looks too standalone or strategically distant.")


def collect_missing_info(facts: dict[str, Any]) -> list[str]:
    missing = []
    if not facts.get("company_name"):
        missing.append("company name")
    if not facts.get("revenue"):
        missing.append("revenue")
    if not facts.get("ebitda"):
        missing.append("EBITDA")
    if not facts.get("gross_margin"):
        missing.append("gross margin")
    if not facts.get("customer_concentration"):
        missing.append("customer concentration")
    if not facts.get("management_team_mentioned"):
        missing.append("management org depth")
    if not facts.get("recurring_revenue"):
        missing.append("recurring/repeat revenue profile")
    return missing


def collect_red_flags(facts: dict[str, Any], text: str) -> list[str]:
    flags = []
    if facts.get("founder_owned"):
        flags.append("Seller dependence may be high.")
    if facts.get("addbacks_mentioned"):
        flags.append("Adjusted EBITDA / addbacks need a hard quality-of-earnings check.")
    if facts.get("customer_concentration"):
        flags.append("Customer concentration needs verification.")
    if re.search(r"data room|nda|follow this link", text, re.I):
        flags.append("Core diligence materials may sit outside the email and require manual retrieval.")
    return flags


def collect_knockout_flags(facts: dict[str, Any], text: str) -> list[str]:
    flags = []
    if facts.get("founder_owned") and not facts.get("management_team_mentioned"):
        flags.append("Autonomy risk looks extreme.")
    if re.search(r"one customer|single customer|top customer.*(?:40|50|60)%", text, re.I):
        flags.append("Customer concentration may be extreme.")
    if re.search(r"litigation|environmental claim|regulatory action|going concern", text, re.I):
        flags.append("Potential legal or solvency issue.")
    return flags


def recommend_next_step(total_score: int, missing_info: list[str], knockout_flags: list[str]) -> NextStep:
    if knockout_flags and total_score < 18:
        return NextStep.REQUEST_INFO if total_score >= 15 else NextStep.PASS
    if total_score > 18:
        return NextStep.SELLER_CALL if len(missing_info) <= 3 else NextStep.REQUEST_INFO
    if 15 <= total_score <= 17:
        return NextStep.REQUEST_INFO
    return NextStep.PASS


def build_executive_view(scorecard: DealScorecard) -> str:
    if scorecard.total_score > 18:
        return "Good candidate. Push forward, but clean up the missing data fast before you romanticize it."
    if 15 <= scorecard.total_score <= 17:
        return "Interesting enough to keep alive, but only proceed conservatively and close the obvious gaps first."
    return "Not good enough. Go find something better unless price or structure is unusually attractive."


def evaluate_deal(facts: dict[str, Any], raw_text: str) -> DealScorecard:
    text = clean_text(raw_text)
    autonomy = score_autonomy(facts, text)
    cash_flow_quality = score_cash_flow_quality(facts, text)
    growth_reality = score_growth_reality(facts, text)
    downside_risk = score_downside_risk(facts, text)
    strategic_fit = score_strategic_fit(facts, text)

    total = (
        autonomy.score
        + cash_flow_quality.score
        + growth_reality.score
        + downside_risk.score
        + strategic_fit.score
    )

    missing = collect_missing_info(facts)
    red_flags = collect_red_flags(facts, text)
    knockout_flags = collect_knockout_flags(facts, text)
    next_step = recommend_next_step(total, missing, knockout_flags)

    preliminary = DealScorecard(
        autonomy=autonomy,
        cash_flow_quality=cash_flow_quality,
        growth_reality=growth_reality,
        downside_risk=downside_risk,
        strategic_fit=strategic_fit,
        total_score=total,
        score_band=score_band(total),
        knockout_flags=knockout_flags,
        missing_information=missing,
        red_flags=red_flags,
        next_step=next_step,
        executive_view="",
    )
    preliminary.executive_view = build_executive_view(preliminary)
    return preliminary


# -----------------------------
# Persistence
# -----------------------------
def save_deal(record: DealRecord) -> None:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """
        INSERT OR REPLACE INTO deals (
            deal_id, company_name, broker, subject, source_type,
            raw_text, extracted_facts, scorecard, created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            record.deal_id,
            record.company_name,
            record.broker,
            record.subject,
            record.source_type,
            record.raw_text,
            json.dumps(record.extracted_facts),
            record.scorecard.model_dump_json(),
            record.created_at,
            record.updated_at,
        ),
    )
    conn.commit()
    conn.close()


def load_deals() -> list[DealRecord]:
    conn = get_conn()
    cur = conn.cursor()
    rows = cur.execute("SELECT * FROM deals ORDER BY updated_at DESC").fetchall()
    conn.close()
    items: list[DealRecord] = []
    for row in rows:
        items.append(
            DealRecord(
                deal_id=row["deal_id"],
                company_name=row["company_name"],
                broker=row["broker"],
                subject=row["subject"],
                source_type=row["source_type"],
                raw_text=row["raw_text"],
                extracted_facts=json.loads(row["extracted_facts"]),
                scorecard=DealScorecard.model_validate_json(row["scorecard"]),
                created_at=row["created_at"],
                updated_at=row["updated_at"],
            )
        )
    return items


# -----------------------------
# API
# -----------------------------
@app.get("/")
def root() -> dict[str, str]:
    return {"status": "ok", "service": "Cemtrex M&A Copilot MVP"}


@app.get("/deals", response_model=list[DealRecord])
def list_deals() -> list[DealRecord]:
    return load_deals()


@app.post("/ingest", response_model=DealRecord)
def ingest_text(payload: IngestRequest) -> DealRecord:
    now = utc_now_iso()
    facts = heuristic_extract(payload.raw_text, payload.subject, payload.sender)
    scorecard = evaluate_deal(facts, payload.raw_text)

    record = DealRecord(
        deal_id=str(uuid.uuid4()),
        company_name=facts.get("company_name"),
        broker=payload.sender,
        subject=payload.subject,
        source_type=payload.source_type,
        raw_text=payload.raw_text,
        extracted_facts=facts,
        scorecard=scorecard,
        created_at=now,
        updated_at=now,
    )
    save_deal(record)
    return record


@app.post("/ingest-file", response_model=DealRecord)
async def ingest_file(
    file: UploadFile = File(...),
    subject: str | None = Form(default=None),
    sender: str | None = Form(default=None),
    source_type: str = Form(default="attachment"),
) -> DealRecord:
    contents = await file.read()
    try:
        raw_text = contents.decode("utf-8", errors="ignore")
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Could not decode file: {exc}") from exc

    now = utc_now_iso()
    facts = heuristic_extract(raw_text, subject or file.filename, sender)
    scorecard = evaluate_deal(facts, raw_text)

    record = DealRecord(
        deal_id=str(uuid.uuid4()),
        company_name=facts.get("company_name"),
        broker=sender,
        subject=subject or file.filename,
        source_type=source_type,  # type: ignore[arg-type]
        raw_text=raw_text,
        extracted_facts=facts,
        scorecard=scorecard,
        created_at=now,
        updated_at=now,
    )
    save_deal(record)
    return record


@app.get("/deals/{deal_id}", response_model=DealRecord)
def get_deal(deal_id: str) -> DealRecord:
    for deal in load_deals():
        if deal.deal_id == deal_id:
            return deal
    raise HTTPException(status_code=404, detail="Deal not found")


@app.get("/dashboard/summary")
def dashboard_summary() -> dict[str, Any]:
    deals = load_deals()
    total = len(deals)
    by_band = {"go_for_it": 0, "proceed_conservatively": 0, "go_find_something_better": 0}
    next_steps: dict[str, int] = {}

    for deal in deals:
        by_band[deal.scorecard.score_band] += 1
        next_steps[deal.scorecard.next_step] = next_steps.get(deal.scorecard.next_step, 0) + 1

    return {
        "total_deals": total,
        "score_bands": by_band,
        "next_steps": next_steps,
    }


# -----------------------------
# Outlook integration stub
# -----------------------------
# Add this next:
# 1. Poll Acquisitions/2026 via Microsoft Graph or Outlook API.
# 2. Create a deal record for each new email.
# 3. Pull attachments + links.
# 4. If the link points to a data room / teaser / NDA, flag human action.
# 5. Re-run evaluation when new artifacts arrive.
#
# This is intentionally left as a stub because your actual environment,
# authentication method, and attachment flow will determine the right implementation.


# -----------------------------
# Sample test payload
# -----------------------------
SAMPLE_EMAIL = """
Subject: Acquisition Opportunity: Fire Suppression Services
Sender: kari@plethorabusinesses.com

Greetings,

I am reaching out to share Project Phoenix, a provider of fire protection systems and underground utility construction.
The Company provides design, installation, inspection, and service for fire suppression systems.
If this is of interest, follow this link to the Teaser and NDA.
The owner is open to a variety of deal structures and is willing to remain with the Company for a period of time.
"""

if __name__ == "__main__":
    facts = heuristic_extract(SAMPLE_EMAIL, "Acquisition: Fire Suppression Services", "kari@plethorabusinesses.com")
    scorecard = evaluate_deal(facts, SAMPLE_EMAIL)
    print(json.dumps({"facts": facts, "scorecard": scorecard.model_dump()}, indent=2))
