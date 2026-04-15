"""
Main deal pipeline: CIM → Score → Excel → Diligence → LOI → Dropbox → Slack.
"""

from __future__ import annotations

import io
import json
from typing import Any, Callable

from openai import OpenAI

from prompts import (
    SCORE_CIM_PROMPT,
    EXTRACT_FINANCIALS_PROMPT,
    DILIGENCE_QUESTIONS_PROMPT,
)
from excel_model import build_financial_model_from_dict
from doc_generators import build_score_memo, build_diligence_doc
from loi_builder import LOITerms, build_loi, LOI_VARIABLES_PROMPT
import dropbox_client


def _clean_json_response(text: str) -> str:
    """Strip markdown fences from an LLM JSON response."""
    text = text.strip()
    if text.startswith("```"):
        # Remove opening fence (```json or ```)
        first_newline = text.index("\n") if "\n" in text else 3
        text = text[first_newline + 1:]
    if text.endswith("```"):
        text = text[:-3]
    return text.strip()


def _call_openai_with_pdf(client: OpenAI, file_id: str, prompt: str) -> str:
    """Call OpenAI responses API with a PDF file reference."""
    response = client.responses.create(
        model="gpt-4.1-mini",
        input=[{
            "role": "user",
            "content": [
                {"type": "input_file", "file_id": file_id},
                {"type": "input_text", "text": prompt},
            ],
        }],
    )
    return response.output_text


def _call_openai_text(client: OpenAI, prompt: str) -> str:
    """Call OpenAI chat completions with a text prompt."""
    response = client.chat.completions.create(
        model="gpt-4.1-mini",
        messages=[{"role": "user", "content": prompt}],
        temperature=0,
    )
    return response.choices[0].message.content


# ── Slack notification formatters ─────────────────────────────────────────

def _format_slack_go(company: str, total: int, financials: dict,
                     valuation: dict, folder_link: str) -> str:
    rev = financials.get("revenue", [None])[-1]
    ebitda_vals = financials.get("ebitda", [None])
    ebitda = ebitda_vals[-1] if ebitda_vals else None
    rev_str = f"${rev / 1_000_000:.1f}M" if rev else "N/A"
    ebitda_str = f"${ebitda / 1_000_000:.1f}M" if ebitda else "N/A"

    gp = financials.get("gross_profit", [None])[-1]
    margin_str = f"{gp / rev * 100:.0f}%" if gp and rev and rev > 0 else "N/A"

    val_low = valuation.get("low", "?")
    val_mid = valuation.get("mid", "?")
    val_high = valuation.get("high", "?")

    return (
        f"🏭 *NEW DEAL: {company}*\n"
        f"🟢 Score: *{total}/25* — *GO*\n"
        f"💰 Revenue: {rev_str} | EBITDA: {ebitda_str} | Margin: {margin_str}\n"
        f"🎯 Valuation Range: {val_low} - {val_high} (mid: {val_mid})\n"
        f"📁 Dropbox: {folder_link}\n"
        f"✅ Full package ready: Score Memo | Financial Model | Diligence Qs | LOI Draft"
    )


def _format_slack_cautious(company: str, total: int, financials: dict,
                            valuation: dict, folder_link: str) -> str:
    rev = financials.get("revenue", [None])[-1]
    ebitda_vals = financials.get("ebitda", [None])
    ebitda = ebitda_vals[-1] if ebitda_vals else None
    rev_str = f"${rev / 1_000_000:.1f}M" if rev else "N/A"
    ebitda_str = f"${ebitda / 1_000_000:.1f}M" if ebitda else "N/A"

    gp = financials.get("gross_profit", [None])[-1]
    margin_str = f"{gp / rev * 100:.0f}%" if gp and rev and rev > 0 else "N/A"

    val_low = valuation.get("low", "?")
    val_mid = valuation.get("mid", "?")
    val_high = valuation.get("high", "?")

    return (
        f"🏭 *NEW DEAL: {company}*\n"
        f"🟡 Score: *{total}/25* — *PROCEED CAUTIOUSLY*\n"
        f"💰 Revenue: {rev_str} | EBITDA: {ebitda_str} | Margin: {margin_str}\n"
        f"🎯 Valuation Range: {val_low} - {val_high} (mid: {val_mid})\n"
        f"📁 Dropbox: {folder_link}\n"
        f"✅ Full package ready: Score Memo | Financial Model | Diligence Qs | LOI Draft"
    )


def _format_slack_pass(company: str, total: int, summary: str) -> str:
    return (
        f"🏭 *NEW DEAL: {company}*\n"
        f"🔴 Score: *{total}/25* — *PASS. Go find something better.*\n"
        f"📋 {summary}"
    )


# ── Main pipeline ─────────────────────────────────────────────────────────

def run_deal_pipeline(
    openai_client: OpenAI,
    pdf_bytes: bytes,
    pdf_name: str,
    slack_say: Callable[[str], None] | None = None,
    slack_channel: str | None = None,
) -> dict[str, Any]:
    """Run the full deal pipeline on a CIM PDF.

    Returns a dict with all results:
        {
            "company_name": str,
            "scorecard": dict,
            "financials": dict,
            "total_score": int,
            "verdict": str,
            "files_uploaded": dict,
            "folder_link": str,
        }
    """
    result: dict[str, Any] = {}

    # Upload PDF to OpenAI for processing
    uploaded_file = openai_client.files.create(
        file=(pdf_name, io.BytesIO(pdf_bytes), "application/pdf"),
        purpose="user_data",
    )
    file_id = uploaded_file.id

    # ── Step 1: Score the CIM ─────────────────────────────────────────
    score_text = _call_openai_with_pdf(openai_client, file_id, SCORE_CIM_PROMPT)
    scorecard = json.loads(_clean_json_response(score_text))
    result["scorecard"] = scorecard

    company = scorecard.get("company_name", "Unknown Company")
    total = scorecard.get("total_score", 0)
    verdict = scorecard.get("verdict", "PASS")
    result["company_name"] = company
    result["total_score"] = total
    result["verdict"] = verdict

    # ── Dedup check against existing Dropbox folders ──────────────────
    dup = dropbox_client.check_duplicate(company, openai_client)
    if dup:
        result["duplicate"] = dup
        match_name = dup.get("match", "")
        confidence = dup.get("confidence", "")
        reason = dup.get("reason", "")
        if slack_say:
            slack_say(
                f"⚠️ *Possible duplicate for {company}*\n"
                f"Matched existing folder: *{match_name}* ({confidence} confidence)\n"
                f"Reason: {reason}\n"
                f"Proceeding with pipeline — files will go to the existing *{match_name}* folder."
            )
        # Use the existing folder name so files land in the same place
        company = match_name
        result["company_name"] = company

    # ── Step 1b: Extract financials ───────────────────────────────────
    fin_text = _call_openai_with_pdf(openai_client, file_id, EXTRACT_FINANCIALS_PROMPT)
    financials = json.loads(_clean_json_response(fin_text))
    result["financials"] = financials

    # ── Score < 15: PASS — Slack only, no files ───────────────────────
    if total < 15:
        summary = scorecard.get("executive_summary", "Low score across multiple dimensions.")
        if slack_say:
            slack_say(_format_slack_pass(company, total, summary))
        result["files_uploaded"] = {}
        result["folder_link"] = ""
        return result

    # ── Score >= 15: Full package ─────────────────────────────────────
    files_to_upload: dict[str, bytes] = {}

    # Score memo
    memo_bytes = build_score_memo(scorecard)
    files_to_upload[f"{company}_Score_Memo.docx"] = memo_bytes

    # Excel financial model
    xlsx_bytes = build_financial_model_from_dict(financials)
    files_to_upload[f"{company}_Financial_Model.xlsx"] = xlsx_bytes

    # ── Step 3: Diligence questions ───────────────────────────────────
    dq_prompt = DILIGENCE_QUESTIONS_PROMPT.replace(
        "__SCORECARD_JSON__", json.dumps(scorecard, indent=2)
    )
    dq_text = _call_openai_with_pdf(openai_client, file_id, dq_prompt)
    questions = json.loads(_clean_json_response(dq_text))
    result["diligence_questions"] = questions

    dq_bytes = build_diligence_doc(company, questions)
    files_to_upload[f"{company}_Diligence_Questions.docx"] = dq_bytes

    # ── Step 4: LOI draft ─────────────────────────────────────────────
    loi_var_prompt = LOI_VARIABLES_PROMPT.replace(
        "__SCORECARD_JSON__", json.dumps(scorecard, indent=2)
    ).replace(
        "__FINANCIAL_JSON__", json.dumps(financials, indent=2)
    )
    loi_vars_text = _call_openai_with_pdf(openai_client, file_id, loi_var_prompt)
    loi_vars = json.loads(_clean_json_response(loi_vars_text))
    result["loi_vars"] = loi_vars

    terms = LOITerms(
        date=loi_vars.get("date", ""),
        seller_names=loi_vars.get("seller_names", ""),
        seller_greeting=loi_vars.get("seller_greeting", ""),
        company_name=loi_vars.get("company_name", company),
        company_abbreviation=loi_vars.get("company_abbreviation", ""),
        company_address_lines=loi_vars.get("company_address_lines", []),
        buyer_name=loi_vars.get("buyer_name", "Cemtrex"),
        opening_paragraphs=loi_vars.get("opening_paragraphs", []),
        total_consideration=loi_vars.get("total_consideration", ""),
        cash_at_close=loi_vars.get("cash_at_close", ""),
        note_amount=loi_vars.get("note_amount", ""),
        has_earnout=loi_vars.get("has_earnout", False),
        earnout_text=loi_vars.get("earnout_text", ""),
        nwc_floor=loi_vars.get("nwc_floor", "[TO BE DETERMINED DURING DILIGENCE]"),
        property_text=loi_vars.get("property_text", ""),
        employees_text=loi_vars.get("employees_text", ""),
        exclusivity_days=loi_vars.get("exclusivity_days", "sixty (60)"),
        closing_paragraphs=loi_vars.get("closing_paragraphs", []),
    )
    loi_bytes = build_loi(terms)
    files_to_upload[f"{company}_LOI_Draft.docx"] = loi_bytes

    # ── Upload to Dropbox ─────────────────────────────────────────────
    paths = dropbox_client.upload_deal_outputs(company, files_to_upload)
    result["files_uploaded"] = paths

    folder_link = dropbox_client.get_deal_folder_link(company)
    result["folder_link"] = folder_link

    # ── Slack notification ────────────────────────────────────────────
    if slack_say:
        valuation = scorecard.get("valuation_range", {})
        if total >= 18:
            msg = _format_slack_go(company, total, financials, valuation, folder_link)
        else:
            msg = _format_slack_cautious(company, total, financials, valuation, folder_link)
        slack_say(msg)

    # Store the OpenAI file_id for follow-up Q&A
    result["openai_file_id"] = file_id

    return result
