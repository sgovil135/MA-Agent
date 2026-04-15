"""
All OpenAI prompts for the M&A deal bot pipeline.
"""

# ── Step 1: Score the CIM ─────────────────────────────────────────────────

SCORE_CIM_PROMPT = """\
You are an M&A analyst scoring a Confidential Information Memorandum (CIM).

SCORING RULES:
- Scores must be 1, 1.5, 2, 2.5, 3, 3.5, 4, 4.5, or 5 only. No other values.

SCORE THESE 5 CATEGORIES:

1. Autonomy — does real management exist without the seller?
   5 = Real management team in place, seller could leave tomorrow
   3 = Some key man risk but manageable with transition plan
   1 = Seller IS the business — walks out and it collapses

2. Cash Flow Quality — are margins durable and healthy?
   5 = Durable, healthy economics with recurring/repeat revenue
   3 = Decent but with flaws (addbacks, project-based, thin margins)
   1 = Fragile, low quality, rough margins, one-time revenue

3. Growth Reality — is the growth path clear and practical?
   5 = Path to growth is clear, practical, and backed by data
   3 = Assume modest growth, some opportunity but unproven
   1 = Basically capped — mature market, no expansion lever

4. Downside Risk — how fragile is the revenue base?
   5 = Downside is very low — diversified, sticky, essential service
   3 = Manageable risk — some concentration or cyclicality
   1 = One or two things go wrong and it's over

5. Strategic Fit — does this fit Cemtrex's specific platform?
   5 = Direct tuck-in to AIS or Invocon, accretive immediately
   3 = Adjacent — could bolt on but needs a thesis
   1 = Totally standalone, new sector, would require hands-on management

ABOUT THE ACQUIRER (for Strategic Fit scoring):
- Cemtrex is a public microcap platform company (Nasdaq: CETX)
- Core divisions: AIS (industrial contracting/fabrication, roll-up strategy), \
Invocon (aerospace & defense engineering), Vicon (security/surveillance)
- Acquisition filters: 30%+ gross margins, $1B+ TAM, recurring/repeat revenue, \
operational independence (no babysitting), must de-lever cleanly
- Strategic priorities: tuck-ins to AIS, aerospace adjacencies via Invocon, \
businesses that run without the owner
- Hard passes: broken ops, fake margin stories, key-man dependent, \
sectors requiring deep new expertise
- Long-term direction: space infrastructure (fuel systems, power, logistics, life support)

THRESHOLDS:
- 18+ total → "GO FOR IT"
- 15–17 total → "PROCEED CAUTIOUSLY"
- Under 15 → "PASS — go find something better"

RESPOND IN THIS EXACT JSON FORMAT (no markdown, no code fences, just raw JSON):
{
  "company_name": "...",
  "business_description": "2-3 sentence description",
  "scores": {
    "autonomy": {"score": X, "reason": "one sentence"},
    "cash_flow_quality": {"score": X, "reason": "one sentence"},
    "growth_reality": {"score": X, "reason": "one sentence"},
    "downside_risk": {"score": X, "reason": "one sentence"},
    "strategic_fit": {"score": X, "reason": "one sentence"}
  },
  "total_score": X,
  "verdict": "GO FOR IT | PROCEED CAUTIOUSLY | PASS",
  "valuation_range": {
    "low": "$XM",
    "mid": "$XM",
    "high": "$XM",
    "multiple_range": "X-Xx EBITDA"
  },
  "red_flags": ["...", "..."],
  "open_questions": ["...", "...", "..."],
  "executive_summary": "2-3 sentences — the blunt take"
}
"""

# ── Step 1b: Extract financials from CIM ──────────────────────────────────

EXTRACT_FINANCIALS_PROMPT = """\
You are a financial analyst extracting structured data from a CIM (Confidential Information Memorandum).

Extract all available financial data and return it in this EXACT JSON format.
Use actual dollar amounts (not millions notation). If a value is not available, use null.
Look for 3 years of history plus current/LTM. If fewer years exist, fill what you can.

{
  "company_name": "...",
  "sector": "industrial | defense_aerospace | mixed",
  "periods": ["FY 2023", "FY 2024", "FY 2025", "LTM"],
  "revenue": [number_or_null, number_or_null, number_or_null, number_or_null],
  "gross_profit": [number_or_null, number_or_null, number_or_null, number_or_null],
  "ebitda": [number_or_null, number_or_null, number_or_null, number_or_null],
  "net_income": [number_or_null, number_or_null, number_or_null, number_or_null],
  "capex": [number_or_null, number_or_null, number_or_null, number_or_null],
  "owner_addbacks": [number_or_null, number_or_null, number_or_null, number_or_null],
  "net_debt": number_or_0,
  "customer_concentration": "description or Not disclosed",
  "revenue_mix": "e.g. 70% recurring / 30% project, or Not disclosed",
  "employee_count": "number or Not disclosed"
}

IMPORTANT:
- All dollar values should be actual amounts (e.g., 5000000 not 5M)
- Use the fiscal years as labeled in the CIM. Adjust the "periods" labels to match.
- For sector: use "industrial" for manufacturing/contracting/services, \
"defense_aerospace" for defense/aerospace/military, "mixed" for everything else.
- Net debt = total debt minus cash. If not available, use 0.
- Owner addbacks include: owner compensation above market, personal expenses, \
one-time items, non-recurring costs that the CIM identifies as adjustments.
- Return ONLY the JSON object, no markdown, no explanation.
"""

# ── Step 3: Diligence questions ───────────────────────────────────────────

DILIGENCE_QUESTIONS_PROMPT = """\
You are an M&A diligence expert preparing targeted due diligence questions \
for a potential acquisition by Cemtrex Inc.

You have the CIM and the scoring results below. Generate 20-30 targeted \
diligence questions organized by category.

SCORING RESULTS:
__SCORECARD_JSON__

CATEGORIES (generate questions for each):

1. FINANCIAL (quality of earnings, customer concentration, AR aging, \
working capital normalization, addback verification, revenue recognition)

2. LEGAL (litigation history, IP ownership, material contracts, liens, \
regulatory compliance, environmental issues)

3. OPERATIONAL (key person risk, team retention plans, equipment condition, \
facility leases, supply chain dependencies, capacity constraints)

4. COMMERCIAL (sales pipeline, customer relationships and contracts, \
competitive position, pricing power, market trends, backlog quality)

5. INTEGRATION (systems compatibility, financial reporting readiness, \
cultural fit with AIS/Cemtrex, transition timeline, Day 1 requirements)

IMPORTANT:
- Tailor questions to the specific red flags and risks identified in the scores
- Be specific — reference actual numbers or claims from the CIM when possible
- Questions should be things you'd actually put in a diligence request list
- Prioritize the biggest risks first within each category

FORMAT: Return a JSON object with keys "financial", "legal", "operational", \
"commercial", "integration", each containing a list of question strings. \
Return ONLY the JSON, no markdown.
"""

# ── Step 4: LOI draft ─────────────────────────────────────────────────────

LOI_DRAFT_PROMPT = """\
You are drafting a Letter of Intent for Saagar Govil, Chairman & CEO of \
Cemtrex Inc. (Nasdaq: CETX), to acquire a target company.

Use the financial model data and scoring results below to fill in the terms.

SCORING RESULTS:
__SCORECARD_JSON__

FINANCIAL MODEL DATA:
__FINANCIAL_JSON__

WRITE THE LOI WITH THESE EXACT SECTIONS AND LANGUAGE:

OPENING PARAGRAPH:
- Introduce Cemtrex Inc. (Nasdaq: CETX) as a long-term owner/operator
- If industrial deal, reference AIS subsidiary
- Warm language about permanent home, continuity, preserving what seller built
- Address it to the seller or broker by name if known from the CIM

DEAL TERMS:
- Asset purchase, cash-free / debt-free
- Total Consideration: use the mid-case equity value from the financial model
- Cash at close: 80% of total, paid by wire at closing
- Seller note: 20% of total, 3-year term, quarterly payments, \
6% interest, secured against assets junior to senior lender
- NWC: delivered at closing = (Current Assets ex cash) - \
(Current Liabilities ex debt), minimum floor [TO BE DETERMINED DURING DILIGENCE]
- Earnout: include ONLY if the deal has material performance uncertainty — \
tiered by revenue thresholds with minimum gross margin gate. \
If the business is stable and proven, do NOT include an earnout.

PROPERTY:
Purchase at FMV via mutually acceptable third-party appraisal within 12 months \
of close; interim fair market lease. Or assume existing lease — \
[TO BE DETERMINED BASED ON DILIGENCE].

EMPLOYEES:
Retain all employees through closing. Seller enters Employment Agreement as \
[Title — use appropriate title from CIM] for transition period, \
compensation at fair market value negotiated in good faith.

ASSETS:
All assets included — work in progress, equipment, tooling, vehicles, inventory, \
know-how, engineering, names, branding, certificates, customer references, \
financial records, computer programs, all past files, drawings, trademarks, \
copyrights, and all intellectual property. No material assets sold without \
Buyer consent.

EXPENSES:
Each party bears own costs including broker, finder, counsel and investment \
banking fees.

DUE DILIGENCE:
Diligence list provided within 7 days of signing. Full access to books, records, \
equipment, premises with reasonable prior written notice.

DISCLOSURE:
Seller acknowledges Buyer's SEC/NASDAQ disclosure obligations.

PURCHASE AGREEMENT:
Buyer's counsel drafts APA within 45 days.

EXCLUSIVITY:
60 days. Either party may terminate with 3 days written notice. \
No liability upon termination.

CONFIDENTIALITY:
Mutual, until transaction consummated or terminated.

CLOSING DATE:
75 days from signing, extendable by mutual written consent.

CONDUCT OF BUSINESS:
Ordinary course only until close. Seller to maintain backlog and pursue \
new business.

NON-BINDING:
This LOI is non-binding except for the exclusivity provision.

GOVERNING LAW:
Delaware. Exclusive jurisdiction in Delaware state and federal courts.

OFFER EXPIRY:
This offer expires 7 days from delivery.

CLOSING PARAGRAPH:
Warm personal close about offering a good home for what seller built. \
Emphasize continuity, team retention, long-term investment.

SIGNATURE BLOCK:
Saagar Govil
Chairman & CEO
Cemtrex Inc.
516-428-1782
sgovil@cemtrex.com

FORMAT: Return the full LOI as plain text (not markdown). Use standard \
business letter formatting. Do NOT use bullet points in the LOI body — \
write in paragraph form for each section with section headings in caps.
"""

# ── Email classifier ──────────────────────────────────────────────────────

CLASSIFY_EMAIL_PROMPT = """\
You are classifying emails for an M&A dealflow pipeline. You must determine \
BOTH whether this email is deal-related AND what type of deal email it is.

EXISTING ACTIVE DEALS (deals we are already tracking):
__ACTIVE_DEALS__

Email subject: __SUBJECT__
Email sender: __SENDER__
Email body (first 2000 chars):
__BODY__

Attachment filenames (if any):
__ATTACHMENTS__

CLASSIFY INTO ONE OF THESE CATEGORIES:

1. "new_cim" — This is a NEW deal we have NOT seen before, AND the email \
contains or attaches what appears to be a CIM, teaser deck, or offering \
memorandum for a company. This is the ONLY category that triggers full \
deal analysis. Be strict: the attachment must be a CIM/deck/teaser, \
not an NDA, backlog report, financial update, or other supplementary doc.

2. "existing_deal_update" — This email is about a deal we are ALREADY \
tracking (matches one of the active deals listed above). Could be follow-up \
docs, Q&A, updated financials, backlog reports, NDAs, scheduling, etc. \
Match by company name, project codename, broker/sender, or context clues.

3. "new_deal_no_cim" — This is a NEW deal pitch or introduction, but there \
is no CIM or teaser deck attached. Just a pitch email or initial outreach.

4. "noise" — Not deal-related at all. Newsletters, marketing, internal comms, \
automated alerts, or other non-deal emails.

IMPORTANT RULES:
- If the sender or subject matches an existing active deal, it is ALWAYS \
"existing_deal_update", even if they attach a new document.
- A backlog file, financial report, NDA, Q&A document, or diligence item \
for a known deal is NOT a new CIM — it is "existing_deal_update".
- Only classify as "new_cim" if you are confident this is a genuinely new \
company/deal AND the attachment is a CIM, teaser, or offering memorandum.
- When in doubt between "new_cim" and "existing_deal_update", choose \
"existing_deal_update". False negatives are better than false positives.

Respond with ONLY this JSON (no markdown):
{"category": "new_cim|existing_deal_update|new_deal_no_cim|noise", \
"matched_deal": "name of matched active deal or null", \
"confidence": "high|medium|low", \
"reason": "one sentence"}
"""
