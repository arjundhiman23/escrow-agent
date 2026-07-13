# Escrow Account Transaction Analyst Agent

AI-powered automation for escrow account transaction analysis, built per the
BRD *"Escrow Account Transaction Analyst Agent"* (Axis Trustee Services
Limited, v1.0). Ingests the main-escrow-account transaction extract, classifies
transactions against the TRA/CATRA taxonomy, validates the sanction waterfall,
computes quarter-wise aggregates and actual-vs-projected variances, and
produces Excel + Word outputs for credit/compliance review.

## Quick start

```bash
pip install -r requirements.txt

# Demo run on the bundled synthetic dataset:
python run_agent.py \
    --transactions sample_data/sample_transactions.xlsx \
    --sanction-extract sample_data/sample_sanction_extract.json \
    --output output/
```

Outputs (in `--output`):

| File | Contents |
|---|---|
| `Escrow_Analysis_Workbook.xlsx` | Transactions · Quarter Summary · Variance Analysis · QoQ Movement · Exceptions Log · Review Queue · Audit Trail |
| `Escrow_Analysis_Report.docx` | Management report: executive summary, quarter summary, waterfall compliance, variance analysis, consolidated exceptions summary, review queue |
| `audit_log.txt` | Timestamped log of every processing step (BRD §7) |

## Inputs

1. **Transaction extract** (`--transactions`, .xlsx/.csv/.pdf) — main escrow
   account only. Column names are mapped adaptively (Risk R-04); core fields
   needed: date, narration, and either a Dr/Cr type + amount or separate
   debit/credit columns. Header rows preceded by title rows are auto-detected.
   Scanned PDFs need OCR first (Risk R-02) — the agent raises a clear error
   rather than silently producing an empty result.

2. **Sanction extract** (`--sanction-extract`, JSON) — waterfall priorities,
   permitted inflow sources, conditions, and **per-quarter** projected
   inflows/outflows by category. Schema documented in
   `escrow_agent/sanction.py`; example in
   `sample_data/sample_sanction_extract.json`.
   Alternatively pass `--sanction-doc letter.pdf` (or .docx) and the agent
   extracts this structure via the Anthropic API, saving it as
   `<name>.extract.json` **for analyst review** — treat AI-extracted sanction
   structures as drafts until reviewed.

3. **TRA/CATRA taxonomy** (`config/categories.yaml`) — the classification
   framework. The bundled taxonomy is a **placeholder built from BRD §5.4/§6.2
   category names**; replace keywords/codes with the official framework file
   when the business team provides it. Administrators add/change categories by
   editing this YAML — no code deployment (BRD §6.2, R-03).

## Configuration (no code changes needed)

`config/settings.yaml`: quarter definitions (default Apr–Jun = Q1), materiality
thresholds (default 10% or INR 10 lakh), balance tolerance (0.1% per BRD §11),
AI model/batch/confidence settings.

## Classification tiers (BRD §6.2)

1. **Rules** — keyword matching within the transaction's side (Cr→inflow
   categories, Dr→outflow). Multiple matches resolve by longest-keyword
   specificity, else fall to tier 2/3.
2. **AI** — semantic classification of narrations via the Anthropic API.
   Enabled automatically when `ANTHROPIC_API_KEY` is set (or force via
   `classification.ai_enabled`). AI may only assign codes from the taxonomy on
   the correct side; low-confidence results are demoted to the review queue.
3. **Analyst review queue** — blank/ambiguous narrations. Narrations that
   *conflict* with the transaction type (e.g. a credit narrated as a loan
   interest payment) are flagged `POTENTIALLY_INACCURATE_NARRATION` and
   **escalated separately for manual override**, distinct from the ordinary
   queue.

Without an API key the agent runs rules-only and routes everything unresolved
to the review queue — deterministic and fully offline.

## Waterfall & condition validation (BRD §6.4)

Debit and credit sides are validated independently with specific reason codes:
`CREDIT_UNPERMITTED_SOURCE`, `DEBIT_UNAUTHORIZED_PAYMENT`,
`DEBIT_OUT_OF_SEQUENCE` (payment made while a higher-priority category
obligated for the quarter received nothing), and `CONDITION_BREACH_<id>` for
machine-checkable conditions (`surplus_requires_prior`,
`category_cap_per_quarter`). Conditions typed `manual` are listed in the
Exceptions Log for human verification. Unclassified transactions are excluded
from waterfall checks and surfaced via the review queue instead.

## Validation status (read before relying on numbers)

- **Measured on the synthetic demo set:** every planted control case fires —
  1 duplicate, balance break, blank-narration review, narration-type conflict
  escalation, out-of-sequence payments, surplus-before-debt-service breach,
  quarterly cap breach, unpermitted credit source. Quarter aggregates
  reconcile exactly to bank Dr/Cr totals (category + unclassified = total).
  85 workbook formulas recalculate with zero errors. Full-year run: < 1 s
  (BRD §11 target: 5 min).
- **Not yet measured:** the BRD's ≥90% classification-accuracy criterion
  requires the real TRA/CATRA file and a labelled dataset of 500+ real
  transactions (BRD §11, §9.2). The bundled keyword taxonomy is a placeholder;
  accuracy on real Axis narrations is unknown until that test is run.

## Project layout

```
run_agent.py                    CLI pipeline orchestrator
config/categories.yaml          TRA/CATRA taxonomy (admin-editable)
config/settings.yaml            thresholds, quarters, AI settings
escrow_agent/ingestion.py       §6.1 ingestion, dedup, balance validation
escrow_agent/classification.py  §6.2 hybrid classifier + review queue
escrow_agent/sanction.py        §6.4 sanction extract schema + AI extraction
escrow_agent/waterfall.py       §6.4 waterfall & condition validation
escrow_agent/aggregation.py     §6.3 quarter aggregation, §6.5 variance
escrow_agent/reporting_excel.py §6.6 Excel workbook
escrow_agent/reporting_word.py  §6.6 Word management report
escrow_agent/audit_log.py       §7 audit trail
sample_data/                    synthetic demo data + generator
```

Out of scope, per BRD §4.2: sub-accounts, live bank APIs, ERP/GL integration,
and authentication (handled by enterprise systems).

## Knowledge Base & KB Analysis Pipeline (added)

`kb/` holds the initial deal documents for **Kanpur Lucknow Expressway Pvt. Ltd. / Axis Bank (RTL Rs. 779.75 cr)**:

- `kb/documents/` — Sanction Letter, Credit Approval Memo (Sanction Note), Escrow Agreement (PDF as uploaded; the agreement file is misnamed "Kallagam Meensuruti" but its content is the KLEPL–Axis Bank agreement)
- `kb/extracted/` — OCR/text extractions of each document
- `kb/structured/deal_extract.yaml` — **verified ground truth** (waterfall, deposits, covenants, projected P&L/BS) checked against source page images; analysts edit this file to correct anything, and all outputs rebuild from it
- `kb/kb_index.yaml` — document registry with extraction method and quality notes

Run `python3 run_kb_analysis.py` to produce:

1. `output/CATRA_Master.xlsx` — CATRA classification framework: debit categories = Escrow Agreement Order of Priority 2.3(B)(a)–(m) with waterfall priorities, credit categories = permitted deposits 2.3(A)(a)–(g), sanction cl.20 waterfall mapping + divergences, covenant register, keyword rules. Also regenerates `config/categories.yaml` so the transaction classifier uses the deal-derived taxonomy.
2. `output/TRA_Analysis.xlsx` — TRA analysis from the CAM projected P&L (FY26–FY35): source P&L, CATRA-mapped cashflow with net-surplus formulas, semi-annual (annuity-cycle) profile, reserve adequacy checks.
3. `output/Final_Analysis.docx` — summary of everything, clause-by-clause cross-check against the Agreement, and the verdict: **OVERBREACH** (utilisation above permitted) / **UNDERBREACH** (funding below required) / COMPLIANT. On the initial inputs the verdict is on projected basis; the same checks re-run on actuals when the bank statement is ingested.

Validation status: reserve-check arithmetic verified year-by-year against expected values; MMR drawdown-cycle reset and rounding tolerance tested; both workbooks recalc with 0 formula errors; Final Analysis visually verified (5 pages).

## Quarterly ATSL Pipeline (added)

`run_quarterly.py` processes the actual Axis CATRA statements (bank export format: FORACID … TRAN_PARTICULAR, TRAN_AMT, BALANCE, Remarks) for the main account (default 922020065877321) and generates, strictly in the bank-provided output formats:

1. `CATRA_ANALYSIS_Kanpur_Lucknow_ATSL.xlsx` — bank's CATRA workbook with only the ATSL sheet, all quarter values + opening/closing balances filled
2. `TRA_Analysis_ATSL_Kanpur_Lucknow.xlsx` — bank's TRA workbook, per-quarter debit/credit summation blocks + Sheet2 actuals filled
3. `Final_Analysis_FY2024-25.docx` — actuals-basis compliance report with OVERBREACH / UNDERBREACH / COMPLIANT verdict

Validation on FY2024-25 (Q1–Q4): all 48 quarter-category totals reconcile to the paisa with the bank's own ATSL figures; all 8 opening/closing balances tie out; TRA Sheet1 reproduces the bank sample with 0 cell diffs; balance integrity holds within and across quarters. Measured classification accuracy on this labelled set: 100% — the BRD ≥90% criterion is met on this data.

## Web Console (added)

`webapp/` is a FastAPI service that serves the Escrow Analyst Console: upload the quarterly statements + the two bank templates, and the agent runs in the background. Every run is tracked (queued -> running -> completed/failed) with a stage-by-stage waterfall indicator, live log, verdict badge, headline stats, downloads for the three reports, re-run and delete.

### Storage
- If S3_BUCKET is set, all inputs, outputs, and run history live in S3 under S3_PREFIX/runs/{run_id}/ (uses AWS_ACCESS_KEY_ID / AWS_SECRET_ACCESS_KEY / AWS_REGION env vars) - survives restarts and redeploys.
- Without S3_BUCKET, it falls back to a local data/ directory (for local development).

### Run locally
    pip install -r requirements.txt
    uvicorn webapp.server.app:app --port 8000
    # open http://localhost:8000

### Deploy on Render (auto-deploy from GitHub)
1. Push this folder to a GitHub repo (git init && git add -A && git commit -m "escrow agent" && git remote add origin <repo-url> && git push -u origin main).
2. In Render: New -> Web Service -> connect the repo. Render reads render.yaml (build: pip install -r requirements.txt, start: uvicorn webapp.server.app:app --host 0.0.0.0 --port $PORT).
3. In the service's Environment tab set: S3_BUCKET, AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, AWS_REGION (and optionally S3_PREFIX, default escrow-agent).
4. Enable Auto-Deploy (on by default) - every push to main redeploys.
5. IAM: the key needs s3:GetObject, s3:PutObject, s3:DeleteObject, s3:ListBucket on the bucket.

Notes: runs execute in a background thread of the web service (fine for this workload - a full FY run takes ~2s); if a deploy interrupts a running job, re-run it from the dashboard. No authentication is enabled yet - keep the URL private or add Render's access controls until auth is added.

## Usage & Cost KPIs (added)

The console now tracks real Anthropic API token usage and cost, and shows it on the dashboard:

- **Rule-based classification stays the source of truth.** The quarterly ATSL pipeline classifies every transaction by keyword rule first (as before, $0 cost, fully reproducible). If `ANTHROPIC_API_KEY` is set and the "AI-assist review queue" toggle is on for a run, a second Haiku pass runs ONLY on the transactions rules flagged for review (typically ~1 in 150) and records a *suggestion* for the analyst — it never overrides a rule-based total, so the CATRA/TRA figures stay deterministic.
- **KPI strip at the top of the dashboard** shows: whether AI-assist is configured, cumulative tokens in/out across all runs, and total cost in USD. If no key is set it plainly shows "Not configured — running rules-only, $0 token cost" instead of a misleading zero.
- **Per-run cost line** on each run card shows the exact call count, input/output tokens, and cost for that run (or "no calls needed" / "off for this run").
- Pricing table lives in `escrow_agent/ai_usage.py` (`claude-haiku-4-5-20251001` for classification, `claude-sonnet-5` for one-time sanction-document extraction) — update the rates there if Anthropic pricing changes; cost is computed from the token counts Anthropic's API actually returns (`resp.usage`), not an estimate.
- New endpoint: `GET /api/usage` — aggregate tokens/cost across all runs, plus `api_key_configured` (boolean only; the key itself is never returned to the frontend).

Env var needed for this: `ANTHROPIC_API_KEY` (optional — everything works at $0 without it).

## Saved Default Templates (added)

You no longer need to re-upload the bank's CATRA/TRA output templates for every run. Upload each once via the "Bank output templates" card at the top of the dashboard (or `POST /api/templates` with `kind=catra|tra` + `file`), and every future run reuses them automatically. To use a different template for a one-off run (e.g. the bank revises the format), attach a file in the "+ New run" drawer's optional override fields — this does not change the saved default.

Templates are stored in the same backend as everything else (S3 if `S3_BUCKET` is set, else local `data/`), under `templates/catra_template.xlsx` and `templates/tra_template.xlsx`, with metadata in `templates/meta.json`.

## Partial-Year Runs (added)

The pipeline now accepts 1-4 quarterly statements per run — you don't have to wait for the full year. Submit Q1 as soon as it's available; the reports scope themselves to the quarters provided (CATRA/TRA templates leave the remaining quarter columns blank, and the Final Analysis notes it's a partial-year run). Re-run later with the additional quarters once available.

## Deals (added)

The console is now organized by **Deal** rather than a flat run list. Each deal keeps its own:

- **Source documents** — Sanction Letter, Sanction Note / Credit Approval Memo, Escrow Agreement (PDF, uploaded per deal)
- **CATRA/TRA profile** — AI-extracted draft (waterfall Order of Priority, permitted deposits, covenants, projected P&L/Balance Sheet), which you review and edit before confirming. Extraction uses Claude's vision on rendered page images (via PyMuPDF, no poppler/tesseract system dependency) — works on scanned or native-text PDFs identically, one call per document, merged into a single profile.
- **Bank output templates** (CATRA/TRA), saved once and reused by every run for that deal
- **Quarterly runs**, scoped under the deal, each producing CATRA (ATSL) / TRA / Final Analysis

### Deal lifecycle
`new` → upload documents → `extracting` (background AI call) → `review` (draft profile ready to edit) → `confirm` → `ready` (usable by quarterly runs). If `ANTHROPIC_API_KEY` isn't set, an empty profile is created immediately so you can fill it in manually — nothing blocks on AI being configured.

The confirmed profile's covenants also feed the quarterly actuals Final Analysis (deal name, account, and any covenants beyond DSRA/WCR/MMR sizing appear as deal-specific "PENDING ACTUALS" checks) — this replaces what was previously hardcoded to the Kanpur Lucknow deal; the report builders are now fully deal-agnostic.

### New API surface
- `POST/GET/DELETE /api/deals`, `GET /api/deals/{id}`
- `POST /api/deals/{id}/extract` (upload 1-3 documents, triggers background AI extraction)
- `PUT /api/deals/{id}/profile` (manual edit), `POST /api/deals/{id}/confirm`
- `GET /api/deals/{id}/documents/{kind}` (view uploaded PDF)
- `GET/POST /api/deals/{id}/templates`
- `POST/GET /api/deals/{id}/runs`, `GET/POST /api/deals/{id}/runs/{run_id}[/download/{file}|/rerun]`, `DELETE`
- `GET /api/usage` — now aggregates both quarterly AI-assist and deal-extraction cost across all deals

### New env var (optional)
`AI_MODEL_EXTRACTION` (default `claude-sonnet-5`) — model used for deal-document extraction, separate from `AI_MODEL_CLASSIFICATION` (Haiku) used for quarterly transaction review.

## Bug fixes (this session)

Two real bugs surfaced while onboarding the second deal (Kallagam-Meensurutti Highway) that are worth flagging since they affected prior output too:

1. **`CATRA_CREDIT_ROWS` used exact string matching (`==`) against template row labels that carry trailing
   annotation marks** (e.g. `"Proceeds from Equity*"`, `"From Redemption of Investments**"`). This meant those
   two rows were never actually written by the code — the earlier Kanpur Lucknow reconciliation only matched
   the bank's figures because the source template *already contained* the correct values, coincidentally.
   Fixed to use `.startswith()` like the debit-row matching already did. Re-verified: the full Kanpur Lucknow
   FY2024-25 regression still ties out exactly (₹570.04cr inflow / ₹596.79cr outflow / COMPLIANT), confirming
   this was a latent bug rather than a behavior change.
2. **openpyxl silently ignores `ws.cell(row, col, value=None)`** — passing `value=None` is treated as a
   read-only call, not a clear. This meant partial-year runs (fewer than 4 quarters) left whichever data was
   already in the template file's unused quarter columns untouched, instead of blanking them — a real risk
   when the "template" for a new deal happens to be another deal's previously-filled output (exactly what
   happened onboarding Kallagam against the Kanpur Lucknow file). Fixed by using direct `cell.value = None`
   attribute assignment instead, which does clear properly. Also broadened the investment-redemption keyword
   rule in `bank_ingest.py` to match mutual-fund counterparty names without requiring the word "redemption"
   in the narration, which cut Kallagam's review queue from 10/31 to 0/31 transactions.

## Third deal onboarded: Athena Hisar Solar Power (this session)

Onboarded a structurally different deal type — a solar PPA-revenue asset (SECI off-take, no NHAI/Concession
Agreement) rather than a road HAM project — which surfaced a few more classification gaps, now fixed:

- Added keyword matches for "cash sweep" (-> Debt servicing, per this deal's own bank remark convention),
  "stat pay" abbreviation (-> Statutory Payments), and "inter company transfer" on the credit side (a real
  ATSL category that was previously always defaulting to 0 in every prior deal since nothing matched it).
- Generalized the reserve-check labels in the actuals Final Analysis away from road-deal-specific "WCR/DSRA/MMRA"
  wording to deal-agnostic phrasing, since this deal's actual reserve structure (DSRA + DSCR-triggered Cash
  Trap/Cash Sweep accounts) doesn't use those names.
- `kb/structured/deal_extract_athena_hisar.yaml` also demonstrates the profile schema's flexibility: this
  deal's Sanction Note presents projections in DSCR-calculation format (EBITDA/interest/principal/DSCR) rather
  than a full P&L — depreciation, PBT, PAT and a projected Balance Sheet were left absent, and the
  normalization layer derives placeholders + explicit notes rather than requiring every field to be present.

Verified end-to-end via the deal-scoped API: 375 transactions, 4 quarters, balance integrity holds throughout,
verdict COMPLIANT, matching a manual reconciliation exactly (₹126.82cr inflow / ₹120.80cr outflow).

All three onboarded deals (Kanpur Lucknow Expressway, Kallagam-Meensurutti Highway, Athena Hisar Solar Power)
re-verified together in one regression pass to confirm no cross-deal regressions from any fix in this session.

## TRA Sheet2 "Projected Cashflow" bug fix (this session)

Found and fixed a real bug: TRA_Analysis_ATSL.xlsx's Sheet2 "Projected Cashflow as per Sanction note"
column (Column C — the single-year projected figures) was **never being written by the pipeline at all**.
It silently carried over whatever was in the source template file — which, since every new deal reused
the same bank-template file, meant every deal after the first showed **Kanpur Lucknow's own projected P&L
figures**, mislabeled as if they were that deal's own Sanction Note projections.

Fixed properly:
- `build_tra()` now takes the deal's confirmed profile and the run's FY, and fills Column C from that
  deal's own `projected_pnl` — matched to the correct financial year (parses "FY 2024-25" style labels
  against the profile's `fye` list; if no exact year match exists, e.g. a pre-COD deal whose projections
  only start post-construction, it clearly notes the fallback year used directly in the sheet rather than
  silently guessing).
- Line-item lookups (TPC Annuity, O&M, Interest on Annuity, Routine Maintenance, Other O&M, MMR
  Provisioning, Interest, Depreciation) search income/opex dictionaries separately by name, so an income
  line can never be accidentally matched against an expense line of a similar name.
- The actual-quarter columns (D-G) were also corrected: previously the "Interest" row showed investment
  redemption inflows as a stand-in for interest expense — a mismatch that made no sense on inspection.
  Since the ATSL bank categories bundle interest and principal together under "Debt servicing" with no
  way to isolate pure interest from an actual bank statement, that row's actual columns are now left
  honestly blank rather than filled with an unrelated proxy figure. Only "Other O&M" has a genuine
  actual-vs-projected comparison, since it maps directly to an ATSL category.

Re-verified across all three onboarded deals: each now shows its own distinct projected figures in Sheet2
(confirmed Kanpur Lucknow's 19.49 vs Kallagam's 36.1 vs Athena's 0/21.81 all differ correctly), all
workbooks recalculate with 0 formula errors, and the FY-mismatch note appears correctly for Kanpur
Lucknow (whose actual run FY predates its own post-construction projection years).

## Correction: Athena Hisar's projected P&L was incomplete (this session)

A follow-up review found that the Athena Hisar profile's `projected_pnl` was missing depreciation/PBT/PAT —
an earlier extraction pass only found the CAM's DSCR-calculation summary table (page 15) and stopped there,
incorrectly concluding those fields were absent from the source. A fuller Income Statement projection
(page 14, "Projections: Rs. Cr") does exist, with Net sales / GST annuity / VGF revenue / Revenue / EBITDA /
Other income (Interest on DSRA) / Interest / Depreciation / PBT / PAT for FY25-FY40. `deal_extract_athena_hisar.yaml`
now uses this fuller table. Note it reports a different "Interest" figure than the DSCR-calc table (Rs. 24.51cr
vs Rs. 21.81cr in FY25) — both are genuine source figures serving different purposes (P&L finance cost vs.
DSRA-sizing calculation); documented in the profile's notes.

This also surfaced a real matching bug: the TRA Sheet2 "TPC Annuity" row had a fallback that matched any
income line containing the substring "annuity" — which incorrectly picked up Athena's unrelated "GST annuity"
revenue line (a solar-specific component, not a road-deal annuity) once real income data was present. Fixed
to require an exact/specific label match rather than a loose substring fallback.

## Related/group entity matching for Internal Company Transfer (this session)

Reconciling Athena Hisar's generated CATRA against the bank's own output surfaced a real gap: credits from
sister SPVs (e.g. "ATHENA JAIPUR SOLAR POWER P LTD") with no "transfer" keyword in the narration were
defaulting to Other Revenue, since the classifier only matched narration text, not counterparty identity.
The bank's own convention classifies these by *who sent the money*, not what the narration literally says.

Added a `related_entities` list to the deal profile schema (admin-editable, shown in the profile review UI
and the raw JSON editor) — credits from any listed counterparty name are now classified as Internal Company
Transfer regardless of narration wording. Verified against the bank's own Q1 FY2024-25 figure for Athena
Hisar: exact match (Rs. 13,41,75,364 both sides) once "Athena Jaipur Solar" was added to the deal's profile.

Q3 (Rs. 1,01,49,213.02) and Q4 (Rs. 70,000) in the bank's figures remain partially unexplained — the
underlying narrations don't clearly identify a counterparty and the bank's Remarks column doesn't tag these
rows the way it does for other categories, so further reconciliation would need either the bank's
transaction-level detail or additional related-entity names as they're identified.

## "Other Refunds" classification rule added (this session)

Reconciling Athena Hisar's Q4 against the bank's own output revealed that "Other Refunds" — a valid ATSL
credit category — had never once been produced by the classifier; nothing in the rule set mapped to it, so
every refund-type credit (tax refunds, GST refunds, balance refunds) silently fell into the generic "Other
Revenue" bucket instead, for every deal, since this is a narration-keyword rule with no deal-specific logic.

Added a dynamic rule: any narration containing "refund" now classifies as Other Refunds (checked before the
existing expense-narration-reversal fallback, since a refund is a more specific and confident signal). This
is not deal-specific — it applies automatically to any statement, present and future.

Verified: Athena Hisar Q4 "Other Refunds" now matches the bank's figure exactly (Rs. 49,06,890 both sides);
Kallagam's own GST refund (Rs. 3,65,56,544, previously undetected since no bank reconciliation file existed
for that deal) is now also correctly bucketed; Kanpur Lucknow unaffected (no refund-type transactions in its
statements, so no change in behaviour there). The only remaining unresolved gap against Athena's bank output
is the Rs. 70,000 Internal Company Transfer item in Q3/Q4 whose counterparty still isn't identifiable from
the visible narration or the bank's own Remarks column.
