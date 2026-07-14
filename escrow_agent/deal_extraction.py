"""Extract a structured deal profile from a new deal's source documents using
Claude's vision (page images via PyMuPDF — works on scanned or native-text
PDFs identically, no OCR system dependency). One call per document type; the
three partial results are merged into a single draft profile for the analyst
to review/edit before it's used to generate CATRA/TRA/Final Analysis.

Schema mirrors kb/structured/deal_extract.yaml (built manually for Kanpur
Lucknow) so the same catra_builder/tra_builder/final_analysis modules work
unchanged for any deal.
"""
import json
from escrow_agent.pdf_render import pdf_to_images
from escrow_agent.ai_usage import UsageTracker

MAX_PAGES_PER_DOC = 60   # generous cap; flags truncation rather than failing silently

SCHEMA_NOTE = """Return ONLY a single JSON object (no prose, no markdown fences) with this shape
(omit any top-level key you find no evidence for — do not invent values):

{
  "deal": {"borrower": str, "sponsor": str, "lender": str, "authority": str, "facility": str,
           "facility_amount": str, "project": str, "concession_agreement_signed": str},
  "order_of_priority": [ {"priority": int, "code": SHORT_CODE, "item": "a"/"b".., "name": str,
                          "sub_account": str, "agreement_item": str, "desc": str} , ... ],
  "permitted_deposits": [ {"code": SHORT_CODE, "item": "a"/"b".., "name": str} , ... ],
  "sanction_waterfall": [ {"step": int, "text": str, "maps_to": SHORT_CODE (best-guess match to an
                          order_of_priority code)} , ... ],
  "covenants": [ {"id": "CVn", "source": "clause ref", "rule": str} , ... ],
  "projected_pnl": {"fye": [...FY labels...], "income": {line: [values...]}, "opex": {line: [values...]},
                    "interest": [...], "depreciation": [...], "pbt": [...], "pat": [...]},
  "projected_balance_sheet": {"fye": [...], "long_term_debt": [...], "cmltd": [...], "dsra_fund": [...],
                              "mmr_fund": [...], "working_capital_reserve": [...], "cash_and_bank": [...]},
  "annuity_terms": {"structure": str, "first_annuity": str, "termination_payment": str},
  "notes": [str, ...]   // anything ambiguous, low-confidence, or worth an analyst's attention
}

SHORT_CODE means a short uppercase snake identifier you invent consistently across
order_of_priority / sanction_waterfall / permitted_deposits (e.g. TAXES, DEBT_SERV, DSRA, EQUITY).
Use "notes" liberally for anything uncertain — do not guess numeric figures."""

DOC_PROMPTS = {
    "escrow_agreement": (
        "This is an Escrow Agreement for an infrastructure financing deal. Extract the debit-side "
        "Order of Priority / waterfall for withdrawals (usually a numbered/lettered clause listing "
        "taxes, construction, O&M, debt service, reserves, distributions in priority order), the "
        "permitted credit-side deposits, and any covenants governing reserve sizing (DSRA/WCR/MMR) "
        "or payment conditions. " + SCHEMA_NOTE
    ),
    "sanction_letter": (
        "This is a Bank Sanction Letter for an infrastructure financing deal. Extract the cash-flow "
        "waterfall / order of application of funds (often a numbered list under a 'cash flow waterfall' "
        "or 'usage of revenues' clause), any conditions precedent/subsequent relevant to escrow "
        "operation, and facility terms (amount, borrower, lender). " + SCHEMA_NOTE
    ),
    "sanction_note": (
        "This is a Credit Approval Memo / Sanction Note for an infrastructure financing deal. Extract "
        "the projected Profit & Loss statement and projected Balance Sheet (by financial year — look "
        "for a multi-year table with income/opex/EBITDA/interest/PBT/PAT rows, and a balance sheet "
        "table with debt/reserve/cash rows), and the annuity or repayment schedule if present. "
        + SCHEMA_NOTE
    ),
}


def extract_document(client, model, doc_kind, pdf_bytes, tracker: UsageTracker, max_pages=MAX_PAGES_PER_DOC):
    images, n_pages = pdf_to_images(pdf_bytes, dpi=110, max_pages=max_pages)
    truncated = n_pages > len(images)
    content = [{"type": "image", "source": {"type": "base64", "media_type": im["media_type"],
                                            "data": im["base64"]}} for im in images]
    content.append({"type": "text", "text": DOC_PROMPTS[doc_kind]})
    resp = client.messages.create(model=model, max_tokens=16000, messages=[{"role": "user", "content": content}])
    tracker.record("deal_extraction", model, resp.usage)
    text = "".join(b.text for b in resp.content if b.type == "text").strip()
    text = text.removeprefix("```json").removeprefix("```").removesuffix("```").strip()
    hit_token_limit = getattr(resp, "stop_reason", None) == "max_tokens"
    try:
        parsed = json.loads(text)
    except json.JSONDecodeError:
        reason = (f"response was cut off at the {resp.usage.output_tokens}-token limit before the JSON "
                  f"could be completed — increase max_tokens" if hit_token_limit else
                  "did not return valid JSON for an unrelated reason — raw output saved for review")
        parsed = {"notes": [f"AI extraction for {doc_kind} failed: {reason}"], "_raw": text}
    if truncated:
        parsed.setdefault("notes", []).append(
            f"Document has {n_pages} pages; only the first {len(images)} were processed — "
            f"review whether later pages contain relevant clauses.")
    return parsed


def merge_profiles(parts: dict) -> dict:
    """Merge per-document partial extractions into one draft profile.
    List-valued keys concatenate (dedup by an identity key where sensible);
    dict-valued keys take the first non-empty source; conflicts go to notes."""
    merged = {"deal": {}, "order_of_priority": [], "permitted_deposits": [], "sanction_waterfall": [],
              "covenants": [], "projected_pnl": {}, "projected_balance_sheet": {}, "annuity_terms": {},
              "notes": [], "sources": {}}
    for doc_kind, data in parts.items():
        if not isinstance(data, dict):
            continue
        merged["sources"][doc_kind] = "ok" if "_raw" not in data else "unparsed"
        for k, v in data.get("deal", {}).items():
            if v and not merged["deal"].get(k):
                merged["deal"][k] = v
        for key in ("order_of_priority", "permitted_deposits", "sanction_waterfall", "covenants"):
            for row in data.get(key, []) or []:
                merged[key].append(row)
        for key in ("projected_pnl", "projected_balance_sheet", "annuity_terms"):
            if data.get(key) and not merged[key]:
                merged[key] = data[key]
        for note in data.get("notes", []) or []:
            merged["notes"].append(f"[{doc_kind}] {note}")
        if "_raw" in data:
            merged["notes"].append(f"[{doc_kind}] extraction returned non-JSON output — needs manual review")
    # de-dup order_of_priority by (priority, name) keeping first occurrence
    seen = set()
    deduped = []
    for row in merged["order_of_priority"]:
        key = (row.get("priority"), row.get("name"))
        if key not in seen:
            seen.add(key)
            deduped.append(row)
    merged["order_of_priority"] = sorted(deduped, key=lambda r: r.get("priority") or 999)
    return merged


def run_extraction(documents: dict, model="claude-sonnet-5"):
    """documents: {"escrow_agreement": bytes|None, "sanction_letter": bytes|None, "sanction_note": bytes|None}
    Returns (merged_profile, usage_summary_dict)."""
    import anthropic
    client = anthropic.Anthropic()
    tracker = UsageTracker()
    parts = {}
    for kind, data in documents.items():
        if not data:
            continue
        parts[kind] = extract_document(client, model, kind, data, tracker)
    return merge_profiles(parts), tracker.summary()
