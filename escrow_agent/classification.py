"""Transaction classification (BRD §6.2).

Hybrid approach:
  Tier 1 — keyword rules from the TRA/CATRA taxonomy (config/categories.yaml)
  Tier 2 — AI semantic analysis of narrations (Anthropic API), if enabled
  Tier 3 — analyst review queue for ambiguous/missing narrations

Conflict detection: narrations whose keywords belong to the OPPOSITE side of
the transaction type (e.g. a Credit whose narration reads like Debt Servicing)
are flagged POTENTIALLY_INACCURATE and escalated separately for manual
override, distinct from the ordinary review queue (BRD §6.2 tiered fallback).
"""
from __future__ import annotations

from .models import Taxonomy, Transaction, resolve_ai_enabled

REVIEW_AMBIGUOUS = "AMBIGUOUS_OR_MISSING_NARRATION"
REVIEW_MULTI = "MULTIPLE_RULE_MATCHES"
REVIEW_LOW_CONF = "LOW_AI_CONFIDENCE"
CONFLICT_FLAG = "POTENTIALLY_INACCURATE_NARRATION"


def _keyword_hits(narration: str, categories) -> list:
    n = narration.lower()
    hits = []
    for cat in categories:
        if any(k in n for k in cat.keywords):
            hits.append(cat)
    return hits


def classify(txns: list[Transaction], taxonomy: Taxonomy, settings: dict,
             audit=None, usage=None) -> None:
    ai_enabled = resolve_ai_enabled(settings)
    by_code = taxonomy.by_code()
    ai_pending: list[Transaction] = []

    for t in txns:
        narration = t.narration.strip()

        # --- Tier 3 (early): missing narration -> review queue -------------
        if not narration:
            t.review_reason = REVIEW_AMBIGUOUS
            t.classification_source = "UNCLASSIFIED"
            continue

        same_side = _keyword_hits(narration, taxonomy.for_side(t.txn_type))
        opp_side = _keyword_hits(
            narration, taxonomy.for_side("Cr" if t.txn_type == "Dr" else "Dr"))

        # --- Conflict: narration reads like the opposite side --------------
        if not same_side and opp_side:
            t.conflict_reason = (
                f"{CONFLICT_FLAG}: {t.txn_type} transaction but narration matches "
                f"{'/'.join(c.name for c in opp_side[:2])} "
                f"({'inflow' if t.txn_type == 'Dr' else 'outflow'} category)."
            )
            t.classification_source = "UNCLASSIFIED"
            continue

        # --- Tier 1: rule-based --------------------------------------------
        if len(same_side) == 1:
            cat = same_side[0]
            t.category_code, t.category_name = cat.code, cat.name
            t.classification_source, t.confidence = "RULE", 1.0
            continue
        if len(same_side) > 1:
            # deterministic tiebreak: the category with the strictly longest
            # matching keyword wins (a more specific phrase beats a substring,
            # e.g. "debt service reserve" over "debt service")
            n = narration.lower()
            best = sorted(((max(len(k) for k in c.keywords if k in n), c)
                           for c in same_side), key=lambda x: -x[0])
            if best[0][0] > best[1][0]:
                cat = best[0][1]
                t.category_code, t.category_name = cat.code, cat.name
                t.classification_source, t.confidence = "RULE", 0.9
            elif ai_enabled:
                ai_pending.append(t)     # AI tiebreak
            else:
                t.review_reason = REVIEW_MULTI + ": " + ", ".join(c.code for c in same_side)
                t.classification_source = "UNCLASSIFIED"
            continue

        # --- No rule hit: Tier 2 (AI) or Tier 3 (review) --------------------
        if ai_enabled:
            ai_pending.append(t)
        else:
            t.review_reason = REVIEW_AMBIGUOUS
            t.classification_source = "UNCLASSIFIED"

    if ai_pending:
        _ai_classify(ai_pending, taxonomy, settings, usage=usage)
        # anything AI couldn't resolve confidently -> review queue
        min_conf = float(settings["classification"].get("min_ai_confidence", 0.75))
        for t in ai_pending:
            if t.classification_source == "AI" and t.confidence < min_conf:
                t.review_reason = REVIEW_LOW_CONF
                t.category_code = t.category_name = ""
                t.classification_source = "UNCLASSIFIED"
            elif t.classification_source != "AI" and not t.review_reason:
                t.review_reason = REVIEW_AMBIGUOUS
                t.classification_source = "UNCLASSIFIED"
        _ = by_code  # (kept for future per-code validation hooks)

    if audit:
        classified = sum(1 for t in txns if t.category_code)
        audit.log("CLASSIFICATION",
                  f"{classified}/{len(txns)} classified "
                  f"(RULE: {sum(t.classification_source == 'RULE' for t in txns)}, "
                  f"AI: {sum(t.classification_source == 'AI' for t in txns)}); "
                  f"review queue: {sum(t.needs_review for t in txns)}; "
                  f"escalated conflicts: {sum(t.is_escalated_conflict for t in txns)}. "
                  f"AI enabled: {ai_enabled}.")


# ---------------------------------------------------------------------------
# Tier 2: AI semantic classification via Anthropic API
# ---------------------------------------------------------------------------
def _ai_classify(txns: list[Transaction], taxonomy: Taxonomy, settings: dict, usage=None) -> None:
    import json
    try:
        import anthropic
    except ImportError:
        for t in txns:
            t.review_reason = t.review_reason or REVIEW_AMBIGUOUS
        return

    client = anthropic.Anthropic()
    model = settings["classification"].get("ai_model_classification", "claude-haiku-4-5-20251001")
    batch_size = int(settings["classification"].get("ai_batch_size", 25))
    valid_codes = set(taxonomy.by_code())

    taxonomy_text = "\n".join(
        f"- {c.code} ({c.side}): {c.name} — {c.description}" for c in taxonomy.all())

    for start in range(0, len(txns), batch_size):
        batch = txns[start:start + batch_size]
        items = "\n".join(
            f'{i}. type={t.txn_type} amount={t.amount:.2f} narration="{t.narration}"'
            for i, t in enumerate(batch))
        prompt = f"""You are classifying escrow account bank transactions into a TRA/CATRA taxonomy.

Categories (Cr transactions must map to an inflow code, Dr to an outflow code):
{taxonomy_text}

Transactions:
{items}

Respond with ONLY a JSON array, one object per transaction, in order:
[{{"i": 0, "code": "<category code or null>", "confidence": 0.0-1.0}}]
Use null when the narration is too ambiguous to classify. No other text."""
        try:
            resp = client.messages.create(
                model=model, max_tokens=64 * len(batch) + 200,
                messages=[{"role": "user", "content": prompt}])
            if usage is not None and getattr(resp, "usage", None):
                usage.add(model, resp.usage.input_tokens, resp.usage.output_tokens)
            text = "".join(b.text for b in resp.content if b.type == "text")
            text = text.strip().removeprefix("```json").removeprefix("```").removesuffix("```").strip()
            results = json.loads(text)
        except Exception:
            continue  # batch fails -> those txns fall through to review queue

        by_code = taxonomy.by_code()
        for r in results:
            try:
                t = batch[int(r["i"])]
            except (KeyError, IndexError, ValueError, TypeError):
                continue
            code = r.get("code")
            if code in valid_codes:
                cat = by_code[code]
                expected = "inflow" if t.txn_type == "Cr" else "outflow"
                if cat.side != expected:   # AI must respect the Dr/Cr side
                    continue
                t.category_code, t.category_name = cat.code, cat.name
                t.classification_source = "AI"
                t.confidence = float(r.get("confidence", 0.0))
