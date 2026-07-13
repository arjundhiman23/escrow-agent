"""Normalize an AI-extracted (or analyst-edited) deal profile into the complete
shape the report builders expect. AI extraction is necessarily partial/messy —
this fills every required field with a safe default and records what it had
to assume, rather than letting a builder crash on a missing key.
"""


def _zeros(n):
    return [0.0] * n


def normalize_profile(raw: dict) -> dict:
    notes = list(raw.get("notes", []))
    p = {
        "deal": dict(raw.get("deal", {})),
        "order_of_priority": list(raw.get("order_of_priority", [])),
        "permitted_deposits": list(raw.get("permitted_deposits", [])),
        "sanction_waterfall": list(raw.get("sanction_waterfall", [])),
        "covenants": list(raw.get("covenants", [])),
        "related_entities": list(raw.get("related_entities", [])),
        "annuity_terms": dict(raw.get("annuity_terms", {})),
    }

    # --- waterfall divergences: derive by comparing codes, since AI extraction
    # doesn't produce this directly the way the manual Kanpur Lucknow build did ---
    op_codes = {r.get("code") for r in p["order_of_priority"] if r.get("code")}
    sw_codes = {r.get("maps_to") for r in p["sanction_waterfall"] if r.get("maps_to")}
    divergences = []
    only_in_agreement = op_codes - sw_codes
    only_in_sanction = sw_codes - op_codes
    if only_in_agreement:
        divergences.append(
            f"Escrow Agreement Order of Priority includes {', '.join(sorted(only_in_agreement))} "
            f"with no corresponding Sanction Letter waterfall step — confirm treatment.")
    if only_in_sanction:
        divergences.append(
            f"Sanction Letter waterfall references {', '.join(sorted(only_in_sanction))} "
            f"not found in the Escrow Agreement Order of Priority — confirm treatment.")
    if not p["order_of_priority"]:
        divergences.append("No Order of Priority was extracted from the Escrow Agreement — "
                           "CATRA debit categories could not be generated; upload/re-extract.")
    p["waterfall_divergences"] = divergences or ["None identified."]

    # --- projected P&L: ensure every field the builders index exists ---
    raw_pnl = raw.get("projected_pnl") or {}
    fye = raw_pnl.get("fye") or []
    n = len(fye)
    if n == 0:
        notes.append("No projected P&L was extracted from the Sanction Note/CAM — "
                     "TRA analysis and reserve-adequacy checks will be empty until this is added.")
    income = raw_pnl.get("income") or {}
    opex = raw_pnl.get("opex") or {}
    income_total = raw_pnl.get("income_total") or (
        [round(sum(vals[i] for vals in income.values()), 2) for i in range(n)] if income else _zeros(n))
    opex_total = raw_pnl.get("opex_total") or (
        [round(sum(vals[i] for vals in opex.values()), 2) for i in range(n)] if opex else _zeros(n))
    interest = raw_pnl.get("interest") or _zeros(n)
    if not raw_pnl.get("interest") and n:
        notes.append("Interest line not extracted from projected P&L — defaulted to 0; DSRA/WCR checks will read as trivially compliant until corrected.")
    depreciation = raw_pnl.get("depreciation") or _zeros(n)
    ebitda = raw_pnl.get("ebitda") or [round(income_total[i] - opex_total[i], 2) for i in range(n)]
    pbt = raw_pnl.get("pbt") or [round(ebitda[i] - interest[i] - depreciation[i], 2) for i in range(n)]
    csr = raw_pnl.get("csr") or _zeros(n)
    tax = raw_pnl.get("tax") or _zeros(n)
    pat = raw_pnl.get("pat") or [round(pbt[i] - csr[i] - tax[i], 2) for i in range(n)]
    p["projected_pnl"] = {
        "fye": fye, "income": income, "opex": opex, "income_total": income_total, "opex_total": opex_total,
        "ebitda": ebitda, "interest": interest, "depreciation": depreciation, "pbt": pbt,
        "csr": csr, "tax": tax, "pat": pat,
    }

    # --- projected balance sheet: same fill-with-zero-and-note approach ---
    raw_bs = raw.get("projected_balance_sheet") or {}
    bs = {"fye": fye}
    for key in ("long_term_debt", "cmltd", "dsra_fund", "mmr_fund", "working_capital_reserve", "cash_and_bank"):
        if raw_bs.get(key):
            bs[key] = raw_bs[key]
        else:
            bs[key] = _zeros(n)
            if n:
                notes.append(f"Balance sheet line '{key}' not extracted — defaulted to 0 for all years; "
                             f"reserve-adequacy checks involving it need analyst verification.")
    p["projected_balance_sheet"] = bs

    p["notes"] = notes
    return p


def representative_years(fye: list, count=5) -> list:
    """Pick up to `count` evenly-spaced FY labels for summary tables,
    instead of a hardcoded list that only matches one specific deal."""
    if not fye:
        return []
    if len(fye) <= count:
        return fye
    step = (len(fye) - 1) / (count - 1)
    idx = sorted({round(i * step) for i in range(count)})
    return [fye[i] for i in idx]
