"""Waterfall mechanism & sanction condition validation (BRD §6.4).

Debit-side and credit-side transactions are validated INDEPENDENTLY, and each
violation carries a specific type reported separately in the exceptions log:

  Debit side:
    DEBIT_OUT_OF_SEQUENCE       payment to a lower-priority category in a cycle
                                where a higher-priority category (with a
                                projected obligation) received no payment
    DEBIT_UNAUTHORIZED_PAYMENT  debit category absent from the waterfall
    CONDITION_BREACH_<id>       machine-checkable sanction condition failed

  Credit side:
    CREDIT_UNPERMITTED_SOURCE   inflow category not in permitted sources

Unclassified / review-queue transactions are excluded from waterfall checks
(they surface via the review queue instead) — checking them would produce
noise on top of an unresolved classification.
"""
from __future__ import annotations

from collections import defaultdict

from .models import Transaction

CYCLE_FMT = "%Y-%m"


def validate(txns: list[Transaction], sanction: dict, settings: dict,
             audit=None) -> list[dict]:
    """Returns exception records:
    {side, violation_type, cycle, row_no, date, category, amount, detail}"""
    exceptions: list[dict] = []
    priority = {lv["category_code"]: int(lv["priority"]) for lv in sanction["waterfall"]}
    permitted_inflows = set(sanction["permitted_inflow_sources"])

    classified = [t for t in txns if t.category_code and not t.is_duplicate]

    # ---- Credit side ------------------------------------------------------
    for t in classified:
        if t.txn_type == "Cr" and t.category_code not in permitted_inflows:
            t.waterfall_violation = "CREDIT_UNPERMITTED_SOURCE"
            exceptions.append(_rec(t, "Credit", "CREDIT_UNPERMITTED_SOURCE",
                                   f"Inflow category '{t.category_name}' is not a "
                                   f"permitted source under the Sanction Letter."))

    # ---- Debit side: unauthorized categories ------------------------------
    for t in classified:
        if t.txn_type == "Dr" and t.category_code not in priority:
            t.waterfall_violation = "DEBIT_UNAUTHORIZED_PAYMENT"
            exceptions.append(_rec(t, "Debit", "DEBIT_UNAUTHORIZED_PAYMENT",
                                   f"Outflow category '{t.category_name}' does not "
                                   f"appear in the sanctioned waterfall."))

    # ---- Debit side: priority-of-payment sequence ---------------------------
    # A payment to priority level N is flagged out-of-sequence when a
    # HIGHER-priority category with a projected obligation for that quarter
    # received NO payment anywhere in the quarter. (Obligations are projected
    # per quarter, so quarterly-frequency items — FD placements, reserve
    # top-ups — are not falsely treated as skipped in months where they simply
    # weren't due. Same-cycle ordering rules, e.g. surplus only after debt
    # service in the month, are expressed as 'surplus_requires_prior'
    # conditions.)
    cycles: dict[str, list[Transaction]] = defaultdict(list)
    paid_in_quarter: dict[str, set] = defaultdict(set)
    for t in classified:
        if t.txn_type == "Dr" and t.category_code in priority:
            cycles[t.txn_date.strftime(CYCLE_FMT)].append(t)
            paid_in_quarter[t.quarter].add(t.category_code)

    proj = sanction.get("projections", {})
    for cycle, cycle_txns in sorted(cycles.items()):
        quarter = cycle_txns[0].quarter
        paid_codes = paid_in_quarter[quarter]
        obligated = {code for code, amt in
                     proj.get(quarter, {}).get("outflows", {}).items()
                     if amt and code in priority}
        for t in cycle_txns:
            lvl = priority[t.category_code]
            skipped = [c for c in obligated
                       if priority[c] < lvl and c not in paid_codes]
            if skipped:
                t.waterfall_violation = t.waterfall_violation or "DEBIT_OUT_OF_SEQUENCE"
                names = ", ".join(f"{c} (priority {priority[c]})" for c in sorted(skipped, key=priority.get))
                exceptions.append(_rec(t, "Debit", "DEBIT_OUT_OF_SEQUENCE",
                                       f"Paid priority-{lvl} '{t.category_name}' in cycle {cycle} "
                                       f"while higher-priority categories obligated for {quarter} "
                                       f"received no payment in the quarter: {names}."))

    # ---- Machine-checkable sanction conditions -----------------------------
    manual_conditions = []
    for cond in sanction.get("conditions", []):
        ctype = cond.get("type")
        if ctype == "surplus_requires_prior":
            exceptions.extend(_check_surplus_requires_prior(classified, cond, priority))
        elif ctype == "category_cap_per_quarter":
            exceptions.extend(_check_category_cap(classified, cond))
        else:
            manual_conditions.append(cond)

    if audit:
        by_type = defaultdict(int)
        for e in exceptions:
            by_type[e["violation_type"]] += 1
        audit.log("WATERFALL_VALIDATION",
                  f"{len(exceptions)} violations flagged "
                  f"({dict(by_type) if by_type else 'none'}); "
                  f"{len(manual_conditions)} conditions require manual verification.")

    return exceptions


def manual_conditions(sanction: dict) -> list[dict]:
    return [c for c in sanction.get("conditions", [])
            if c.get("type") not in ("surplus_requires_prior", "category_cap_per_quarter")]


def _check_surplus_requires_prior(txns, cond, priority) -> list[dict]:
    """Target category (e.g. SURPLUS_DIST) is only permitted in a cycle where
    the required category (e.g. DEBT_SERV) has already been paid."""
    req = cond["params"]["required_code"]
    tgt = cond["params"]["target_code"]
    out = []
    cycles = defaultdict(list)
    for t in txns:
        if t.txn_type == "Dr":
            cycles[t.txn_date.strftime(CYCLE_FMT)].append(t)
    for cycle, cycle_txns in cycles.items():
        req_dates = [t.txn_date for t in cycle_txns if t.category_code == req]
        for t in cycle_txns:
            if t.category_code != tgt:
                continue
            if not req_dates or min(req_dates) > t.txn_date:
                vtype = f"CONDITION_BREACH_{cond['id']}"
                t.waterfall_violation = t.waterfall_violation or vtype
                out.append(_rec(t, "Debit", vtype,
                                f"{cond.get('description', '')} — '{tgt}' paid on "
                                f"{t.txn_date:%d-%m-%Y} without prior '{req}' payment in cycle {cycle}."))
    _ = priority
    return out


def _check_category_cap(txns, cond) -> list[dict]:
    code = cond["params"]["category_code"]
    cap = float(cond["params"]["cap"])
    out = []
    totals = defaultdict(float)
    for t in txns:
        if t.txn_type == "Dr" and t.category_code == code:
            totals[t.quarter] += t.amount
            if totals[t.quarter] > cap:
                vtype = f"CONDITION_BREACH_{cond['id']}"
                t.waterfall_violation = t.waterfall_violation or vtype
                out.append(_rec(t, "Debit", vtype,
                                f"{cond.get('description', '')} — cumulative {t.quarter} total "
                                f"{totals[t.quarter]:,.2f} exceeds cap {cap:,.2f} at this transaction."))
    return out


def _rec(t: Transaction, side: str, vtype: str, detail: str) -> dict:
    return {"side": side, "violation_type": vtype, "cycle": t.txn_date.strftime(CYCLE_FMT),
            "row_no": t.row_no, "date": t.txn_date, "quarter": t.quarter,
            "category": t.category_name or "(unclassified)", "amount": t.amount,
            "detail": detail}
