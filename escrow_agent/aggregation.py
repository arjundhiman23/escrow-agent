"""Quarter-wise aggregation (BRD §6.3) and actual-vs-projected variance
analysis (BRD §6.5)."""
from __future__ import annotations

from collections import defaultdict

from .models import Taxonomy, Transaction

QUARTERS = ("Q1", "Q2", "Q3", "Q4")


# ---------------------------------------------------------------------------
# §6.3 Quarter-wise aggregation
# ---------------------------------------------------------------------------
def aggregate(txns: list[Transaction], taxonomy: Taxonomy, audit=None) -> dict:
    """Returns:
    {
      "totals":   {Q: {"Dr": x, "Cr": y}},                    (all non-duplicate txns)
      "by_category": {Q: {code: amount}},                     (classified txns)
      "internal_transfers": {Q: {"Dr": x, "Cr": y}},          (flagged separately, §5.4)
      "unclassified": {Q: {"Dr": x, "Cr": y}},
    }
    Duplicates are excluded from all aggregates (they remain visible on the
    Transactions sheet with a DUPLICATE flag)."""
    totals = {q: {"Dr": 0.0, "Cr": 0.0} for q in QUARTERS}
    by_cat = {q: defaultdict(float) for q in QUARTERS}
    internal = {q: {"Dr": 0.0, "Cr": 0.0} for q in QUARTERS}
    unclassified = {q: {"Dr": 0.0, "Cr": 0.0} for q in QUARTERS}

    for t in txns:
        if t.is_duplicate or t.quarter not in totals:
            continue
        totals[t.quarter][t.txn_type] += t.amount
        if t.is_internal_transfer:
            internal[t.quarter][t.txn_type] += t.amount
        if t.category_code:
            by_cat[t.quarter][t.category_code] += t.amount
        else:
            unclassified[t.quarter][t.txn_type] += t.amount

    result = {"totals": totals,
              "by_category": {q: dict(v) for q, v in by_cat.items()},
              "internal_transfers": internal,
              "unclassified": unclassified}
    if audit:
        fy_dr = sum(v["Dr"] for v in totals.values())
        fy_cr = sum(v["Cr"] for v in totals.values())
        audit.log("AGGREGATION",
                  f"Quarter-wise totals computed. FY Dr: {fy_dr:,.2f}, FY Cr: {fy_cr:,.2f}.")
    _ = taxonomy
    return result


# ---------------------------------------------------------------------------
# §6.5 Actual vs projected variance
# ---------------------------------------------------------------------------
def variance_analysis(agg: dict, sanction: dict, taxonomy: Taxonomy,
                      settings: dict, audit=None) -> dict:
    """Returns:
      rows: [{quarter, side, code, name, actual, projected, var_abs, var_pct,
              material, remarks}]
      qoq:  [{side, code, name, quarter_pair, actual_delta, projected_delta}]
      exceptions_summary: {Q: [top material rows]}"""
    pct_thr = float(settings["variance"]["material_pct_threshold"])
    abs_thr = float(settings["variance"]["material_abs_threshold"])
    top_n = int(settings["variance"].get("top_exceptions_per_quarter", 3))
    proj = sanction.get("projections", {})
    by_code = taxonomy.by_code()

    rows = []
    for q in QUARTERS:
        actual_cat = agg["by_category"].get(q, {})
        proj_q = proj.get(q, {"inflows": {}, "outflows": {}})
        for side, side_key in (("Inflow", "inflows"), ("Outflow", "outflows")):
            cats = taxonomy.inflows if side == "Inflow" else taxonomy.outflows
            proj_side = proj_q.get(side_key, {})
            codes = [c.code for c in cats]
            # include projected codes not in taxonomy order (defensive)
            codes += [c for c in proj_side if c not in codes]
            for code in codes:
                actual = float(actual_cat.get(code, 0.0))
                projected = float(proj_side.get(code, 0.0))
                if actual == 0.0 and projected == 0.0:
                    continue
                var_abs = actual - projected
                var_pct = (var_abs / projected * 100.0) if projected else None
                material = (abs(var_abs) >= abs_thr or
                            (var_pct is not None and abs(var_pct) >= pct_thr) or
                            (projected == 0.0 and actual != 0.0))
                rows.append({
                    "quarter": q, "side": side, "code": code,
                    "name": by_code[code].name if code in by_code else code,
                    "actual": actual, "projected": projected,
                    "var_abs": var_abs, "var_pct": var_pct, "material": material,
                    "remarks": _remark(actual, projected, var_pct, material, pct_thr),
                })

    # Quarter-on-quarter deltas in both actuals and projections (§6.5)
    qoq = []
    for side, side_key in (("Inflow", "inflows"), ("Outflow", "outflows")):
        cats = taxonomy.inflows if side == "Inflow" else taxonomy.outflows
        for c in cats:
            for i in range(1, len(QUARTERS)):
                q_prev, q_cur = QUARTERS[i - 1], QUARTERS[i]
                a_prev = float(agg["by_category"].get(q_prev, {}).get(c.code, 0.0))
                a_cur = float(agg["by_category"].get(q_cur, {}).get(c.code, 0.0))
                p_prev = float(proj.get(q_prev, {}).get(side_key, {}).get(c.code, 0.0))
                p_cur = float(proj.get(q_cur, {}).get(side_key, {}).get(c.code, 0.0))
                if any((a_prev, a_cur, p_prev, p_cur)):
                    qoq.append({"side": side, "code": c.code, "name": c.name,
                                "quarter_pair": f"{q_prev}\u2192{q_cur}",
                                "actual_delta": a_cur - a_prev,
                                "projected_delta": p_cur - p_prev})

    # Consolidated exceptions summary: most material deviations per quarter (§6.5)
    exceptions_summary = {}
    for q in QUARTERS:
        material_rows = [r for r in rows if r["quarter"] == q and r["material"]]
        material_rows.sort(key=lambda r: abs(r["var_abs"]), reverse=True)
        exceptions_summary[q] = material_rows[:top_n]

    if audit:
        audit.log("VARIANCE_ANALYSIS",
                  f"{len(rows)} category-quarter variance rows computed; "
                  f"{sum(r['material'] for r in rows)} material deviations "
                  f"(thresholds: {pct_thr}% / {abs_thr:,.0f}).")
    return {"rows": rows, "qoq": qoq, "exceptions_summary": exceptions_summary}


def _remark(actual, projected, var_pct, material, pct_thr) -> str:
    if not material:
        return "Within threshold."
    if projected == 0.0 and actual != 0.0:
        return "Unbudgeted: actuals recorded against a category with no projection this quarter."
    if actual == 0.0 and projected != 0.0:
        return "No actuals against a projected obligation/inflow — verify timing or missed booking."
    direction = "above" if actual > projected else "below"
    return (f"Material deviation: actuals {direction} projection by "
            f"{abs(var_pct):.1f}% (threshold {pct_thr:.0f}%).")
