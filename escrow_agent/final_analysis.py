"""Build the Final Analysis document.

Summarises the KB (sanction letter + sanction note/CAM + escrow agreement),
cross-checks the CATRA/TRA outputs against the Agreement clauses, and issues a
per-item and overall verdict: OVERBREACH / UNDERBREACH / COMPLIANT.

Breach semantics used throughout the agent:
  OVERBREACH  — utilisation/outflow above what the clause permits (cap exceeded,
                unauthorised withdrawal, payment out of Order of Priority,
                distribution made while senior obligations unfunded).
  UNDERBREACH — funding/obligation below what the clause requires (reserve
                shortfall, unfunded senior priority, deposit not routed to escrow).
On the initial inputs (no bank statement yet) the verdicts are on PROJECTED basis
from the CAM; the same checks re-run on actuals when the statement is ingested.
"""
from datetime import date
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

DARK = RGBColor(0x1F, 0x4E, 0x79)
RED = RGBColor(0xC0, 0x00, 0x00)
GREEN = RGBColor(0x00, 0x7A, 0x33)
AMBER = RGBColor(0xB8, 0x86, 0x00)


def _style(doc):
    s = doc.styles["Normal"]
    s.font.name = "Arial"
    s.font.size = Pt(10)


def _h(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for r in h.runs:
        r.font.name = "Arial"
        r.font.color.rgb = DARK
    return h


def _table(doc, headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        c.text = h
        for p in c.paragraphs:
            for r in p.runs:
                r.font.bold = True
                r.font.size = Pt(9)
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = "" if v is None else str(v)
            for p in cells[i].paragraphs:
                for r in p.runs:
                    r.font.size = Pt(9)
    if widths:
        for i, w in enumerate(widths):
            for row in t.rows:
                row.cells[i].width = Inches(w)
    return t


def _status_color(t, col):
    for row in t.rows[1:]:
        cell = row.cells[col]
        txt = cell.text.strip()
        if txt == "COMPLIANT":
            color = GREEN
        elif txt in ("UNDERBREACH", "OVERBREACH"):
            color = RED
        else:  # MARGINAL, N-A, PENDING ACTUALS, ATTENTION
            color = AMBER
        for p in cell.paragraphs:
            for r in p.runs:
                r.font.color.rgb = color
                r.font.bold = True


def build_final_analysis(extract, reserve_checks, out_path="output/Final_Analysis.docx"):
    doc = Document()
    _style(doc)

    deal_info = extract.get("deal", {})
    borrower = deal_info.get("borrower", "Borrower not extracted — check source documents")
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run(f"FINAL ANALYSIS — ESCROW ACCOUNT COMPLIANCE\n{borrower}")
    r.font.size = Pt(16); r.font.bold = True; r.font.color.rgb = DARK
    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = sub.add_run(f"Generated from Knowledge Base (Sanction Letter, Credit Approval Memo, Escrow Agreement) — {date.today().strftime('%d-%m-%Y')}")
    r.font.size = Pt(9); r.font.italic = True

    # ---------------- 1. Executive summary & verdict ----------------
    _h(doc, "1. Executive Summary and Verdict")
    fails = [c for c in reserve_checks if c["status"] in ("UNDERBREACH", "OVERBREACH")]
    marginal = [c for c in reserve_checks if c["status"] == "MARGINAL"]
    over = [c for c in fails if c["status"] == "OVERBREACH"]
    under = [c for c in fails if c["status"] == "UNDERBREACH"]
    if over:
        verdict, vcol = "OVERBREACH", RED
    elif under:
        verdict, vcol = "UNDERBREACH", RED
    else:
        verdict, vcol = "COMPLIANT", GREEN
    p = doc.add_paragraph()
    p.add_run("Overall account status (projected basis): ").font.size = Pt(11)
    vr = p.add_run(verdict)
    vr.font.size = Pt(13); vr.font.bold = True; vr.font.color.rgb = vcol
    doc.add_paragraph(
        f"Basis: {len(reserve_checks)} covenant checks computed from the CAM base-case projections against the "
        f"Escrow Agreement requirements — {len(fails)} fail ({len(under)} UNDERBREACH, {len(over)} OVERBREACH), "
        f"{len(marginal)} are marginal (within proxy tolerance, analyst to verify against the actual debt-service "
        "schedule), and the rest are compliant or not applicable in the year concerned. No actual bank statement "
        "has been ingested yet — once the escrow account statement is provided, every check below re-runs on "
        "actuals and the verdict is restated on an actuals basis. UNDERBREACH means a reserve or senior obligation "
        "is funded below the level the Agreement requires; OVERBREACH would mean utilisation above a permitted cap "
        "or a payment made out of the Order of Priority."
    )

    # ---------------- 2. Deal summary ----------------
    _h(doc, "2. Deal Summary (from KB)")
    deal_rows = []
    for label, key in [("Borrower", "borrower"), ("Sponsor", "sponsor"), ("Lender", "lender"),
                       ("Authority", "authority"), ("Facility", "facility"),
                       ("Facility Amount", "facility_amount"), ("Project", "project"),
                       ("Concession Agreement Signed", "concession_agreement_signed")]:
        val = deal_info.get(key)
        if val:
            deal_rows.append([label, val])
    if extract.get("annuity_terms", {}).get("structure"):
        deal_rows.append(["Annuity structure", extract["annuity_terms"]["structure"]])
    if not deal_rows:
        deal_rows.append(["Note", "No deal-level fields were extracted from the source documents — verify manually"])
    _table(doc, ["Item", "Detail"], deal_rows, widths=[1.6, 5.4])

    # ---------------- 3. CATRA summary ----------------
    _h(doc, "3. CATRA Classification Framework (generated)")
    doc.add_paragraph(
        "Debit-side categories follow the Escrow Agreement Order of Priority (authoritative); the Sanction Letter "
        "clause-20 waterfall maps onto it step-for-step. Credit-side categories follow the permitted deposits in "
        "clause 2.3(A). Full detail, keyword rules and clause references are in CATRA_Master.xlsx; the transaction "
        "classifier now runs against this taxonomy (config/categories.yaml)."
    )
    _table(doc, ["Priority", "Code", "Category", "Agreement", "Sanction cl.20"],
           [[c["priority"], c["code"], c["name"], f"2.3(B)({c['item']})",
             ", ".join(str(s["step"]) for s in extract["sanction_waterfall"] if s["maps_to"] == c["code"]) or "—"]
            for c in extract["order_of_priority"]],
           widths=[0.7, 1.1, 2.6, 1.1, 1.2])
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.add_run("Waterfall divergences identified: ").font.bold = True
    for d in extract["waterfall_divergences"]:
        doc.add_paragraph(d, style="List Bullet")

    # ---------------- 4. TRA summary ----------------
    _h(doc, "4. TRA Analysis (from Sanction Note / CAM Projected P&L)")
    pnl = extract["projected_pnl"]
    bs = extract["projected_balance_sheet"]
    if not pnl.get("fye"):
        doc.add_paragraph("No projected P&L was extracted from the Sanction Note / CAM for this deal — "
                          "this section will populate once that data is added (see notes in section 6).")
    else:
        doc.add_paragraph(
            "Summary of the projected cashflow and reserve position by financial year. Full detail, the "
            "CATRA-mapped cashflow view and the semi-annual profile are in TRA_Analysis.xlsx."
        )
        from escrow_agent.profile_normalize import representative_years
        fy_show = representative_years(pnl["fye"], count=min(5, len(pnl["fye"])))
        idx = [pnl["fye"].index(f) for f in fy_show]
        _table(doc, ["Rs. crore"] + fy_show, [
            ["Total income"] + [pnl["income_total"][i] for i in idx],
            ["EBITDA"] + [pnl["ebitda"][i] for i in idx],
            ["Interest (facility)"] + [pnl["interest"][i] for i in idx],
            ["Principal (CMLTD)"] + [bs["cmltd"][i] for i in idx],
            ["PAT"] + [pnl["pat"][i] for i in idx],
            ["DSRA fund"] + [bs["dsra_fund"][i] for i in idx],
            ["WCR fund"] + [bs["working_capital_reserve"][i] for i in idx],
        ], widths=[1.8] + [1.04] * len(fy_show))

    # ---------------- 5. Clause cross-check ----------------
    _h(doc, "5. Cross-Check Against Escrow Agreement Clauses")
    rows = []
    # aggregate reserve checks by covenant: worst year
    for cvid, label in [("CV2", "DSRA = 2 quarters debt service (2.3(B)(j))"),
                        ("CV1", "WCR = 6 months interest at COD (2.3(B)(i))"),
                        ("CV3", "MMRA per base case (2.3(B)(k))")]:
        sub = [c for c in reserve_checks if c["covenant"].startswith(cvid)]
        bad = [c for c in sub if c["status"] in ("UNDERBREACH", "OVERBREACH")]
        marg = [c for c in sub if c["status"] == "MARGINAL"]
        if bad:
            worst = min(bad, key=lambda c: c["gap"])
            extra = f"; {len(marg)} further year(s) marginal" if marg else ""
            rows.append([label, worst["status"],
                         f"Worst year {worst['fy']}: required {worst['required']} vs provided {worst['provided']} (gap {worst['gap']}){extra}. {worst['note']}"])
        elif marg:
            worst = min(marg, key=lambda c: c["gap"])
            rows.append([label, "MARGINAL",
                         f"{len(marg)} year(s) within proxy tolerance (worst {worst['fy']}: gap {worst['gap']}). {worst['note']}"])
        else:
            rows.append([label, "COMPLIANT", "Projected fund meets requirement in all applicable years"])
    handled_ids = {"CV1", "CV2", "CV3"}   # covered above via reserve_checks (DSRA/WCR/MMRA)
    other_covenants = [cv for cv in extract.get("covenants", []) if cv.get("id") not in handled_ids]
    for cv in other_covenants:
        rows.append([f"{cv.get('rule', 'Covenant')} ({cv.get('source', 'source not extracted')})",
                     "PENDING ACTUALS",
                     "Requires the escrow account bank statement to verify against actual transactions/dates."])
    if not other_covenants:
        rows.append(["Other covenants beyond reserve sizing", "NONE EXTRACTED",
                     "No further covenants were extracted from the source documents beyond DSRA/WCR/MMRA sizing — "
                     "review the Escrow Agreement and Sanction Letter manually for payment-routing, distribution, "
                     "or prepayment conditions and add them to the deal profile."])
    for d in extract.get("waterfall_divergences", []):
        if d != "None identified.":
            rows.append(["Waterfall divergence", "ATTENTION", d])
    t = _table(doc, ["Clause / Covenant", "Status", "Finding"], rows, widths=[2.3, 1.2, 3.5])
    _status_color(t, 1)

    # ---------------- 6. Detailed reserve table ----------------
    _h(doc, "6. Reserve Adequacy Detail (projected basis)")
    t2 = _table(doc, ["Covenant", "FY", "Required", "Provided", "Gap", "Status"],
                [[c["covenant"], c["fy"], c["required"], c["provided"], c["gap"], c["status"]]
                 for c in reserve_checks],
                widths=[2.4, 0.7, 1.0, 1.0, 0.9, 1.0])
    _status_color(t2, 5)

    # ---------------- 7. Next inputs ----------------
    _h(doc, "7. Inputs Required to Move from Projected to Actuals Basis")
    generic_items = [
        "Escrow account bank statement (xlsx/csv/pdf) — enables classification against the generated CATRA, "
        "waterfall-order validation, and actual-vs-TRA variance",
        "COD date / annuity or revenue payment schedule as invoiced — to pin the TRA profile to actual dates",
    ]
    if other_covenants:
        generic_items.append("Confirmation from the lender on the covenants listed as PENDING ACTUALS above")
    for d in extract.get("waterfall_divergences", []):
        if d != "None identified.":
            generic_items.append(f"Resolution of: {d}")
    if not pnl.get("fye"):
        generic_items.append("Projected P&L / Balance Sheet from the Sanction Note or CAM — none was extracted; "
                             "TRA and reserve-adequacy checks are placeholders until this is added")
    for item in generic_items:
        doc.add_paragraph(item, style="List Bullet")

    doc.save(out_path)
    return out_path, verdict
