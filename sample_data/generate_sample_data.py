#!/usr/bin/env python3
"""Generates SYNTHETIC demo data: a full-FY transaction extract and a matching
sanction extract JSON. Purely illustrative — replace with real Axis extracts,
the TRA/CATRA framework file, and reviewed sanction extracts.

Deliberately embedded edge cases (to exercise every BRD control):
  - one exact duplicate row                       -> duplicate flag
  - one blank narration                           -> analyst review queue
  - one Cr narrated like Debt Servicing           -> escalated conflict
  - one artificially wrong running balance        -> balance break
  - a Q2 month with surplus paid, no debt service -> DEBIT_OUT_OF_SEQUENCE + condition breach
  - an 'insurance claim' credit (unmapped)        -> review queue (rules-only mode)
  - INTERCO transfers exceeding the Q3 cap        -> CONDITION_BREACH
"""
import json
import random
from datetime import date
from pathlib import Path

from openpyxl import Workbook

HERE = Path(__file__).parent
random.seed(42)

rows = []  # (date, drcr, amount, narration)


def add(d, m, y, t, amt, narr):
    rows.append((date(y, m, d), t, float(amt), narr))


# ---- FY 2025-26 --------------------------------------------------------------
for m, y in [(4, 2025), (5, 2025), (6, 2025), (7, 2025), (8, 2025), (9, 2025),
             (10, 2025), (11, 2025), (12, 2025), (1, 2026), (2, 2026), (3, 2026)]:
    add(5, m, y, "Cr", 9000000 + random.randint(-800000, 800000),
        f"NEFT RECEIPT FROM CUSTOMER TARIFF COLLECTION {m:02d}{y}")
    add(12, m, y, "Dr", 1200000 + random.randint(-50000, 50000),
        "VENDOR PAYMENT O&M CONTRACTOR MONTHLY BILL")
    add(18, m, y, "Dr", 450000, "GST PAYMENT CBIC PORTAL")
    add(20, m, y, "Dr", 150000, "TDS PAYMENT INCOME TAX")
    if m != 8:  # August: debt service deliberately skipped
        add(25, m, y, "Dr", 4000000, "INTEREST PAYMENT TERM LOAN A/C 4021 TO LENDER")

# quarterly principal repayments
for d, m, y in [(28, 6, 2025), (28, 9, 2025), (28, 12, 2025), (28, 3, 2026)]:
    add(d, m, y, "Dr", 2500000, "PRINCIPAL REPAYMENT TERM LOAN INSTALMENT")

# FD activity
add(10, 5, 2025, "Dr", 5000000, "FD BOOKING TERM DEPOSIT 91 DAYS")
add(11, 8, 2025, "Cr", 5093000, "FD MATURITY PROCEEDS WITH INTEREST CREDIT")
add(15, 11, 2025, "Dr", 6000000, "FD PLACED TERM DEPOSIT 180 DAYS")

# equity & term loan proceeds
add(8, 4, 2025, "Cr", 20000000, "EQUITY INFUSION PROMOTER CONTRIBUTION TRANCHE 3")
add(14, 7, 2025, "Cr", 15000000, "TERM LOAN DISBURSEMENT TRANCHE 5 FACILITY B")

# reserves
add(30, 6, 2025, "Dr", 3000000, "TRANSFER TO DSRA DEBT SERVICE RESERVE")

# surplus distribution: Aug (no debt service that month -> violations) + Mar (clean)
add(27, 8, 2025, "Dr", 2000000, "SURPLUS DISTRIBUTION TO BORROWER AS PER WATERFALL")
add(30, 3, 2026, "Dr", 1500000, "SURPLUS DISTRIBUTION TO BORROWER Q4")

# inter-company transfers; Q3 total 5.8mn breaches the 5mn cap
add(9, 10, 2025, "Dr", 3000000, "INTER COMPANY TRANSFER TO GROUP COMPANY SUBSIDIARY")
add(19, 12, 2025, "Dr", 2800000, "IC TRANSFER GROUP COMPANY WORKING CAPITAL SUPPORT")

# internal own-account transfers (flag separately)
add(6, 9, 2025, "Dr", 1000000, "INTERNAL TRANSFER TO OWN ACCOUNT CURRENT A/C SWEEP")
add(7, 9, 2025, "Cr", 1000000, "TRANSFER TO OWN ACCOUNT REVERSAL SWEEP RETURN")

# --- edge cases ---------------------------------------------------------------
add(16, 5, 2025, "Cr", 750000, "")                                        # blank narration -> review
add(21, 7, 2025, "Cr", 300000, "INTEREST PAYMENT TERM LOAN")              # Cr with Dr-style narration -> conflict
add(23, 10, 2025, "Cr", 1250000, "INSURANCE CLAIM SETTLEMENT RECEIVED")   # unmapped -> review
add(18, 1, 2026, "Dr", 450000, "GST PAYMENT CBIC PORTAL")                 # exact duplicate of Jan GST row

rows.sort(key=lambda r: r[0])

# ---- write xlsx with running balance ------------------------------------------
wb = Workbook()
ws = wb.active
ws.title = "Statement"
ws.append(["Axis Bank – Escrow Account Statement (SYNTHETIC DEMO DATA)"])
ws.append([])
ws.append(["Txn Date", "Dr/Cr", "Amount (INR)", "Narration/Remarks", "Balance"])
bal = 12000000.0
BREAK_ROW = 30  # inject a wrong balance at this data row -> balance break flag
for i, (d, t, amt, narr) in enumerate(rows):
    bal = bal + amt if t == "Cr" else bal - amt
    reported = bal + (500000 if i == BREAK_ROW else 0)
    ws.append([d.strftime("%d/%m/%Y"), t, round(amt, 2), narr, round(reported, 2)])
wb.save(HERE / "sample_transactions.xlsx")

# ---- sanction extract JSON -----------------------------------------------------
sanction = {
    "deal_name": "DEMO — Solar SPV Escrow (Synthetic)",
    "waterfall": [
        {"priority": 1, "category_code": "STAT_PAY", "description": "Statutory dues and taxes"},
        {"priority": 2, "category_code": "OM_EXP", "description": "O&M and project construction payments"},
        {"priority": 3, "category_code": "DEBT_SERV", "description": "Interest and principal to lenders"},
        {"priority": 4, "category_code": "RESERVE", "description": "DSRA / reserve top-ups"},
        {"priority": 5, "category_code": "PERM_INV", "description": "Permitted investments (FD/TD)"},
        {"priority": 6, "category_code": "INTERCO", "description": "Inter-company transfers (capped)"},
        {"priority": 7, "category_code": "CO_TRF", "description": "Own-account transfers"},
        {"priority": 8, "category_code": "SURPLUS_DIST", "description": "Surplus distribution to borrower"},
    ],
    "permitted_inflow_sources": ["OTH_REV", "PROC_TL", "FD_REDEMPTION"],  # equity must route via sponsor account -> direct equity credit is unpermitted
    "conditions": [
        {"id": "C1", "type": "surplus_requires_prior",
         "params": {"required_code": "DEBT_SERV", "target_code": "SURPLUS_DIST"},
         "description": "Surplus may be distributed only after debt servicing in the same monthly cycle."},
        {"id": "C2", "type": "category_cap_per_quarter",
         "params": {"category_code": "INTERCO", "cap": 5000000},
         "description": "Inter-company transfers capped at INR 50 lakh per quarter."},
        {"id": "C3", "type": "manual",
         "description": "Borrower to submit insurance policy renewals to the trustee annually."},
    ],
    "projections": {
        "Q1": {"inflows": {"OTH_REV": 27000000, "FD_REDEMPTION": 0},
               "outflows": {"DEBT_SERV": 14500000, "STAT_PAY": 1800000, "OM_EXP": 3600000,
                            "RESERVE": 3000000, "PERM_INV": 5000000}},
        "Q2": {"inflows": {"OTH_REV": 27000000, "PROC_TL": 15000000, "FD_REDEMPTION": 5100000},
               "outflows": {"DEBT_SERV": 14500000, "STAT_PAY": 1800000, "OM_EXP": 3600000,
                            "RESERVE": 2000000, "SURPLUS_DIST": 2000000}},
        "Q3": {"inflows": {"OTH_REV": 27000000},
               "outflows": {"DEBT_SERV": 14500000, "STAT_PAY": 1800000, "OM_EXP": 3600000,
                            "PERM_INV": 6000000, "INTERCO": 5000000}},
        "Q4": {"inflows": {"OTH_REV": 27000000},
               "outflows": {"DEBT_SERV": 14500000, "STAT_PAY": 1800000, "OM_EXP": 3600000,
                            "SURPLUS_DIST": 1500000}},
    },
}
(HERE / "sample_sanction_extract.json").write_text(json.dumps(sanction, indent=2), encoding="utf-8")
print(f"Wrote {len(rows)} synthetic transactions + sanction extract to {HERE}")
