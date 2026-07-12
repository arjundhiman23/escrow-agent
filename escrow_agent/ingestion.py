"""Data ingestion & preprocessing (BRD §6.1).

Handles Excel (.xlsx), CSV, and PDF transaction extracts for the MAIN escrow
account. Performs:
  - adaptive column mapping (field names vary across Axis extracts — Risk R-04)
  - date standardization to DD-MM-YYYY
  - duplicate detection
  - opening/closing balance consistency validation
  - internal company transfer flagging
"""
from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path

import pandas as pd

from .models import Taxonomy, Transaction

# ---------------------------------------------------------------------------
# Adaptive column mapping (Risk R-04: format inconsistencies across extracts)
# ---------------------------------------------------------------------------
COLUMN_ALIASES = {
    "date":      ["transaction date", "txn date", "tran date", "value date", "date", "posting date"],
    "type":      ["type", "dr/cr", "cr/dr", "debit/credit", "txn type", "transaction type", "d/c"],
    "amount":    ["amount", "transaction amount", "txn amount", "amount (inr)", "amt"],
    "debit":     ["debit", "withdrawal", "withdrawal amt", "debit amount", "withdrawal amount", "dr amount"],
    "credit":    ["credit", "deposit", "deposit amt", "credit amount", "deposit amount", "cr amount"],
    "narration": ["narration", "remarks", "description", "particulars", "transaction remarks", "narration/remarks"],
    "balance":   ["balance", "running balance", "closing balance", "available balance", "balance (inr)"],
}

DATE_FORMATS = ["%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%b-%Y", "%d %b %Y",
                "%d.%m.%Y", "%m/%d/%Y", "%d-%b-%y", "%d/%m/%y"]


def _norm(s) -> str:
    return re.sub(r"\s+", " ", str(s)).strip().lower()


def _map_columns(columns) -> dict[str, str]:
    """Map logical field -> actual column name found in the file."""
    normed = {_norm(c): c for c in columns}
    mapping = {}
    for field, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in normed:
                mapping[field] = normed[alias]
                break
    return mapping


def _parse_date(value) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    s = str(value).strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    # last resort: pandas parser (dayfirst per Indian convention)
    parsed = pd.to_datetime(s, dayfirst=True, errors="coerce")
    if pd.isna(parsed):
        raise ValueError(f"Unparseable transaction date: {value!r}")
    return parsed.date()


def _parse_amount(value) -> float:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    s = str(value).replace(",", "").replace("₹", "").strip()
    if s in ("", "-", "nan", "None"):
        return 0.0
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    amt = float(s)
    return -amt if neg else amt


# ---------------------------------------------------------------------------
# Readers
# ---------------------------------------------------------------------------
def _read_dataframe(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix in (".xlsx", ".xlsm", ".xls"):
        # Header row may not be row 0 in bank extracts — probe first 15 rows.
        raw = pd.read_excel(path, header=None, nrows=15)
        header_row = _find_header_row(raw)
        return pd.read_excel(path, header=header_row)
    if suffix in (".csv", ".tsv"):
        sep = "\t" if suffix == ".tsv" else ","
        raw = pd.read_csv(path, header=None, nrows=15, sep=sep, dtype=str)
        header_row = _find_header_row(raw)
        return pd.read_csv(path, header=header_row, sep=sep)
    if suffix == ".pdf":
        return _read_pdf_tables(path)
    raise ValueError(f"Unsupported transaction file format: {suffix}")


def _find_header_row(raw: pd.DataFrame) -> int:
    """Locate the header row: the first row containing a date-alias AND a
    narration-alias cell (bank extracts often carry title/address rows first)."""
    date_aliases = set(COLUMN_ALIASES["date"])
    narr_aliases = set(COLUMN_ALIASES["narration"])
    for i in range(len(raw)):
        cells = {_norm(v) for v in raw.iloc[i].tolist()}
        if cells & date_aliases and cells & narr_aliases:
            return i
    return 0


def _read_pdf_tables(path: Path) -> pd.DataFrame:
    """Best-effort table extraction from PDF statements (BRD §6.1).
    Scanned/image PDFs (Risk R-02) raise a clear error directing to OCR."""
    import pdfplumber

    frames = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables() or []:
                if not table or len(table) < 2:
                    continue
                df = pd.DataFrame(table[1:], columns=table[0])
                frames.append(df)
    if not frames:
        raise ValueError(
            f"No extractable tables found in {path.name}. If this is a scanned "
            "PDF, apply OCR pre-processing first (Risk R-02) or supply the "
            "Excel/CSV extract."
        )
    return pd.concat(frames, ignore_index=True)


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------
def ingest(path: str | Path, taxonomy: Taxonomy, settings: dict,
           audit=None) -> tuple[list[Transaction], list[str]]:
    """Ingest a transaction file. Returns (transactions, ingestion_warnings)."""
    path = Path(path)
    warnings: list[str] = []
    df = _read_dataframe(path)
    df = df.dropna(how="all")
    mapping = _map_columns(df.columns)

    missing_core = [f for f in ("date", "narration") if f not in mapping]
    has_amount = "amount" in mapping and "type" in mapping
    has_drcr_cols = "debit" in mapping and "credit" in mapping
    if missing_core or not (has_amount or has_drcr_cols):
        raise ValueError(
            f"Could not map required fields in {path.name}. Found columns: "
            f"{list(df.columns)}. Need date + narration + (type&amount, or "
            f"separate debit/credit columns). Update COLUMN_ALIASES if the "
            f"bank extract uses new headers."
        )

    txns: list[Transaction] = []
    for i, row in df.iterrows():
        narr_val = row.get(mapping["narration"], "")
        narration = "" if pd.isna(narr_val) else str(narr_val).strip()
        date_val = row.get(mapping["date"])
        if pd.isna(date_val) or str(date_val).strip() == "":
            continue  # trailer/summary rows
        try:
            txn_date = _parse_date(date_val)
        except ValueError:
            warnings.append(f"Row {i + 2}: unparseable date {date_val!r} — row skipped.")
            continue

        if has_drcr_cols and not has_amount:
            dr = _parse_amount(row.get(mapping["debit"]))
            cr = _parse_amount(row.get(mapping["credit"]))
            if dr > 0 and cr > 0:
                warnings.append(f"Row {i + 2}: both debit and credit populated — row skipped, verify manually.")
                continue
            if dr == 0 and cr == 0:
                continue
            txn_type, amount = ("Dr", dr) if dr > 0 else ("Cr", cr)
        else:
            raw_type = _norm(row.get(mapping["type"], ""))
            if raw_type.startswith(("dr", "d", "debit", "withdraw")):
                txn_type = "Dr"
            elif raw_type.startswith(("cr", "c", "credit", "deposit")):
                txn_type = "Cr"
            else:
                warnings.append(f"Row {i + 2}: unrecognised Dr/Cr type {raw_type!r} — row skipped.")
                continue
            amount = abs(_parse_amount(row.get(mapping["amount"])))

        balance = None
        if "balance" in mapping and not pd.isna(row.get(mapping["balance"])):
            try:
                balance = _parse_amount(row.get(mapping["balance"]))
            except ValueError:
                pass

        t = Transaction(row_no=int(i) + 2, txn_date=txn_date, txn_type=txn_type,
                        amount=round(amount, 2), narration=narration, balance=balance)
        t.is_internal_transfer = any(k in narration.lower()
                                     for k in taxonomy.internal_transfer_keywords)
        txns.append(t)

    txns.sort(key=lambda t: (t.txn_date, t.row_no))
    _flag_duplicates(txns)
    warnings.extend(_validate_balances(txns, settings))
    _assign_quarters(txns, settings)

    if audit:
        audit.log("INGESTION", f"{path.name}: {len(txns)} transactions parsed, "
                               f"{sum(t.is_duplicate for t in txns)} duplicates flagged, "
                               f"{sum(t.balance_break for t in txns)} balance breaks, "
                               f"{sum(t.is_internal_transfer for t in txns)} internal transfers, "
                               f"{len(warnings)} warnings.")
    return txns, warnings


def _flag_duplicates(txns: list[Transaction]) -> None:
    seen: dict[tuple, int] = {}
    for t in txns:
        key = (t.txn_date, t.txn_type, round(t.amount, 2), t.narration.lower())
        if key in seen:
            t.is_duplicate = True
        else:
            seen[key] = t.row_no


def _validate_balances(txns: list[Transaction], settings: dict) -> list[str]:
    """Running-balance consistency: prev_balance +Cr / -Dr must equal the
    reported balance within tolerance (BRD §6.1, §11)."""
    tol_pct = float(settings.get("balance_tolerance_pct", 0.1)) / 100.0
    warnings = []
    prev = None
    for t in txns:
        if t.balance is None:
            continue
        if prev is not None:
            expected = prev + t.amount if t.txn_type == "Cr" else prev - t.amount
            tolerance = max(abs(expected) * tol_pct, 1.0)
            if abs(expected - t.balance) > tolerance:
                t.balance_break = True
                warnings.append(
                    f"Balance break at row {t.row_no} ({t.txn_date:%d-%m-%Y}): "
                    f"expected {expected:,.2f}, reported {t.balance:,.2f}."
                )
        prev = t.balance
    return warnings


def _assign_quarters(txns: list[Transaction], settings: dict) -> None:
    month_to_q = {m: q for q, months in settings["quarters"].items() for m in months}
    for t in txns:
        t.quarter = month_to_q.get(t.txn_date.month, "?")
