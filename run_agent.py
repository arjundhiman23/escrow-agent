#!/usr/bin/env python3
"""Escrow Account Transaction Analyst Agent — CLI entrypoint.

Usage:
  python run_agent.py \
      --transactions sample_data/sample_transactions.xlsx \
      --sanction-extract sample_data/sample_sanction_extract.json \
      --output output/

  # or, extract structure from a raw sanction document via AI (needs ANTHROPIC_API_KEY):
  python run_agent.py --transactions txns.xlsx --sanction-doc sanction_letter.pdf --output output/
"""
from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from escrow_agent import aggregation, classification, ingestion, sanction as sanction_mod, waterfall  # noqa: E402
from escrow_agent import reporting_excel, reporting_word  # noqa: E402
from escrow_agent.audit_log import AuditLog  # noqa: E402
from escrow_agent.models import load_settings, load_taxonomy  # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser(description="Escrow Account Transaction Analyst Agent")
    ap.add_argument("--transactions", required=True,
                    help="Transaction extract for the MAIN escrow account (.xlsx/.csv/.pdf)")
    grp = ap.add_mutually_exclusive_group(required=True)
    grp.add_argument("--sanction-extract", help="Reviewed sanction extract JSON")
    grp.add_argument("--sanction-doc", help="Raw Sanction Letter/Note (.pdf/.docx) — AI extraction")
    ap.add_argument("--config", default=str(Path(__file__).parent / "config"),
                    help="Directory containing categories.yaml and settings.yaml")
    ap.add_argument("--output", default="output", help="Output directory for reports")
    args = ap.parse_args()

    t0 = time.time()
    cfg_dir = Path(args.config)
    out_dir = Path(args.output)
    out_dir.mkdir(parents=True, exist_ok=True)

    audit = AuditLog(out_dir)
    audit.log("RUN_START", f"transactions={args.transactions}, "
                           f"sanction={args.sanction_extract or args.sanction_doc}, config={cfg_dir}")

    taxonomy = load_taxonomy(cfg_dir / "categories.yaml")
    settings = load_settings(cfg_dir / "settings.yaml")
    audit.log("CONFIG", f"Taxonomy loaded: {len(taxonomy.inflows)} inflow / "
                        f"{len(taxonomy.outflows)} outflow categories.")

    # 1. Ingestion & preprocessing (§6.1)
    txns, warnings = ingestion.ingest(args.transactions, taxonomy, settings, audit)
    if not txns:
        print("No transactions parsed — aborting.", file=sys.stderr)
        return 2

    # 2. Classification (§6.2)
    classification.classify(txns, taxonomy, settings, audit)

    # 3. Sanction document analysis (§6.4)
    if args.sanction_extract:
        sanction = sanction_mod.load_sanction_extract(args.sanction_extract, taxonomy, audit)
    else:
        sanction = sanction_mod.extract_from_document(args.sanction_doc, taxonomy, settings, audit)

    # 4. Waterfall & condition validation (§6.4)
    wf_exceptions = waterfall.validate(txns, sanction, settings, audit)
    manual_conds = waterfall.manual_conditions(sanction)

    # 5. Quarter-wise aggregation (§6.3)
    agg = aggregation.aggregate(txns, taxonomy, audit)

    # 6. Variance analysis (§6.5)
    var = aggregation.variance_analysis(agg, sanction, taxonomy, settings, audit)

    # 7. Report generation (§6.6)
    xlsx_path = reporting_excel.generate(
        out_dir / "Escrow_Analysis_Workbook.xlsx", txns, taxonomy, agg, var,
        wf_exceptions, manual_conds, sanction, audit)
    docx_path = reporting_word.generate(
        out_dir / "Escrow_Analysis_Report.docx", txns, taxonomy, agg, var,
        wf_exceptions, manual_conds, sanction, warnings, audit)

    elapsed = time.time() - t0
    audit.log("RUN_COMPLETE", f"Elapsed {elapsed:.1f}s (BRD §11 target: < 300s).")

    print(f"Done in {elapsed:.1f}s")
    print(f"  Excel workbook : {xlsx_path}")
    print(f"  Word report    : {docx_path}")
    print(f"  Audit log      : {audit.path}")
    review = sum(t.needs_review or t.is_escalated_conflict for t in txns)
    if review:
        print(f"  NOTE: {review} transactions await analyst review (see 'Review Queue' sheet).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
