#!/usr/bin/env python3
"""Run the KB -> CATRA -> TRA -> Final Analysis pipeline.

Usage: python run_kb_analysis.py [--extract kb/structured/deal_extract.yaml] [--output output]
"""
import argparse, os, sys, yaml

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from escrow_agent.catra_builder import load_extract, build_categories_yaml, build_catra_xlsx
from escrow_agent.tra_builder import build_tra_xlsx
from escrow_agent.final_analysis import build_final_analysis


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--extract", default="kb/structured/deal_extract.yaml")
    ap.add_argument("--output", default="output")
    args = ap.parse_args()
    os.makedirs(args.output, exist_ok=True)

    extract = load_extract(args.extract)

    n = build_categories_yaml(extract, "config/categories.yaml")
    print(f"[1/4] categories.yaml regenerated from KB ({n} categories)")

    catra_path = build_catra_xlsx(extract, os.path.join(args.output, "CATRA_Master.xlsx"))
    print(f"[2/4] CATRA master -> {catra_path}")

    tra_path, checks = build_tra_xlsx(extract, os.path.join(args.output, "TRA_Analysis.xlsx"))
    fails = sum(1 for c in checks if c["status"] in ("UNDERBREACH", "OVERBREACH"))
    print(f"[3/4] TRA analysis -> {tra_path} ({len(checks)} reserve checks, {fails} failing)")

    fa_path, verdict = build_final_analysis(extract, checks, os.path.join(args.output, "Final_Analysis.docx"))
    print(f"[4/4] Final analysis -> {fa_path}")
    print(f"\nOVERALL VERDICT (projected basis): {verdict}")


if __name__ == "__main__":
    main()
