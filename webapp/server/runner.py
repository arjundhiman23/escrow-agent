"""Background run executor: pulls inputs from storage, runs the pipeline in a temp
dir, streams stage + log updates into meta.json, uploads the three reports.
Runs are now nested under a deal (deals/{deal_id}/runs/{run_id}/...)."""
import os, sys, tempfile, traceback
from datetime import datetime, timezone

BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
sys.path.insert(0, BASE)

from escrow_agent.bank_ingest import read_quarter, classify_all, quarter_summary, CREDIT_CATS, DEBIT_CATS  # noqa
from escrow_agent.atsl_reports import build_catra, build_tra, build_final_analysis_actuals, QUARTERS  # noqa
from escrow_agent.ai_usage import UsageTracker  # noqa
from webapp.server.storage import read_run, write_run, read_deal  # noqa

AI_MODEL_CLASSIFICATION = os.environ.get("AI_MODEL_CLASSIFICATION", "claude-haiku-4-5-20251001")

STAGES = ["INGEST", "CLASSIFY", "VALIDATE", "CATRA", "TRA", "FINAL"]
CR = 1e7


def _now():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _log(st, deal_id, run_id, meta, line):
    meta["log"].append(f"[{_now()}] {line}")
    write_run(st, deal_id, run_id, meta)


def _stage(st, deal_id, run_id, meta, stage):
    meta["stage"] = stage
    meta["stages_done"] = STAGES[:STAGES.index(stage)]
    _log(st, deal_id, run_id, meta, f"stage -> {stage}")


def execute_run(st, deal_id, run_id):
    meta = read_run(st, deal_id, run_id)
    meta.update(status="running", started_at=_now(), log=meta.get("log", []))
    write_run(st, deal_id, run_id, meta)
    try:
        deal = read_deal(st, deal_id)
        deal_name = deal.get("name", "Borrower not specified")
        deal_covenants = (deal.get("profile") or {}).get("covenants", [])

        with tempfile.TemporaryDirectory() as tmp:
            ind, outd = os.path.join(tmp, "in"), os.path.join(tmp, "out")
            os.makedirs(ind); os.makedirs(outd)
            for f in meta["inputs"]["statements"] + [meta["inputs"]["catra_template"], meta["inputs"]["tra_template"]]:
                with open(os.path.join(ind, f), "wb") as fh:
                    fh.write(st.get_bytes(f"deals/{deal_id}/runs/{run_id}/inputs/{f}"))

            _stage(st, deal_id, run_id, meta, "INGEST")
            txns = []
            for f in sorted(meta["inputs"]["statements"]):
                q = next((q for q in QUARTERS if q in f.upper()), None)
                if not q:
                    raise ValueError(f"cannot infer quarter from file name: {f} (name must contain Q1..Q4)")
                rows = read_quarter(os.path.join(ind, f), q, meta["account"])
                if not rows:
                    raise ValueError(f"no transactions for account {meta['account']} in {f}")
                txns += rows
                _log(st, deal_id, run_id, meta, f"{f} -> {q}: {len(rows)} main-account transactions")

            _stage(st, deal_id, run_id, meta, "CLASSIFY")
            tracker = UsageTracker()
            ai_assist = bool(os.environ.get("ANTHROPIC_API_KEY")) and meta.get("ai_assist", True)
            classify_all(txns, ai_assist=ai_assist, ai_model=AI_MODEL_CLASSIFICATION, tracker=tracker)
            summary = quarter_summary(txns)
            n_review = sum(1 for t in txns if t.review or t.conflict)
            n_ai_suggested = sum(1 for t in txns if t.ai_suggestion)
            _log(st, deal_id, run_id, meta, f"classified {len(txns)} transactions; {n_review} flagged for review")
            if ai_assist:
                u = tracker.summary()
                _log(st, deal_id, run_id, meta,
                    f"AI-assist ({AI_MODEL_CLASSIFICATION}): {u['calls']} call(s), "
                    f"{u['input_tokens']} in / {u['output_tokens']} out tokens, "
                    f"${u['cost_usd']:.4f}, {n_ai_suggested} suggestion(s) for analyst review")
            else:
                u = {"input_tokens": 0, "output_tokens": 0, "cost_usd": 0.0, "calls": 0, "by_purpose": {}}

            _stage(st, deal_id, run_id, meta, "VALIDATE")
            qs = [q for q in QUARTERS if q in summary]
            for q in qs:
                s = summary[q]
                if abs(s["opening"] + s["total_credit"] - s["total_debit"] - s["closing"]) >= 0.01:
                    raise AssertionError(f"{q} balance break")
            for i in range(len(qs) - 1):
                if abs(summary[qs[i]]["closing"] - summary[qs[i + 1]]["opening"]) >= 0.01:
                    _log(st, deal_id, run_id, meta, f"WARNING: closing {qs[i]} != opening {qs[i+1]} (gap across quarters)")
            _log(st, deal_id, run_id, meta, "balance integrity OK")

            _stage(st, deal_id, run_id, meta, "CATRA")
            catra_path, _ = build_catra(os.path.join(ind, meta["inputs"]["catra_template"]), summary,
                                        os.path.join(outd, "CATRA_ANALYSIS_ATSL.xlsx"),
                                        deal_name=deal_name, account=meta["account"], fy=meta.get("fy", ""))
            _stage(st, deal_id, run_id, meta, "TRA")
            tra_path, _ = build_tra(os.path.join(ind, meta["inputs"]["tra_template"]), summary,
                                    os.path.join(outd, "TRA_Analysis_ATSL.xlsx"),
                                    extract=deal.get("profile"), fy=meta.get("fy", ""))
            _stage(st, deal_id, run_id, meta, "FINAL")
            recon = {"n_txns": len(txns), "n_cells": 12 * len(qs),
                     "accuracy": "totals reconcile to statement balances",
                     "n_review": n_review}
            fa_path, verdict = build_final_analysis_actuals(
                summary, txns, recon, os.path.join(outd, "Final_Analysis.docx"),
                deal_name=deal_name, account=meta["account"], fy=meta.get("fy", ""),
                deal_covenants=deal_covenants)

            outputs = []
            for p in (catra_path, tra_path, fa_path):
                name = os.path.basename(p)
                with open(p, "rb") as fh:
                    st.put_bytes(f"deals/{deal_id}/runs/{run_id}/outputs/{name}", fh.read())
                outputs.append(name)

            meta.update(
                status="completed", stage="DONE", stages_done=STAGES, verdict=verdict,
                outputs=outputs, finished_at=_now(),
                ai_usage=u,
                stats={
                    "transactions": len(txns), "review_items": n_review,
                    "ai_suggestions": n_ai_suggested, "quarters": qs,
                    "total_inflow_cr": round(sum(summary[q]["total_credit"] for q in qs) / CR, 2),
                    "total_outflow_cr": round(sum(summary[q]["total_debit"] for q in qs) / CR, 2),
                    "opening_cr": round(summary[qs[0]]["opening"] / CR, 2),
                    "closing_cr": round(summary[qs[-1]]["closing"] / CR, 2),
                    "quarter_table": {q: {
                        "inflow_cr": round(summary[q]["total_credit"] / CR, 2),
                        "outflow_cr": round(summary[q]["total_debit"] / CR, 2),
                        "closing_cr": round(summary[q]["closing"] / CR, 2)} for q in qs},
                },
            )
            _log(st, deal_id, run_id, meta, f"run completed — verdict {verdict}")
    except Exception as e:
        meta.update(status="failed", error=f"{type(e).__name__}: {e}", finished_at=_now())
        meta["log"].append(f"[{_now()}] FAILED: {e}")
        meta["log"].append(traceback.format_exc(limit=3))
        write_run(st, deal_id, run_id, meta)
