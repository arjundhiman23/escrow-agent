"""Escrow Analyst Console — FastAPI backend, organized by Deal.

Deal lifecycle: create -> upload Sanction Letter/Note + Escrow Agreement ->
AI extraction (draft profile) -> analyst reviews/edits -> confirm -> deal is
"ready". Quarterly runs (statements + bank templates -> CATRA/TRA/Final
Analysis) live under a deal once it exists (profile is optional but improves
the actuals Final Analysis with deal-specific covenant checks).
"""
import io, os, re, threading, uuid, traceback
from datetime import datetime, timezone
from typing import Optional

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse, FileResponse
from fastapi.staticfiles import StaticFiles

from webapp.server.storage import (
    get_storage, read_deal, write_deal, list_deals,
    get_deal_template_meta, set_deal_default_template, get_deal_default_template,
    read_run, write_run, list_runs,
)
from webapp.server.runner import execute_run, STAGES
from escrow_agent.profile_normalize import normalize_profile

app = FastAPI(title="Escrow Analyst Console")
ST = get_storage()
STATIC = os.path.join(os.path.dirname(__file__), "..", "static")
SAFE = re.compile(r"[^A-Za-z0-9._ -]")


def _now():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _safe(name):
    return SAFE.sub("_", os.path.basename(name or "file"))


# ============================================================== DEALS
@app.post("/api/deals")
def create_deal(name: str = Form(...), account: str = Form("")):
    deal_id = datetime.now().strftime("%Y%m%d-%H%M%S") + "-" + uuid.uuid4().hex[:6]
    meta = {
        "deal_id": deal_id, "name": name.strip(), "account": account.strip(),
        "status": "new",   # new -> extracting -> review -> ready
        "created_at": _now(), "documents": {}, "profile": None, "extraction_log": [],
    }
    write_deal(ST, deal_id, meta)
    return {"deal_id": deal_id}


@app.get("/api/deals")
def get_deals():
    deals = []
    for did in list_deals(ST):
        try:
            m = read_deal(ST, did)
        except Exception:
            continue
        n_runs = len(list_runs(ST, did))
        deals.append({**{k: m.get(k) for k in
                        ("deal_id", "name", "account", "status", "created_at", "documents", "ai_usage")},
                      "n_runs": n_runs})
    deals.sort(key=lambda d: d.get("created_at") or "", reverse=True)
    return {"deals": deals}


@app.get("/api/deals/{deal_id}")
def get_deal(deal_id: str):
    try:
        return read_deal(ST, _safe(deal_id))
    except Exception:
        raise HTTPException(404, "deal not found")


@app.delete("/api/deals/{deal_id}")
def delete_deal(deal_id: str):
    ST.delete_prefix(f"deals/{_safe(deal_id)}")
    return {"ok": True}


# ============================================================== DEAL DOCUMENTS + EXTRACTION
DOC_KINDS = ("escrow_agreement", "sanction_letter", "sanction_note")


@app.post("/api/deals/{deal_id}/documents/{kind}/extract")
async def upload_and_extract_one(deal_id: str, kind: str, file: UploadFile = File(...)):
    """Upload and extract exactly one document (escrow_agreement | sanction_letter |
    sanction_note). Independent of the other two — each can be uploaded and
    extracted separately, in any order, and its result is merged into the deal's
    existing profile without disturbing what other documents already contributed.
    Saving each document's result as soon as it individually completes also means
    a restart/redeploy mid-extraction only costs the one document in flight, not
    the whole batch."""
    deal_id, kind = _safe(deal_id), _safe(kind)
    if kind not in DOC_KINDS:
        raise HTTPException(400, f"kind must be one of {DOC_KINDS}")
    try:
        meta = read_deal(ST, deal_id)
    except Exception:
        raise HTTPException(404, "deal not found")

    name = _safe(file.filename)
    data = await file.read()
    ST.put_bytes(f"deals/{deal_id}/documents/{kind}.pdf", data)

    documents = dict(meta.get("documents", {}))
    documents[kind] = name
    meta["documents"] = documents
    doc_status = dict(meta.get("document_status", {}))
    doc_status[kind] = "extracting"
    meta["document_status"] = doc_status
    if meta["status"] == "new":
        meta["status"] = "extracting"
    write_deal(ST, deal_id, meta)

    threading.Thread(target=_run_single_extraction, args=(deal_id, kind), daemon=True).start()
    return {"ok": True, "status": "extracting"}


@app.get("/api/deals/{deal_id}/documents/{kind}/status")
def document_extract_status(deal_id: str, kind: str):
    """Lightweight poll target for a single document's extraction widget —
    lets the frontend refresh just that one card instead of the whole deal."""
    deal_id, kind = _safe(deal_id), _safe(kind)
    try:
        meta = read_deal(ST, deal_id)
    except Exception:
        raise HTTPException(404, "deal not found")
    return {
        "status": (meta.get("document_status") or {}).get(kind, "new"),
        "log": (meta.get("document_log") or {}).get(kind, []),
        "profile": meta.get("profile"),
        "deal_status": meta.get("status"),
    }


def _run_single_extraction(deal_id, kind):
    meta = read_deal(ST, deal_id)
    doc_status = dict(meta.get("document_status", {}))
    doc_log = dict(meta.get("document_log", {}))
    try:
        if not os.environ.get("ANTHROPIC_API_KEY"):
            doc_status[kind] = "done"
            doc_log[kind] = doc_log.get(kind, []) + [
                f"[{_now()}] ANTHROPIC_API_KEY not set — skipping AI extraction for this document."]
            meta["document_status"] = doc_status
            meta["document_log"] = doc_log
            if not meta.get("profile"):
                meta["profile"] = normalize_profile({})
                meta["status"] = "review"
            write_deal(ST, deal_id, meta)
            return

        from escrow_agent.deal_extraction import run_single_extraction, merge_profiles
        model = os.environ.get("AI_MODEL_EXTRACTION", "claude-sonnet-5")
        pdf_bytes = ST.get_bytes(f"deals/{deal_id}/documents/{kind}.pdf")
        result, usage = run_single_extraction(kind, pdf_bytes, model=model)

        profile_parts = dict(meta.get("profile_parts", {}))
        profile_parts[kind] = result
        merged = merge_profiles(profile_parts)
        profile = normalize_profile(merged)

        # accumulate token usage across all per-document calls for this deal
        prior = meta.get("ai_usage") or {"input_tokens": 0, "output_tokens": 0, "cost_usd": 0.0, "calls": 0}
        combined_usage = {
            "input_tokens": prior.get("input_tokens", 0) + usage["input_tokens"],
            "output_tokens": prior.get("output_tokens", 0) + usage["output_tokens"],
            "cost_usd": round(prior.get("cost_usd", 0.0) + usage["cost_usd"], 6),
            "calls": prior.get("calls", 0) + usage["calls"],
        }

        doc_status[kind] = "done"
        doc_log[kind] = doc_log.get(kind, []) + [
            f"[{_now()}] Extraction complete: {usage['calls']} call(s), "
            f"{usage['input_tokens']} in / {usage['output_tokens']} out tokens, ${usage['cost_usd']:.4f}."]

        meta["profile_parts"] = profile_parts
        meta["profile"] = profile
        meta["ai_usage"] = combined_usage
        meta["document_status"] = doc_status
        meta["document_log"] = doc_log
        meta["status"] = "review"
        write_deal(ST, deal_id, meta)
    except Exception as e:
        doc_status[kind] = "failed"
        doc_log[kind] = doc_log.get(kind, []) + [
            f"[{_now()}] FAILED: {type(e).__name__}: {e}", traceback.format_exc(limit=3)]
        meta["document_status"] = doc_status
        meta["document_log"] = doc_log
        write_deal(ST, deal_id, meta)


@app.put("/api/deals/{deal_id}/profile")
def update_profile(deal_id: str, profile: dict):
    deal_id = _safe(deal_id)
    try:
        meta = read_deal(ST, deal_id)
    except Exception:
        raise HTTPException(404, "deal not found")
    meta["profile"] = normalize_profile(profile)
    if meta["status"] not in ("review", "ready"):
        meta["status"] = "review"
    write_deal(ST, deal_id, meta)
    return {"ok": True}


@app.post("/api/deals/{deal_id}/confirm")
def confirm_profile(deal_id: str):
    deal_id = _safe(deal_id)
    try:
        meta = read_deal(ST, deal_id)
    except Exception:
        raise HTTPException(404, "deal not found")
    if not meta.get("profile"):
        raise HTTPException(400, "no profile to confirm — extract or enter one first")
    meta["status"] = "ready"
    meta["confirmed_at"] = _now()
    write_deal(ST, deal_id, meta)
    return {"ok": True}


@app.get("/api/deals/{deal_id}/documents/{kind}")
def download_document(deal_id: str, kind: str):
    deal_id, kind = _safe(deal_id), _safe(kind)
    key = f"deals/{deal_id}/documents/{kind}.pdf"
    if not ST.exists(key):
        raise HTTPException(404, "document not found")
    data = ST.get_bytes(key)
    return StreamingResponse(io.BytesIO(data), media_type="application/pdf",
                             headers={"Content-Disposition": f'inline; filename="{kind}.pdf"'})


# ============================================================== DEAL TEMPLATES
@app.get("/api/deals/{deal_id}/templates")
def deal_templates_status(deal_id: str):
    meta = get_deal_template_meta(ST, _safe(deal_id))
    return {"catra": meta.get("catra"), "tra": meta.get("tra")}


@app.post("/api/deals/{deal_id}/templates")
async def upload_deal_template(deal_id: str, kind: str = Form(...), file: UploadFile = File(...)):
    if kind not in ("catra", "tra"):
        raise HTTPException(400, "kind must be 'catra' or 'tra'")
    name = _safe(file.filename)
    data = await file.read()
    saved = set_deal_default_template(ST, _safe(deal_id), kind, name, data)
    return {"ok": True, kind: saved}


# ============================================================== DEAL RUNS (quarterly)
@app.post("/api/deals/{deal_id}/runs")
async def create_run(
    deal_id: str,
    fy: str = Form(...),
    account: str = Form(""),
    ai_assist: bool = Form(True),
    statements: list[UploadFile] = File(...),
    catra_template: Optional[UploadFile] = File(None),
    tra_template: Optional[UploadFile] = File(None),
):
    deal_id = _safe(deal_id)
    try:
        deal = read_deal(ST, deal_id)
    except Exception:
        raise HTTPException(404, "deal not found")
    if not (1 <= len(statements) <= 4):
        raise HTTPException(400, "upload 1-4 quarterly statement files (names must contain Q1..Q4)")

    run_id = datetime.now().strftime("%Y%m%d-%H%M%S") + "-" + uuid.uuid4().hex[:6]
    stmt_names = []
    for f in statements:
        name = _safe(f.filename)
        ST.put_bytes(f"deals/{deal_id}/runs/{run_id}/inputs/{name}", await f.read())
        stmt_names.append(name)

    template_source = {}
    async def _resolve_template(kind, upload):
        if upload is not None and upload.filename:
            name = _safe(upload.filename)
            data = await upload.read()
            ST.put_bytes(f"deals/{deal_id}/runs/{run_id}/inputs/{name}", data)
            template_source[kind] = "uploaded for this run"
            return name
        name, data = get_deal_default_template(ST, deal_id, kind)
        if name is None:
            raise HTTPException(400, f"no {kind.upper()} template provided and no saved default template "
                                     f"exists yet for this deal — upload one for this run, or save a default "
                                     f"in this deal's Templates first")
        ST.put_bytes(f"deals/{deal_id}/runs/{run_id}/inputs/{name}", data)
        template_source[kind] = "saved default"
        return name

    ct = await _resolve_template("catra", catra_template)
    tt = await _resolve_template("tra", tra_template)

    meta = {
        "run_id": run_id, "deal_id": deal_id, "fy": fy.strip(),
        "account": (account.strip() or deal.get("account", "")), "status": "queued", "stage": "QUEUED",
        "stages_done": [], "stages": STAGES, "created_at": _now(), "log": [],
        "ai_assist": bool(ai_assist),
        "inputs": {"statements": stmt_names, "catra_template": ct, "tra_template": tt},
        "template_source": template_source,
        "outputs": [],
    }
    write_run(ST, deal_id, run_id, meta)
    threading.Thread(target=execute_run, args=(ST, deal_id, run_id), daemon=True).start()
    return {"run_id": run_id}


@app.get("/api/deals/{deal_id}/runs")
def get_runs(deal_id: str):
    deal_id = _safe(deal_id)
    runs = []
    for rid in list_runs(ST, deal_id):
        try:
            m = read_run(ST, deal_id, rid)
        except Exception:
            continue
        runs.append({k: m.get(k) for k in
                     ("run_id", "deal_id", "fy", "account", "status", "stage", "stages_done",
                      "verdict", "created_at", "finished_at", "stats", "error", "outputs",
                      "ai_assist", "ai_usage", "template_source")})
    runs.sort(key=lambda m: m.get("created_at") or "", reverse=True)
    return {"runs": runs, "stages": STAGES}


@app.get("/api/deals/{deal_id}/runs/{run_id}")
def get_run(deal_id: str, run_id: str):
    try:
        return read_run(ST, _safe(deal_id), _safe(run_id))
    except Exception:
        raise HTTPException(404, "run not found")


@app.get("/api/deals/{deal_id}/runs/{run_id}/download/{filename}")
def download_run_file(deal_id: str, run_id: str, filename: str):
    deal_id, run_id, filename = _safe(deal_id), _safe(run_id), _safe(filename)
    for folder in ("outputs", "inputs"):
        key = f"deals/{deal_id}/runs/{run_id}/{folder}/{filename}"
        if ST.exists(key):
            data = ST.get_bytes(key)
            mt = ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                  if filename.endswith(".xlsx") else
                  "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                  if filename.endswith(".docx") else "application/octet-stream")
            return StreamingResponse(io.BytesIO(data), media_type=mt,
                                     headers={"Content-Disposition": f'attachment; filename="{filename}"'})
    raise HTTPException(404, "file not found")


@app.post("/api/deals/{deal_id}/runs/{run_id}/rerun")
def rerun(deal_id: str, run_id: str):
    deal_id, run_id = _safe(deal_id), _safe(run_id)
    try:
        meta = read_run(ST, deal_id, run_id)
    except Exception:
        raise HTTPException(404, "run not found")
    meta.update(status="queued", stage="QUEUED", stages_done=[], outputs=[],
                verdict=None, error=None, log=meta.get("log", []) + [f"[{_now()}] re-run requested"])
    write_run(ST, deal_id, run_id, meta)
    threading.Thread(target=execute_run, args=(ST, deal_id, run_id), daemon=True).start()
    return {"ok": True}


@app.delete("/api/deals/{deal_id}/runs/{run_id}")
def delete_run(deal_id: str, run_id: str):
    ST.delete_prefix(f"deals/{_safe(deal_id)}/runs/{_safe(run_id)}")
    return {"ok": True}


# ============================================================== GLOBAL USAGE (across all deals/runs)
@app.get("/api/usage")
def usage_summary():
    total_in = total_out = 0
    total_cost = 0.0
    total_calls = 0
    runs_with_ai = 0
    deals_with_extraction = 0
    for did in list_deals(ST):
        try:
            d = read_deal(ST, did)
        except Exception:
            continue
        u = d.get("ai_usage")
        if u and u.get("calls"):
            deals_with_extraction += 1
            total_in += u.get("input_tokens", 0); total_out += u.get("output_tokens", 0)
            total_cost += u.get("cost_usd", 0.0); total_calls += u.get("calls", 0)
        for rid in list_runs(ST, did):
            try:
                m = read_run(ST, did, rid)
            except Exception:
                continue
            ru = m.get("ai_usage")
            if ru and ru.get("calls"):
                runs_with_ai += 1
                total_in += ru.get("input_tokens", 0); total_out += ru.get("output_tokens", 0)
                total_cost += ru.get("cost_usd", 0.0); total_calls += ru.get("calls", 0)
    return {
        "api_key_configured": bool(os.environ.get("ANTHROPIC_API_KEY")),
        "runs_with_ai_assist": runs_with_ai,
        "deals_with_extraction": deals_with_extraction,
        "total_input_tokens": total_in,
        "total_output_tokens": total_out,
        "total_cost_usd": round(total_cost, 4),
        "total_calls": total_calls,
    }


@app.get("/")
def index():
    return FileResponse(os.path.join(STATIC, "index.html"))


app.mount("/static", StaticFiles(directory=STATIC), name="static")
