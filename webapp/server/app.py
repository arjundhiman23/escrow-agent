"""Escrow Analyst Console — FastAPI backend.

POST /api/runs                   create a run (multipart) and execute in background
GET  /api/runs                   list runs (newest first)
GET  /api/runs/{id}              run detail (meta incl. log, stats, verdict)
GET  /api/runs/{id}/download/{f} download a generated report or input
POST /api/runs/{id}/rerun        re-execute with the same inputs
DELETE /api/runs/{id}            remove a run and its files
GET  /                           frontend
"""
import io, os, re, threading, uuid
from datetime import datetime, timezone

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse, FileResponse
from fastapi.staticfiles import StaticFiles

from webapp.server.storage import get_storage, read_meta, write_meta
from webapp.server.runner import execute_run, STAGES

app = FastAPI(title="Escrow Analyst Console")
ST = get_storage()
STATIC = os.path.join(os.path.dirname(__file__), "..", "static")
SAFE = re.compile(r"[^A-Za-z0-9._ -]")


def _now():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _safe(name):
    return SAFE.sub("_", os.path.basename(name or "file"))


@app.post("/api/runs")
async def create_run(
    deal_name: str = Form(...),
    fy: str = Form(...),
    account: str = Form("922020065877321"),
    ai_assist: bool = Form(True),
    statements: list[UploadFile] = File(...),
    catra_template: UploadFile = File(...),
    tra_template: UploadFile = File(...),
):
    if not (1 <= len(statements) <= 4):
        raise HTTPException(400, "upload 1-4 quarterly statement files (names must contain Q1..Q4)")
    run_id = datetime.now().strftime("%Y%m%d-%H%M%S") + "-" + uuid.uuid4().hex[:6]
    stmt_names = []
    for f in statements:
        name = _safe(f.filename)
        ST.put_bytes(f"runs/{run_id}/inputs/{name}", await f.read())
        stmt_names.append(name)
    ct, tt = _safe(catra_template.filename), _safe(tra_template.filename)
    ST.put_bytes(f"runs/{run_id}/inputs/{ct}", await catra_template.read())
    ST.put_bytes(f"runs/{run_id}/inputs/{tt}", await tra_template.read())

    meta = {
        "run_id": run_id, "deal_name": deal_name.strip(), "fy": fy.strip(),
        "account": account.strip(), "status": "queued", "stage": "QUEUED",
        "stages_done": [], "stages": STAGES, "created_at": _now(), "log": [],
        "ai_assist": bool(ai_assist),
        "inputs": {"statements": stmt_names, "catra_template": ct, "tra_template": tt},
        "outputs": [],
    }
    write_meta(ST, run_id, meta)
    threading.Thread(target=execute_run, args=(ST, run_id), daemon=True).start()
    return {"run_id": run_id}


@app.get("/api/runs")
def list_runs():
    runs = []
    for rid in ST.list_dirs("runs"):
        try:
            m = read_meta(ST, rid)
        except Exception:
            continue
        runs.append({k: m.get(k) for k in
                     ("run_id", "deal_name", "fy", "account", "status", "stage", "stages_done",
                      "verdict", "created_at", "finished_at", "stats", "error", "outputs",
                      "ai_assist", "ai_usage")})
    runs.sort(key=lambda m: m.get("created_at") or "", reverse=True)
    return {"runs": runs, "stages": STAGES}


@app.get("/api/usage")
def usage_summary():
    """Aggregate Anthropic API token usage + cost across all runs, plus
    whether an API key is configured (never returns the key itself)."""
    total_in = total_out = 0
    total_cost = 0.0
    total_calls = 0
    runs_with_ai = 0
    by_model = {}
    for rid in ST.list_dirs("runs"):
        try:
            m = read_meta(ST, rid)
        except Exception:
            continue
        u = m.get("ai_usage")
        if not u or not u.get("calls"):
            continue
        runs_with_ai += 1
        total_in += u.get("input_tokens", 0)
        total_out += u.get("output_tokens", 0)
        total_cost += u.get("cost_usd", 0.0)
        total_calls += u.get("calls", 0)
        for ev in u.get("events", []):
            bm = by_model.setdefault(ev["model"], {"input_tokens": 0, "output_tokens": 0, "cost_usd": 0.0, "calls": 0})
            bm["input_tokens"] += ev["input_tokens"]
            bm["output_tokens"] += ev["output_tokens"]
            bm["cost_usd"] = round(bm["cost_usd"] + ev["cost_usd"], 6)
            bm["calls"] += 1
    return {
        "api_key_configured": bool(os.environ.get("ANTHROPIC_API_KEY")),
        "runs_with_ai_assist": runs_with_ai,
        "total_input_tokens": total_in,
        "total_output_tokens": total_out,
        "total_cost_usd": round(total_cost, 4),
        "total_calls": total_calls,
        "by_model": by_model,
    }


@app.get("/api/runs/{run_id}")
def run_detail(run_id: str):
    try:
        return read_meta(ST, _safe(run_id))
    except Exception:
        raise HTTPException(404, "run not found")


@app.get("/api/runs/{run_id}/download/{filename}")
def download(run_id: str, filename: str):
    run_id, filename = _safe(run_id), _safe(filename)
    for folder in ("outputs", "inputs"):
        key = f"runs/{run_id}/{folder}/{filename}"
        if ST.exists(key):
            data = ST.get_bytes(key)
            mt = ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                  if filename.endswith(".xlsx") else
                  "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                  if filename.endswith(".docx") else "application/octet-stream")
            return StreamingResponse(io.BytesIO(data), media_type=mt,
                                     headers={"Content-Disposition": f'attachment; filename="{filename}"'})
    raise HTTPException(404, "file not found")


@app.post("/api/runs/{run_id}/rerun")
def rerun(run_id: str):
    run_id = _safe(run_id)
    try:
        meta = read_meta(ST, run_id)
    except Exception:
        raise HTTPException(404, "run not found")
    meta.update(status="queued", stage="QUEUED", stages_done=[], outputs=[],
                verdict=None, error=None, log=meta.get("log", []) + [f"[{_now()}] re-run requested"])
    write_meta(ST, run_id, meta)
    threading.Thread(target=execute_run, args=(ST, run_id), daemon=True).start()
    return {"ok": True}


@app.delete("/api/runs/{run_id}")
def delete_run(run_id: str):
    ST.delete_prefix(f"runs/{_safe(run_id)}")
    return {"ok": True}


@app.get("/")
def index():
    return FileResponse(os.path.join(STATIC, "index.html"))


app.mount("/static", StaticFiles(directory=STATIC), name="static")
