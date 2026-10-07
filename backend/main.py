"""FTIR report service. Each request owns its own PowerPoint file."""
from __future__ import annotations

import uuid
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse

from backend import engine, jobs

ROOT = Path(__file__).resolve().parents[1]
DIST = ROOT / "frontend" / "dist"


@asynccontextmanager
async def lifespan(_app: FastAPI):
    jobs.JOB_DIR.mkdir(parents=True, exist_ok=True)
    engine.load()
    yield


app = FastAPI(title="FTIR KPI", version="1.0.0", lifespan=lifespan)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://127.0.0.1:5173"],
    allow_methods=["*"],
    allow_headers=["*"],
)


MAX_UPLOAD_BYTES = 100 * 1024 * 1024
UPLOAD_DIR = ROOT / "output" / "uploads"


@app.get("/api/health")
def health() -> dict:
    frame = engine.DATA
    return {
        "ok": frame is not None,
        "rows": 0 if frame is None else int(len(frame)),
        "source": None if engine.SOURCE is None else engine.SOURCE.name,
        "running": sum(1 for item in jobs.activity() if item["status"] in {"queued", "running"}),
    }


@app.get("/api/activity")
def activity() -> dict:
    items = jobs.activity()
    return {
        "jobs": items,
        "running": sum(1 for item in items if item["status"] in {"queued", "running"}),
    }


@app.post("/api/reports", status_code=202)
async def create_report(
    name: str = Form(default="Analyst"),
    file: UploadFile | None = File(default=None),
) -> dict:
    upload_path = None
    label = engine.SOURCE.name if engine.SOURCE else "server dataset"
    has_file = file is not None and bool(file.filename)
    if has_file:
        if Path(file.filename).suffix.lower() != ".csv":
            raise HTTPException(status_code=400, detail="Upload a .csv file")
        payload = await file.read()
        if len(payload) > MAX_UPLOAD_BYTES:
            raise HTTPException(status_code=413, detail="CSV must be 100 MB or smaller")
        if not payload.strip():
            raise HTTPException(status_code=400, detail="The uploaded file is empty")
        UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        upload_path = UPLOAD_DIR / f"{uuid.uuid4().hex}.csv"
        upload_path.write_bytes(payload)
        label = Path(file.filename).name[:120]
    elif engine.DATA is None:
        raise HTTPException(status_code=503, detail="FTIR dataset is not loaded. Upload a CSV.")
    job = jobs.submit(name, upload_path, label)
    return {"id": job.id, "token": job.token, "status": job.status, "name": job.name, "source": label}


@app.get("/api/reports/{job_id}")
def report_status(job_id: str, x_report_token: str | None = Header(default=None)) -> dict:
    job = jobs.get_job(job_id)
    if job is None or not jobs.token_matches(job, x_report_token):
        raise HTTPException(status_code=404, detail="Report not found")
    return job.public()


@app.get("/api/reports/{job_id}/download")
def download_report(job_id: str, x_report_token: str | None = Header(default=None)):
    job = jobs.get_job(job_id)
    if job is None or not jobs.token_matches(job, x_report_token):
        raise HTTPException(status_code=404, detail="Report not found")
    if job.status != "ready" or job.path is None or not job.path.is_file():
        raise HTTPException(status_code=409, detail="Report is not ready")
    safe = "".join(ch if ch.isalnum() else "_" for ch in job.name).strip("_") or "Analyst"
    filename = f"FTIR_KPI_{safe}.pptx"
    return FileResponse(
        job.path,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=filename,
    )


def _frontend_file(path: str) -> FileResponse | None:
    if not DIST.is_dir():
        return None
    candidate = (DIST / path).resolve()
    if DIST.resolve() not in candidate.parents and candidate != DIST.resolve():
        return None
    if candidate.is_file():
        return FileResponse(candidate)
    index = DIST / "index.html"
    if index.is_file():
        return FileResponse(index)
    return None


@app.get("/")
def index():
    page = _frontend_file("index.html")
    if page is not None:
        return page
    return JSONResponse(
        {
            "service": "FTIR KPI",
            "docs": "/docs",
            "hint": "Build the React app with: cd frontend && npm install && npm run build",
        }
    )


@app.get("/{path:path}")
def frontend(path: str):
    if path.startswith("api/"):
        raise HTTPException(status_code=404, detail="Not found")
    page = _frontend_file(path)
    if page is None:
        raise HTTPException(status_code=404, detail="Frontend is not built")
    return page
