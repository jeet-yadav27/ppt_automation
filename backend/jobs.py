"""One report job per request so simultaneous downloads never share a file."""
from __future__ import annotations

import secrets
import threading
import uuid
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path

from backend import engine

ROOT = Path(__file__).resolve().parents[1]
JOB_DIR = ROOT / "output" / "jobs"
EXECUTOR = ThreadPoolExecutor(max_workers=4, thread_name_prefix="ftir-report")
LOCK = threading.Lock()
JOBS: dict[str, "Job"] = {}
MAX_JOBS = 40


@dataclass
class Job:
    id: str
    name: str
    token: str
    status: str = "queued"
    created: str = field(default_factory=lambda: datetime.now(timezone.utc).isoformat(timespec="seconds"))
    error: str | None = None
    summary: dict | None = None
    path: Path | None = None
    source_path: Path | None = None
    source_label: str = "server dataset"

    def public(self) -> dict:
        summary = self.summary or {}
        return {
            "id": self.id,
            "name": self.name,
            "status": self.status,
            "created": self.created,
            "error": self.error,
            "kpis": summary.get("kpis"),
            "charts": summary.get("charts"),
            "insights": summary.get("insights"),
            "recommendations": summary.get("recommendations"),
            "filename": self.path.name if self.path else None,
            "source": self.source_label,
        }

    def activity(self) -> dict:
        return {
            "id": self.id,
            "name": self.name,
            "status": self.status,
            "created": self.created,
        }


def _trim_locked() -> None:
    finished = [job for job in JOBS.values() if job.status in {"ready", "error"}]
    overflow = len(JOBS) - MAX_JOBS
    if overflow <= 0:
        return
    finished.sort(key=lambda job: job.created)
    for job in finished[:overflow]:
        JOBS.pop(job.id, None)
        if job.path and job.path.exists():
            job.path.unlink(missing_ok=True)
        if job.source_path and job.source_path.exists():
            job.source_path.unlink(missing_ok=True)


def _run(job_id: str) -> None:
    with LOCK:
        job = JOBS.get(job_id)
        if job is None:
            return
        job.status = "running"
    path = JOB_DIR / f"{job_id}.pptx"
    try:
        frame = engine.frame_from_csv(job.source_path) if job.source_path else None
        summary = engine.build_report(job.name, path, frame)
    except Exception as exc:
        with LOCK:
            current = JOBS.get(job_id)
            if current is not None:
                current.status = "error"
                current.error = str(exc)
        return
    with LOCK:
        current = JOBS.get(job_id)
        if current is not None:
            current.status = "ready"
            current.summary = summary
            current.path = path


def submit(name: str, source_path: Path | None = None, source_label: str = "server dataset") -> Job:
    clean = " ".join(name.split())[:60] or "Analyst"
    job = Job(
        id=uuid.uuid4().hex,
        name=clean,
        token=secrets.token_urlsafe(24),
        source_path=source_path,
        source_label=source_label,
    )
    with LOCK:
        JOBS[job.id] = job
        _trim_locked()
    EXECUTOR.submit(_run, job.id)
    return job


def get_job(job_id: str) -> Job | None:
    with LOCK:
        return JOBS.get(job_id)


def activity() -> list[dict]:
    with LOCK:
        jobs = sorted(JOBS.values(), key=lambda job: job.created, reverse=True)
        return [job.activity() for job in jobs[:20]]


def token_matches(job: Job, token: str | None) -> bool:
    if not token:
        return False
    return secrets.compare_digest(job.token, token)
