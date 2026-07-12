"""Audit trail (BRD §7): every processing step logged with a timestamp,
persisted to a log file and included as a sheet in the Excel output."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path


class AuditLog:
    def __init__(self, output_dir: str | Path):
        self.entries: list[dict] = []
        self.path = Path(output_dir) / "audit_log.txt"
        self.path.parent.mkdir(parents=True, exist_ok=True)

    def log(self, step: str, detail: str) -> None:
        entry = {"timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                 "step": step, "detail": detail}
        self.entries.append(entry)
        with open(self.path, "a", encoding="utf-8") as fh:
            fh.write(f"{entry['timestamp']} | {step} | {detail}\n")
