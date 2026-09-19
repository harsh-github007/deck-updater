"""Run log: one row per action, written to CSV at the end."""

from __future__ import annotations

import csv
from dataclasses import dataclass, asdict, fields
from pathlib import Path
from typing import List, Optional


@dataclass
class LogEntry:
    deck: str
    slide: int
    shape: str
    directive: str
    reference: str
    action: str
    status: str
    message: str = ""


class RunLog:
    def __init__(self) -> None:
        self.entries: List[LogEntry] = []

    def add(self, **kwargs) -> LogEntry:
        entry = LogEntry(**kwargs)
        self.entries.append(entry)
        return entry

    @property
    def errors(self) -> List[LogEntry]:
        return [e for e in self.entries if e.status == "ERROR"]

    def write_csv(self, path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", newline="", encoding="utf-8") as fh:
            writer = csv.writer(fh)
            writer.writerow([f.name.capitalize() for f in fields(LogEntry)])
            for e in self.entries:
                writer.writerow(asdict(e).values())

    def summary(self) -> str:
        counts = {}
        for e in self.entries:
            counts[e.status] = counts.get(e.status, 0) + 1
        parts = [f"{k}: {v}" for k, v in sorted(counts.items())]
        return ", ".join(parts) if parts else "nothing to do"
