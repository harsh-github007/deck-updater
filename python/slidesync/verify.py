"""Quality control for a finished run.

A run log records what the engine believes it wrote. That is a record, not a
check: if the engine is wrong, the log is wrong in exactly the same way. This
module adds the two comparisons that are not circular.

**Read-back** re-opens the deck that was produced and asks whether the slide
actually says what the log claims. It goes through python-pptx against the
saved file, so it exercises none of the code that produced it. This is what
catches a table that silently truncated, a chart cache that did not take, or a
token replacement that landed in the wrong run of a paragraph.

**Baseline comparison** diffs this run's log against a log from a run that was
signed off. It answers the question an analyst actually has at month end: is
anything different from last month other than the numbers? A token that has
stopped resolving, a shape that has been renamed, a table that changed shape, all of those show up here and nowhere else.

Note what is deliberately not offered: an "ideal log" regenerated from the same
workbook by the same reader would agree with the run log by construction, and
would prove nothing.
"""

from __future__ import annotations

import csv
import re
from dataclasses import dataclass, asdict, fields
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

from pptx import Presentation

from .log import LogEntry

TOKEN_RE = re.compile(r"\{\{\s*[^{}]+\s*\}\}")
_TABLE_SHAPE = re.compile(r"^(\d+)x(\d+)$")
_CHART_SHAPE = re.compile(r"^(\d+) series x (\d+) points$")
_PICTURE_SHAPE = re.compile(r"^(\d+)x(\d+) rendered$")


@dataclass
class Finding:
    """One verification result. ``status`` is OK, MISMATCH, MISSING or SKIPPED."""

    check: str          # readback | baseline
    deck: str
    slide: str
    shape: str
    directive: str
    status: str
    detail: str = ""


class Report:
    def __init__(self) -> None:
        self.findings: List[Finding] = []

    def add(self, **kwargs) -> Finding:
        f = Finding(**kwargs)
        self.findings.append(f)
        return f

    @property
    def problems(self) -> List[Finding]:
        return [f for f in self.findings if f.status in ("MISMATCH", "MISSING")]

    @property
    def ok(self) -> bool:
        return not self.problems

    def write_csv(self, path: Path) -> None:
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            w.writerow([f.name.capitalize() for f in fields(Finding)])
            for f in self.findings:
                w.writerow(asdict(f).values())

    def summary(self) -> str:
        counts: Dict[str, int] = {}
        for f in self.findings:
            counts[f.status] = counts.get(f.status, 0) + 1
        if not counts:
            return "nothing to verify"
        return ", ".join(f"{k}: {v}" for k, v in sorted(counts.items()))


# --------------------------------------------------------------------------- logs
def read_log(path) -> List[LogEntry]:
    """Read a log CSV back into entries. Slide stays a string: shared parts are
    logged as ``layout`` or ``master`` rather than a slide number."""
    rows: List[LogEntry] = []
    with Path(path).open(newline="", encoding="utf-8-sig") as fh:
        for row in csv.DictReader(fh):
            norm = {k.strip().lower(): (v or "") for k, v in row.items() if k}
            rows.append(
                LogEntry(
                    deck=norm.get("deck", ""),
                    slide=norm.get("slide", ""),
                    shape=norm.get("shape", ""),
                    directive=norm.get("directive", ""),
                    reference=norm.get("reference", ""),
                    action=norm.get("action", ""),
                    status=norm.get("status", ""),
                    message=norm.get("message", ""),
                )
            )
    return rows


def _key(e: LogEntry) -> Tuple[str, str, str, str, str]:
    return (str(e.deck), str(e.slide), str(e.shape), str(e.directive), str(e.action))


def _indexed(entries: Iterable[LogEntry]) -> Dict[Tuple, List[LogEntry]]:
    """Group by key. The same token can appear twice in one shape, so each key
    holds a list and the comparison walks them in order."""
    out: Dict[Tuple, List[LogEntry]] = {}
    for e in entries:
        out.setdefault(_key(e), []).append(e)
    return out


# ----------------------------------------------------------------------- baseline
def compare_logs(expected: Iterable[LogEntry], actual: Iterable[LogEntry],
                 report: Optional[Report] = None) -> Report:
    """Diff an approved log against this run's log."""
    report = report or Report()
    exp, act = _indexed(expected), _indexed(actual)

    for key in sorted(set(exp) | set(act)):
        deck, slide, shape, directive, action = key
        if action == "deck":
            continue                      # bookkeeping row, not a value
        e_list, a_list = exp.get(key, []), act.get(key, [])
        common = min(len(e_list), len(a_list))

        for i in range(common):
            e, a = e_list[i], a_list[i]
            if e.status != a.status:
                report.add(check="baseline", deck=deck, slide=slide, shape=shape,
                           directive=directive, status="MISMATCH",
                           detail=f"status was {e.status}, now {a.status}: {a.message}")
            elif e.message != a.message:
                # A changed value is expected at month end; it is reported as
                # information so a reviewer can scan it, not as a failure.
                report.add(check="baseline", deck=deck, slide=slide, shape=shape,
                           directive=directive, status="CHANGED",
                           detail=f"{e.message!r} -> {a.message!r}")
            else:
                report.add(check="baseline", deck=deck, slide=slide, shape=shape,
                           directive=directive, status="OK", detail=a.message)

        for e in e_list[common:]:
            report.add(check="baseline", deck=deck, slide=slide, shape=shape,
                       directive=directive, status="MISSING",
                       detail=f"in the baseline ({e.message!r}) but not written this run")
        for a in a_list[common:]:
            report.add(check="baseline", deck=deck, slide=slide, shape=shape,
                       directive=directive, status="MISMATCH",
                       detail=f"written this run ({a.message!r}) but not in the baseline")
    return report


# ----------------------------------------------------------------------- read-back
def _walk(shapes):
    for sh in shapes:
        yield sh
        if sh.shape_type == 6 and hasattr(sh, "shapes"):   # GROUP
            yield from _walk(sh.shapes)


def _shape_text(shape) -> str:
    parts = []
    if shape.has_text_frame:
        parts.append(shape.text_frame.text)
    if getattr(shape, "has_table", False) and shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                parts.append(cell.text)
    return "\n".join(parts)


def _collect(prs) -> Dict[str, Dict[str, list]]:
    """slide key -> shape name -> shapes. Layouts and masters share one bucket
    each, matching how the engine logs them."""
    index: Dict[str, Dict[str, list]] = {}

    def put(key, shapes):
        bucket = index.setdefault(key, {})
        for sh in _walk(shapes):
            bucket.setdefault(sh.name, []).append(sh)

    for n, slide in enumerate(prs.slides, start=1):
        put(str(n), slide.shapes)
    for layout in prs.slide_layouts:
        put("layout", layout.shapes)
    for master in prs.slide_masters:
        put("master", master.shapes)
    return index


def readback(deck_path, entries: Iterable[LogEntry], deck_name: str = "",
             report: Optional[Report] = None) -> Report:
    """Re-open a produced deck and check it against what the log claims."""
    report = report or Report()
    deck_path = Path(deck_path)
    deck_name = deck_name or deck_path.name

    if not deck_path.exists():
        report.add(check="readback", deck=deck_name, slide="", shape="", directive="",
                   status="MISSING", detail=f"output not found: {deck_path}")
        return report
    try:
        prs = Presentation(str(deck_path))
    except Exception as exc:  # noqa: BLE001 - a corrupt output is a finding
        report.add(check="readback", deck=deck_name, slide="", shape="", directive="",
                   status="MISMATCH", detail=f"cannot open output: {exc}")
        return report

    index = _collect(prs)

    for e in entries:
        if e.action == "deck":
            continue
        slide_key = str(e.slide)
        shapes = index.get(slide_key, {}).get(e.shape, [])
        add = lambda status, detail: report.add(  # noqa: E731 - local shorthand
            check="readback", deck=deck_name, slide=slide_key, shape=e.shape,
            directive=e.directive, status=status, detail=detail)

        if not shapes:
            add("MISSING", f"no shape named {e.shape!r} on {slide_key}")
            continue
        shape = shapes[0]

        if e.status == "ERROR":
            # The promise for a failed reference is that the slide is left
            # exactly as written, so the token must still be there.
            if e.directive and e.directive.startswith("{{"):
                if e.directive in _shape_text(shape):
                    add("OK", "unresolved token left in place, as documented")
                else:
                    add("MISMATCH", f"failed token {e.directive!r} is not on the slide")
            else:
                add("SKIPPED", "failed directive; nothing to read back")
            continue

        if e.action == "text":
            text = _shape_text(shape)
            if e.message == "":
                add("SKIPPED", "resolved to an empty value; nothing to match")
            elif e.message in text:
                add("OK", e.message)
            else:
                add("MISMATCH", f"log says {e.message!r}; the slide does not contain it")
            if TOKEN_RE.search(text) and e.status == "OK":
                # Not a failure by itself, another token in the same shape may
                # legitimately have failed, but worth surfacing.
                add("SKIPPED", "shape still contains an unreplaced token")

        elif e.action == "table":
            m = _TABLE_SHAPE.match(e.message or "")
            if not (getattr(shape, "has_table", False) and shape.has_table):
                add("MISMATCH", "log recorded a table; the shape is not a table")
            elif not m:
                add("SKIPPED", f"cannot parse recorded size {e.message!r}")
            else:
                want = (int(m.group(1)), int(m.group(2)))
                got = (len(shape.table.rows), len(shape.table.columns))
                add("OK" if got == want else "MISMATCH",
                    f"{got[0]}x{got[1]} on the slide, log recorded {want[0]}x{want[1]}")

        elif e.action == "chart":
            m = _CHART_SHAPE.match(e.message or "")
            if not (getattr(shape, "has_chart", False) and shape.has_chart):
                add("MISMATCH", "log recorded a chart; the shape is not a chart")
            elif not m:
                add("SKIPPED", f"cannot parse recorded size {e.message!r}")
            else:
                want = (int(m.group(1)), int(m.group(2)))
                series = list(shape.chart.series)
                points = max((len(list(s.values)) for s in series), default=0)
                got = (len(series), points)
                add("OK" if got == want else "MISMATCH",
                    f"{got[0]} series x {got[1]} points on the slide, "
                    f"log recorded {want[0]} x {want[1]}")

        elif e.action == "picture":
            if getattr(shape, "image", None) is None:
                add("MISMATCH", "log recorded a picture; the shape has no image")
            else:
                # The bytes cannot be checked against the range without
                # re-rendering, which would be circular. Presence and size only.
                add("OK", f"image present, {shape.image.size[0]}x{shape.image.size[1]}px")
        else:
            add("SKIPPED", f"no read-back for action {e.action!r}")

    return report


# --------------------------------------------------------------------------- top
def verify_run(config, log_path=None, expected_path=None) -> Report:
    """Read-back every deck in *config*, and diff against a baseline if given."""
    report = Report()
    log_file = Path(log_path) if log_path else config.log
    if not log_file or not Path(log_file).exists():
        raise FileNotFoundError(
            "no run log to verify; run first with a \"log\" path in the config"
        )
    entries = read_log(log_file)

    outputs = {str(e.deck): e.reference for e in entries
               if e.action == "deck" and e.status == "OK" and e.reference}
    by_deck: Dict[str, List[LogEntry]] = {}
    for e in entries:
        by_deck.setdefault(str(e.deck), []).append(e)

    for deck in config.decks:
        name = deck.template.name
        out = outputs.get(name, str(deck.output))
        readback(out, by_deck.get(name, []), deck_name=name, report=report)

    if expected_path:
        compare_logs(read_log(expected_path), entries, report=report)
    return report
