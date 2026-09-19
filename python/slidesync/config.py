"""Run configuration: which workbooks feed which decks."""

from __future__ import annotations

import json
import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional


@dataclass
class Source:
    alias: str
    path: Path


@dataclass
class Deck:
    template: Path
    output: Path


@dataclass
class Config:
    sources: List[Source]
    decks: List[Deck]
    log: Optional[Path] = None
    base_dir: Path = field(default_factory=Path.cwd)

    @property
    def default_alias(self) -> str:
        return self.sources[0].alias


def _resolve(base: Path, value: str) -> Path:
    p = Path(value).expanduser()
    return p if p.is_absolute() else (base / p)


def load_config(path: str | Path) -> Config:
    """Read a JSON config file.

    Relative paths inside the file are resolved against the file's folder.

    Example::

        {
          "sources": [{"alias": "data", "path": "monthly.xlsx"}],
          "decks":   [{"template": "template.pptx", "output": "output/report.pptx"}],
          "log": "output/log.csv"
        }
    """
    path = Path(path)
    with path.open(encoding="utf-8") as fh:
        raw = json.load(fh)
    base = path.parent.resolve()

    sources = raw.get("sources") or []
    decks = raw.get("decks") or []
    if not sources:
        raise ValueError("config needs at least one entry in 'sources'")
    if not decks:
        raise ValueError("config needs at least one entry in 'decks'")

    cfg_sources = []
    seen = set()
    for i, s in enumerate(sources):
        alias = str(s.get("alias") or f"src{i + 1}")
        # Lookup is case-insensitive, so validation must be too, or one source
        # silently shadows another.
        if alias.lower() in seen:
            raise ValueError(f"duplicate source alias '{alias}' (aliases are case-insensitive)")
        seen.add(alias.lower())
        cfg_sources.append(Source(alias=alias, path=_resolve(base, s["path"])))

    cfg_decks = []
    for d in decks:
        template = _resolve(base, d["template"])
        output = _resolve(base, d.get("output") or _default_output(template))
        cfg_decks.append(Deck(template=template, output=output))

    _check_paths(cfg_sources, cfg_decks)

    log = _resolve(base, raw["log"]) if raw.get("log") else None
    return Config(sources=cfg_sources, decks=cfg_decks, log=log, base_dir=base)


def _key(path: Path) -> str:
    """Comparison key for two paths that may name the same file.

    ``os.path.normcase`` folds case on Windows and leaves POSIX paths alone.
    """
    try:
        resolved = path.resolve()
    except OSError:  # pragma: no cover - unreachable on supported platforms
        resolved = path.absolute()
    return os.path.normcase(str(resolved))


def _check_paths(sources: List[Source], decks: List[Deck]) -> None:
    """Refuse a run that would destroy one of its own inputs.

    Writing a deck over its template consumes the tokens: the first run
    succeeds and every run after it has nothing left to replace.
    """
    inputs = {_key(s.path): f"source '{s.alias}'" for s in sources}
    for d in decks:
        inputs.setdefault(_key(d.template), f"template '{d.template.name}'")

    seen_outputs = {}
    for d in decks:
        out = _key(d.output)
        if out in inputs:
            raise ValueError(
                f"deck '{d.template.name}' would write over its own input "
                f"({inputs[out]}); choose a different output path"
            )
        if out in seen_outputs:
            raise ValueError(
                f"decks '{seen_outputs[out]}' and '{d.template.name}' share the "
                f"output path {d.output}"
            )
        seen_outputs[out] = d.template.name


def _default_output(template: Path) -> str:
    return str(template.with_name(template.stem + "_updated" + template.suffix))
