"""Command line interface: ``slidesync run config.json``."""

from __future__ import annotations

import argparse
import sys

from . import __version__
from .config import load_config
from .engine import Engine
from .verify import verify_run


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="slidesync", description="Refresh PowerPoint decks from Excel data.")
    p.add_argument("--version", action="version", version=f"slidesync {__version__}")
    sub = p.add_subparsers(dest="command", required=True)

    run_p = sub.add_parser("run", help="process every deck in a config file")
    run_p.add_argument("config", help="path to a JSON config file")
    run_p.add_argument("-q", "--quiet", action="store_true", help="only print the summary line")
    run_p.add_argument("--verify", action="store_true",
                       help="after the run, read the decks back and check them against the log")
    run_p.add_argument("--baseline", metavar="LOG",
                       help="also diff the run log against an approved log")
    run_p.add_argument("--report", metavar="CSV", help="write verification findings here")

    ver_p = sub.add_parser("verify", help="check an existing run's output against its log")
    ver_p.add_argument("config", help="path to the JSON config file used for the run")
    ver_p.add_argument("--log", help="run log to verify (defaults to the config's log path)")
    ver_p.add_argument("--baseline", metavar="LOG",
                       help="an approved log to diff this run against")
    ver_p.add_argument("--report", metavar="CSV", help="write verification findings here")
    ver_p.add_argument("-q", "--quiet", action="store_true", help="only print the summary line")
    return p


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)
    if args.command == "run":
        cfg = load_config(args.config)
        log = Engine(cfg).run()
        if not args.quiet:
            for e in log.entries:
                if e.status != "OK" or e.action == "deck":
                    print(f"[{e.status}] {e.deck} slide {e.slide} {e.shape} {e.directive} -> {e.message}")
        print(f"Done. {log.summary()}" + (f" Log: {cfg.log}" if cfg.log else ""))
        code = 1 if log.errors else 0
        if args.verify or args.baseline:
            code = max(code, _verify(cfg, None, args.baseline, args.report, args.quiet))
        return code

    if args.command == "verify":
        cfg = load_config(args.config)
        return _verify(cfg, args.log, args.baseline, args.report, args.quiet)
    return 2


def _verify(cfg, log_path, baseline, report_path, quiet) -> int:
    """Run the checks and print what a reviewer needs to see."""
    try:
        report = verify_run(cfg, log_path=log_path, expected_path=baseline)
    except FileNotFoundError as exc:
        print(f"Cannot verify: {exc}")
        return 2
    if not quiet:
        for f in report.findings:
            if f.status in ("MISMATCH", "MISSING", "CHANGED"):
                where = f"{f.deck} {f.slide} {f.shape}".strip()
                print(f"[{f.status}] ({f.check}) {where} {f.directive} -> {f.detail}")
    if report_path:
        report.write_csv(report_path)
    print(f"Verified. {report.summary()}" + (f" Report: {report_path}" if report_path else ""))
    return 1 if report.problems else 0


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
