"""Command line interface for schedule ICS builder."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .builder import build_ics, read_config


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="schedics")
    sub = parser.add_subparsers(dest="cmd")

    p_build = sub.add_parser("build", help="build .ics file")
    p_build.add_argument("--config", required=True)
    p_build.add_argument("--output", required=True)

    p_print = sub.add_parser("print", help="print .ics to stdout")
    p_print.add_argument("--config", required=True)

    args = parser.parse_args(argv)

    if args.cmd == "build":
        cfg = read_config(args.config)
        ics = build_ics(cfg)
        Path(args.output).write_bytes(ics)
        return 0
    if args.cmd == "print":
        cfg = read_config(args.config)
        ics = build_ics(cfg)
        sys.stdout.write(ics.decode())
        return 0
    parser.print_help()
    return 1


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
