"""Command line interface for schedule ICS builder."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .builder import build_ics, read_config
from . import builder as _builder
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
import openpyxl
import tempfile
from pathlib import Path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="schedics")
    sub = parser.add_subparsers(dest="cmd")

    p_build = sub.add_parser("build", help="build .ics file")
    p_build.add_argument("--config", required=True)
    p_build.add_argument("--output", required=True)

    p_print = sub.add_parser("print", help="print .ics to stdout")
    p_print.add_argument("--config", required=True)

    p_inspect = sub.add_parser("inspect", help="inspect a cell and derive date/time")
    p_inspect.add_argument("--config", required=True)
    p_inspect.add_argument("--cell", required=True, help="Cell address, e.g. D6")

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
    if args.cmd == "inspect":
        cfg = read_config(args.config)
        col_letter, row = coordinate_from_string(args.cell)
        col = column_index_from_string(col_letter)
        with tempfile.NamedTemporaryFile(suffix=".xlsx") as tmp:
            _builder.download_xlsx(cfg.public_link, Path(tmp.name))
            wb = openpyxl.load_workbook(tmp.name, data_only=True)
            ws = wb[cfg.sheet_name] if cfg.sheet_name else wb.active

            # Value at the target cell (respect merged cells)
            val = _builder._cell_value(ws, row, col)

            # Row header (time) at column A, or scan upwards in this column if not found
            row_header = _builder._cell_value(ws, row, 1)
            time_info = None
            def _extract_time(s: str):
                m = _builder.TIME_RE.search(s)
                if m:
                    return (
                        int(m.group(1)), int(m.group(2)),
                        int(m.group(3)), int(m.group(4)),
                    )
                return None
            if isinstance(row_header, str):
                time_info = _extract_time(row_header)
            if not time_info:
                # Scan upward in the same column for a time range
                for r2 in range(row, 0, -1):
                    h2 = _builder._cell_value(ws, r2, col)
                    if isinstance(h2, str):
                        time_info = _extract_time(h2)
                        if time_info:
                            break

            # Find date scanning upward in this column
            found_date = None
            header_text = None
            from datetime import date, datetime
            for r in range(row, max(row - cfg.scan.date_scan_up, 0), -1):
                hval = _builder._cell_value(ws, r, col)
                if isinstance(hval, (datetime, date)):
                    dtt = hval if isinstance(hval, datetime) else datetime(hval.year, hval.month, hval.day)
                    found_date = dtt.date()
                    header_text = str(hval)
                    break
                if isinstance(hval, str):
                    header_text = hval
                    m = _builder.DATE_RE.search(hval)
                    if m:
                        d, mth = int(m.group(1)), int(m.group(2))
                        found_date = date(cfg.year, mth, d)
                        break

            print(f"Cell {args.cell} value: {val!r}")
            print(f"Row header (A{row}): {row_header!r}")
            if time_info:
                print(f"Parsed time: {time_info[0]:02d}:{time_info[1]:02d} - {time_info[2]:02d}:{time_info[3]:02d}")
            else:
                print("Parsed time: (not found)")
            print(f"Column header text up: {header_text!r}")
            print(f"Found date: {found_date}")
            if found_date and time_info and val and str(val).strip():
                from datetime import datetime
                from dateutil import tz
                tzinfo = tz.gettz(cfg.timezone)
                dtstart = datetime(found_date.year, found_date.month, found_date.day, time_info[0], time_info[1], tzinfo=tzinfo)
                dtend = datetime(found_date.year, found_date.month, found_date.day, time_info[2], time_info[3], tzinfo=tzinfo)
                print(f"Event: '{str(val).splitlines()[0]}' @ {dtstart.isoformat()} -> {dtend.isoformat()}")
            else:
                print("Event: insufficient data to build precise datetime")
        return 0
    parser.print_help()
    return 1


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
