"""Schedule builder: download XLSX and convert to ICS."""
from __future__ import annotations

from datetime import date, datetime, time
from pathlib import Path
import re
import tempfile
import uuid

import openpyxl
import requests
from icalendar import Calendar, Event
from dateutil import tz
from pydantic import BaseModel
import yaml

DAYS = [
    "понедельник",
    "вторник",
    "среда",
    "четверг",
    "пятница",
    "суббота",
]


class ScanConfig(BaseModel):
    max_header_rows: int = 8
    date_scan_up: int = 8
    max_row: int | None = None


class Config(BaseModel):
    public_link: str
    sheet_name: str | None = None
    timezone: str = "Europe/Moscow"
    year: int
    uid_prefix: str = ""
    scan: ScanConfig = ScanConfig()


TIME_RE = re.compile(r"(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})")
DATE_RE = re.compile(r"(\d{1,2})[.,](\d{1,2})")


def read_config(path: str | Path) -> Config:
    """Read configuration from YAML file."""
    with open(path, "r", encoding="utf8") as f:
        data = yaml.safe_load(f)
    return Config.model_validate(data)


def download_xlsx(public_link: str, dest: Path) -> Path:
    """Download public XLSX from Yandex.Disk to dest."""
    api = "https://cloud-api.yandex.net/v1/disk/public/resources/download"
    r = requests.get(api, params={"public_key": public_link})
    r.raise_for_status()
    href = r.json()["href"]
    r2 = requests.get(href)
    r2.raise_for_status()
    dest.write_bytes(r2.content)
    return dest


def _cell_value(ws, row: int, col: int):
    cell = ws.cell(row=row, column=col)
    if cell.value is not None:
        return cell.value
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            top = ws.cell(rng.min_row, rng.min_col)
            return top.value
    return None


def _is_top_left(ws, row: int, col: int) -> bool:
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return row == rng.min_row and col == rng.min_col
    return True


def _find_date(ws, row: int, col: int, cfg: Config) -> date:
    for r in range(row, max(row - cfg.scan.date_scan_up, 0), -1):
        val = _cell_value(ws, r, col)
        if isinstance(val, str):
            m = DATE_RE.search(val)
            if m:
                d, mth = int(m.group(1)), int(m.group(2))
                return date(cfg.year, mth, d)
    raise ValueError(f"Date not found for cell {row},{col}")


def _iter_events(ws, cfg: Config):
    # day columns
    day_cols: dict[int, str] = {}
    for row in ws.iter_rows(min_row=1, max_row=cfg.scan.max_header_rows):
        for cell in row:
            if isinstance(cell.value, str):
                name = cell.value.strip().lower()
                if name in DAYS:
                    day_cols[cell.column] = name
    if not day_cols:
        return

    # Respect max_row limit if configured
    max_row = cfg.scan.max_row if getattr(cfg, "scan", None) and cfg.scan.max_row else ws.max_row

    # time rows (classic layout: times in column A)
    time_rows: dict[int, tuple[time, time]] = {}
    for cell in ws["A"]:
        if isinstance(cell.value, str):
            m = TIME_RE.search(cell.value)
            if m:
                start = time(int(m.group(1)), int(m.group(2)))
                end = time(int(m.group(3)), int(m.group(4)))
                if cell.row <= max_row:
                    time_rows[cell.row] = (start, end)

    tzinfo = tz.gettz(cfg.timezone)

    # Helper: scan upward in a column to find time range from any header cell
    def _scan_time_up(row: int, col: int) -> tuple[time, time] | None:
        for r2 in range(row, 0, -1):
            v = _cell_value(ws, r2, col)
            if isinstance(v, str):
                m = TIME_RE.search(v)
                if m:
                    start = time(int(m.group(1)), int(m.group(2)))
                    end = time(int(m.group(3)), int(m.group(4)))
                    return (start, end)
        return None

    # Helper: scan upward in a column to find a date (accept xlsx date/datetime or DD.MM in text)
    def _scan_date_up(row: int, col: int, cfg: Config) -> date:
        for r2 in range(row, max(row - cfg.scan.date_scan_up, 0), -1):
            v = _cell_value(ws, r2, col)
            if isinstance(v, datetime):
                return v.date()
            if isinstance(v, date):
                return v
            if isinstance(v, str):
                m = DATE_RE.search(v)
                if m:
                    d, mth = int(m.group(1)), int(m.group(2))
                    return date(cfg.year, mth, d)
        # fallback to original behavior (scan up and error)
        return _find_date(ws, row, col, cfg)

    def _split_summary_desc(text: str) -> tuple[str, str]:
        import re as _re
        # Convert long sequences of spaces/tabs into newlines
        text = _re.sub(r"[\t ]{3,}", "\n", text)
        # Normalize lines, drop numeric-only lines from description
        parts = [ln.strip() for ln in text.splitlines()]
        clean = [ln for ln in parts if ln]
        # First non-empty line as summary
        summary = clean[0] if clean else ""
        # For description, exclude numeric-only lines
        desc_lines = [ln for ln in clean[1:] if not _re.fullmatch(r"\d+", ln)]
        # Move key markers (НАЧАЛО/ЗАЩИТА ОТЧЕТА/ЗАЧЕТ С ОЦЕНКОЙ/ЭКЗАМЕН/ЗАЧЁТ/ЗАЩИТА) into summary together with first detail line
        KEY_MARKERS = {
            "начало",
            "защита отчета",
            "защита отчёта",
            "зачет с оценкой",
            "зачёт с оценкой",
            "экзамен",
            "зачет",
            "зачёт",
            "защита",
        }
        if summary.strip().lower() in KEY_MARKERS and desc_lines:
            summary = f"{summary} — {desc_lines[0]}"
            desc_lines = desc_lines[1:]
        desc = "\n".join(desc_lines)
        return summary, desc

    if time_rows:
        # Classic layout path
        for r, (start, end) in time_rows.items():
            for c in day_cols.keys():
                val = _cell_value(ws, r, c)
                # Only textual cells are considered events; skip dates/numbers
                if not isinstance(val, str):
                    continue
                if not val.strip():
                    continue
                if not _is_top_left(ws, r, c):
                    continue
                dt = _find_date(ws, r, c, cfg)
                summary, desc = _split_summary_desc(str(val))
                dtstart = datetime(dt.year, dt.month, dt.day, start.hour, start.minute, tzinfo=tzinfo)
                dtend = datetime(dt.year, dt.month, dt.day, end.hour, end.minute, tzinfo=tzinfo)
                uid = f"{cfg.uid_prefix}{dt.isoformat()}-{start.hour:02d}{start.minute:02d}-{uuid.uuid5(uuid.NAMESPACE_DNS, summary)}"
                event = Event()
                event.add("summary", summary)
                if desc:
                    event.add("description", desc)
                event.add("dtstart", dtstart)
                event.add("dtend", dtend)
                event.add("uid", uid)
                yield event
    else:
        # Alternative layout: times and dates are in the day columns headers
        # Iterate over grid: for each day column, scan rows beyond headers
        for c in day_cols.keys():
            for r in range(1, max_row + 1):
                val = _cell_value(ws, r, c)
                # Only textual cells are considered events; skip dates/numbers
                if not isinstance(val, str):
                    continue
                if not val.strip():
                    continue
                name = val.strip().lower()
                # Skip header-like cells: day names or pure time ranges
                if name in DAYS:
                    continue
                m_hdr_time = TIME_RE.search(name)
                if m_hdr_time and m_hdr_time.group(0).strip() == name:
                    continue
                if not _is_top_left(ws, r, c):
                    continue
                dt = _scan_date_up(r, c, cfg)
                tpair = _scan_time_up(r, c)
                if not tpair:
                    # Try fallback to time from column A of this row
                    a_val = _cell_value(ws, r, 1)
                    if isinstance(a_val, str):
                        m = TIME_RE.search(a_val)
                        if m:
                            tpair = (
                                time(int(m.group(1)), int(m.group(2))),
                                time(int(m.group(3)), int(m.group(4))),
                            )
                if not tpair:
                    continue
                start, end = tpair
                summary, desc = _split_summary_desc(str(val))
                dtstart = datetime(dt.year, dt.month, dt.day, start.hour, start.minute, tzinfo=tzinfo)
                dtend = datetime(dt.year, dt.month, dt.day, end.hour, end.minute, tzinfo=tzinfo)
                uid = f"{cfg.uid_prefix}{dt.isoformat()}-{start.hour:02d}{start.minute:02d}-{uuid.uuid5(uuid.NAMESPACE_DNS, summary)}"
                event = Event()
                event.add("summary", summary)
                if desc:
                    event.add("description", desc)
                event.add("dtstart", dtstart)
                event.add("dtend", dtend)
                event.add("uid", uid)
                yield event


def build_ics(cfg: Config) -> bytes:
    """Build calendar ICS from config."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx") as tmp:
        download_xlsx(cfg.public_link, Path(tmp.name))
        wb = openpyxl.load_workbook(tmp.name, data_only=True)
        ws = wb[cfg.sheet_name] if cfg.sheet_name else wb.active
        cal = Calendar()
        cal.add("prodid", "-//schedics//")
        cal.add("version", "2.0")
        for ev in _iter_events(ws, cfg):
            cal.add_component(ev)
        return cal.to_ical()


def build_ics_and_events(cfg: Config) -> tuple[bytes, list[dict]]:
    """Build calendar ICS and return a JSON-serializable list of parsed events."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx") as tmp:
        download_xlsx(cfg.public_link, Path(tmp.name))
        wb = openpyxl.load_workbook(tmp.name, data_only=True)
        ws = wb[cfg.sheet_name] if cfg.sheet_name else wb.active
        cal = Calendar()
        cal.add("prodid", "-//schedics//")
        cal.add("version", "2.0")
        events: list[dict] = []
        for ev in _iter_events(ws, cfg):
            cal.add_component(ev)
            # Extract fields for debugging/verification
            try:
                dtstart = ev.decoded("dtstart")
            except Exception:
                dtstart = None
            try:
                dtend = ev.decoded("dtend")
            except Exception:
                dtend = None
            summary = ev.get("summary")
            desc = ev.get("description")
            events.append({
                "summary": str(summary) if summary is not None else None,
                "description": str(desc) if desc is not None else None,
                "dtstart": dtstart.isoformat() if hasattr(dtstart, "isoformat") else None,
                "dtend": dtend.isoformat() if hasattr(dtend, "isoformat") else None,
            })
        return cal.to_ical(), events
