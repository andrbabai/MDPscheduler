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

    # time rows
    time_rows: dict[int, tuple[time, time]] = {}
    for cell in ws["A"]:
        if isinstance(cell.value, str):
            m = TIME_RE.search(cell.value)
            if m:
                start = time(int(m.group(1)), int(m.group(2)))
                end = time(int(m.group(3)), int(m.group(4)))
                time_rows[cell.row] = (start, end)

    tzinfo = tz.gettz(cfg.timezone)

    for r, (start, end) in time_rows.items():
        for c in day_cols.keys():
            val = _cell_value(ws, r, c)
            if not val or not str(val).strip():
                continue
            if not _is_top_left(ws, r, c):
                continue
            dt = _find_date(ws, r, c, cfg)
            summary, *desc_lines = str(val).strip().splitlines()
            desc = "\n".join(desc_lines)
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
