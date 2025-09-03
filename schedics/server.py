"""FastAPI server serving schedule ICS."""
from __future__ import annotations

import os
from pathlib import Path

from fastapi import FastAPI, Response, Request
from fastapi.responses import HTMLResponse
import datetime as _dt
import calendar as _cal

from .builder import build_ics_and_events, read_config

CONFIG_PATH = Path(os.environ.get("SCHEDICS_CONFIG", "config.yml"))

app = FastAPI()
_cache: dict[str, bytes | None] = {"ics": None}
_events_cache: list[dict] = []


def load_cache() -> None:
    cfg = read_config(CONFIG_PATH)
    ics, events = build_ics_and_events(cfg)
    _cache["ics"] = ics
    # Replace events cache atomically
    global _events_cache
    _events_cache = events


@app.on_event("startup")
async def _startup() -> None:
    load_cache()


@app.get("/schedule.ics")
async def get_schedule() -> Response:
    return Response(content=_cache["ics"], media_type="text/calendar")


@app.post("/refresh")
async def refresh() -> dict[str, str]:
    load_cache()
    return {"status": "ok"}


@app.get("/events")
async def get_events() -> list[dict]:
    # Serve the last parsed events for verification
    return _events_cache


@app.get("/", response_class=HTMLResponse)
async def index(request: Request) -> str:
    # Month selection via ?m=YYYY-MM, clamped between 2025-01 and 2026-01
    _cal.setfirstweekday(_cal.MONDAY)
    today = _dt.date.today()
    mparam = request.query_params.get("m")
    try:
        if mparam:
            y, m = map(int, mparam.split("-", 1))
            base = _dt.date(y, m, 1)
        else:
            base = _dt.date(today.year, today.month, 1)
    except Exception:
        base = _dt.date(today.year, today.month, 1)
    min_month = _dt.date(2025, 1, 1)
    max_month = _dt.date(2026, 1, 1)
    if base < min_month:
        base = min_month
    if base > max_month:
        base = max_month

    grid = _cal.monthcalendar(base.year, base.month)  # list of weeks, 0 = out of month
    month_name = base.strftime("%B %Y")
    weekdays = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]

    # Render calendar rows
    cal_rows = []
    for week in grid:
        tds = []
        for d in week:
            if d == 0:
                tds.append('<td class="muted">&nbsp;</td>')
            else:
                tds.append(f'<td>{d}</td>')
        cal_rows.append("<tr>" + "".join(tds) + "</tr>")
    cal_html = "".join(cal_rows)

    # Recompute calendar with parsed events and navigation
    by_day: dict[int, list[dict]] = {}
    for ev in _events_cache:
        ds = ev.get("dtstart")
        if not ds:
            continue
        try:
            dt = _dt.datetime.fromisoformat(ds)
        except Exception:
            continue
        if dt.year == base.year and dt.month == base.month:
            by_day.setdefault(dt.day, []).append({
                "time": dt.strftime("%H:%M"),
                "summary": ev.get("summary") or "",
            })
    for lst in by_day.values():
        lst.sort(key=lambda x: x["time"])

    cal_rows2 = []
    for week in grid:
        tds = []
        for d in week:
            if d == 0:
                tds.append('<td class="muted">&nbsp;</td>')
            else:
                items = by_day.get(d, [])
                ev_html = "".join(
                    f"<div class=\"ev-item\"><span class=\"ev-time\">{e['time']}</span><span class=\"ev-title\">{e['summary']}</span></div>"
                    for e in items
                )
                tds.append(
                    f'<td><div class="day"><div class="day-num">{d}</div><div class="ev-list">{ev_html}</div></div></td>'
                )
        cal_rows2.append("<tr>" + "".join(tds) + "</tr>")
    cal_html2 = "".join(cal_rows2)

    # Month navigation links
    def _shift_month(d: _dt.date, delta: int) -> _dt.date:
        y = d.year + (d.month + delta - 1) // 12
        m = (d.month + delta - 1) % 12 + 1
        return _dt.date(y, m, 1)
    prev_m = _shift_month(base, -1)
    next_m = _shift_month(base, +1)
    prev_allowed = prev_m >= _dt.date(2025, 1, 1)
    next_allowed = next_m <= _dt.date(2026, 1, 1)
    prev_href = f"/?m={prev_m.strftime('%Y-%m')}" if prev_allowed else "#"
    next_href = f"/?m={next_m.strftime('%Y-%m')}" if next_allowed else "#"

    return (
        "<!doctype html>"
        "<html lang=\"ru\">"
        "<head>"
        "<meta charset=\"utf-8\">"
        "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">"
        "<title>Расписание УЦП-24 РАНХиГС на 3 семестр</title>"
        "<style>"
        "body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;max-width:960px;margin:40px auto;padding:0 16px;line-height:1.6;background:#f5f5dc;color:#333}"
        "h1{margin:0 0 20px;text-align:center;font-size:28px}"
        "h2{margin:24px 0 12px;font-size:20px;color:#333}"
        "h3{margin:16px 0 8px;font-size:16px;color:#333}"
        "p.small, .hint{font-size:13px;color:#666}"
        ""
        ".section{margin:20px 0;padding-top:8px;border-top:1px solid #e0e0e0}"
        ".card{padding:16px;border:1px solid #e0e0e0;border-radius:12px;margin:16px 0;background:#fff}"
        ""
        ".btn{display:block;width:100%;box-sizing:border-box;margin:10px 0;padding:12px 18px;background:#2563eb;color:#fff;text-decoration:none;border-radius:10px;text-align:center;font-weight:600;transition:background .15s,opacity .15s,border-color .15s,color .15s}"
        ".btn:hover{opacity:.9}"
        ".btn:active{opacity:.85}"
        ".btn.download{background:#ff4f00;color:#fff;border:2px solid transparent}"
        ".btn.secondary{background:#d9d99b;color:#000;border:2px solid transparent}"
        ".btn-outline{background:#fff;color:#3a86ff;border:2px solid #3a86ff}"
        ".btn-outline:hover{background:#f3f4f6}"
        ""
        ".calendar{width:100%;border-collapse:collapse;border:1px solid #e0e0e0;border-radius:12px;overflow:hidden}"
        ".calendar th, .calendar td{border:1px solid #e0e0e0;padding:10px;vertical-align:top}"
        ".calendar th{background:#fafafa;font-weight:600;color:#666;text-align:center}"
        ".calendar td{background:#fff;text-align:left}"
        ".calendar td.muted{color:#cccccc;background:#fff}"
        ".calendar td:hover{background:#f3f4f6}"
        ".cal-head{display:flex;justify-content:space-between;align-items:center;margin:8px 0 12px;gap:8px}"
        ".cal-nav{display:flex;gap:8px}"
        ".cal-nav a{padding:6px 10px;border:1px solid #e0e0e0;border-radius:8px;color:#3a86ff;background:#fff}"
        ".cal-nav a.disabled{pointer-events:none;opacity:.5}"
        ".day{min-height:68px} .day-num{font-weight:600;margin-bottom:4px;color:#333;text-align:right}"
        ".ev-item{font-size:12px;color:#333;margin:2px 0;padding:2px 4px;border-radius:6px;background:#fafafa}"
        ".ev-time{color:#666;margin-right:4px}"
        ""
        "label{display:block;margin:8px 0 6px;color:#333}"
        "input[type=text], select{width:100%;padding:10px;border:1px solid #e0e0e0;border-radius:8px;font-size:14px}"
        ".file-wrap{display:flex;gap:12px;align-items:center;flex-wrap:wrap}"
        ".file-wrap input[type=file]{display:none}"
        ".file-wrap label{display:inline-block}"
        ""
        ".callout{background:#fffbea;border-left:4px solid #3a86ff;padding:12px 14px;border-radius:8px;margin:8px 0}"
        ".footer{margin-top:32px;text-align:center;color:#6b7280}"
        "@media (max-width: 640px){ h1{font-size:22px} h2{font-size:18px} .calendar th,.calendar td{padding:6px} .ev-item{font-size:11px} body{margin:20px auto} }"
        "a{color:#3a86ff;text-decoration:none} a:hover{text-decoration:underline}"
        "</style>"
        "</head>"
        "<body>"
        "<header>"
        "<h1>🎓 Расписание УЦП-24 РАНХиГС на 3 семестр</h1>"
        "</header>"

        "<section class=\"section\">"
        "<h2>Просмотр календаря</h2>"
        "<div class=\"card\">"
        f"<div class=cal-head><h3 style=\"margin:0\">{month_name.title()}</h3><div class=cal-nav>" +
        (f"<a href=\"{prev_href}\">← Пред</a>" if prev_allowed else "<a class=\\\"disabled\\\" href=\\\"#\\\">← Пред</a>") +
        (f"<a href=\"{next_href}\">След →</a>" if next_allowed else "<a class=\\\"disabled\\\" href=\\\"#\\\">След →</a>") +
        "</div></div>"
        "<table class=\"calendar\">"
        "<thead><tr>" + "".join(f"<th>{d}</th>" for d in weekdays) + "</tr></thead>"
        f"<tbody>{cal_html2}</tbody>"
        "</table>"
        "</div>"
        "</section>"

        "<section class=\"section\">"
        "<h2>Импорт в свой календарь</h2>"
        "<div class=\"card\">"
        "<a class=\"btn download\" href=\"/schedule.ics\">🗓️ Скачать календарь</a>"
        "</div>"
        "<div class=\"card\">"
        "<strong>Инструкция для установки в Google календарь</strong>"
        "<div class=\"hint\">Вам нужен Шаг 2</div>"
        "<a class=\"btn secondary\" href=\"https://support.google.com/calendar/answer/37118?hl=ru&co=GENIE.Platform%3DDesktop\" target=\"_blank\" rel=\"noopener\">Открыть инструкцию</a>"
        "</div>"
        "<div class=\"card\">"
        "<strong>Инструкция для установки в Яндекс календарь</strong>"
        "<div class=\"hint\">Вам нужен раздел <strong>Импорт из файла</strong></div>"
        "<a class=\"btn secondary\" href=\"https://yandex.ru/support/yandex-360/business/calendar/ru/create#import-iz-fajla\" target=\"_blank\" rel=\"noopener\">Открыть инструкцию</a>"
        "</div>"
        "<div class=\"card callout\"><strong>Совет:</strong> добавляйте данные события в отдельный календарь, чтобы иметь возможность отфильтровать их от других. Таблица могла некорректно распарситься и намного легче удалить целиком календарь, чем события в одиночку.</div>"

        "<div class=\"footer\">Создал Андрей Байков 2025 • <a href=\"https://github.com/andrbabai/MDPscheduler\" target=\"_blank\" rel=\"noopener\">Как это было сделано</a></div>"
        "</body></html>"
    )
