"""FastAPI server serving schedule ICS."""
from __future__ import annotations

import os
from pathlib import Path

from fastapi import FastAPI, Response

from .builder import build_ics, read_config

CONFIG_PATH = Path(os.environ.get("SCHEDICS_CONFIG", "config.yml"))

app = FastAPI()
_cache: dict[str, bytes | None] = {"ics": None}


def load_cache() -> None:
    cfg = read_config(CONFIG_PATH)
    _cache["ics"] = build_ics(cfg)


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
