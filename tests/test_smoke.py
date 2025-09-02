import subprocess
from pathlib import Path

from fastapi.testclient import TestClient

from schedics.server import app, load_cache


def test_cli_build(tmp_path: Path) -> None:
    out = tmp_path / "schedule.ics"
    subprocess.check_call([
        "python",
        "-m",
        "schedics.cli",
        "build",
        "--config",
        "config.yml",
        "--output",
        str(out),
    ])
    text = out.read_text()
    assert "BEGIN:VCALENDAR" in text
    assert "END:VCALENDAR" in text


def test_server_endpoints() -> None:
    load_cache()
    client = TestClient(app)
    resp = client.get("/schedule.ics")
    assert resp.status_code == 200
    assert resp.headers["content-type"].startswith("text/calendar")
    resp2 = client.post("/refresh")
    assert resp2.status_code == 200
