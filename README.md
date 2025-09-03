# MDPscheduler

Generate and serve an `.ics` feed from the R7 Office (Яндекс.Диск) schedule.

## Quickstart

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python -m schedics.cli build --config config.yml --output public/schedule.ics
```

The resulting calendar is written to `public/schedule.ics`.

## CLI

```
python -m schedics.cli build --config config.yml --output public/schedule.ics
python -m schedics.cli print --config config.yml
```

## FastAPI server

```
uvicorn schedics.server:app --reload --port 8080
```

- `GET /schedule.ics` – serve cached calendar
- `POST /refresh` – rebuild and update cache

## VSCode

Launch and tasks files are provided in `.vscode/` to build the feed or run the server directly from VSCode.

## Docker

Build and run with Docker:

```
docker build -t schedics .
docker run -p 8080:8080 schedics
```

or using docker-compose:

```
docker-compose up --build
```

## Cron / refresh

To refresh the calendar periodically, add a cron entry:

```
0 6 * * * /usr/bin/python -m schedics.cli build --config /path/to/config.yml --output /var/www/public/schedule.ics
```

## Calendar subscription

The generated `schedule.ics` file can be served over HTTP and subscribed to from calendar clients (Google Calendar, Apple Calendar, etc.).
