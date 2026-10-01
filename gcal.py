"""Google Calendar access: OAuth, calendar lookup, idempotent upsert, cleanup."""
from __future__ import annotations

import datetime as dt
import os
import time

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from parse_plan import Event

SCOPES = ["https://www.googleapis.com/auth/calendar"]
HERE = os.path.dirname(os.path.abspath(__file__))


def get_service():
    token_path = os.path.join(HERE, "token.json")
    creds_path = os.path.join(HERE, "credentials.json")
    creds = None
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(creds_path):
                raise SystemExit("credentials.json not found – see README.md (Google Cloud setup).")
            creds = InstalledAppFlow.from_client_secrets_file(creds_path, SCOPES).run_local_server(port=0)
        with open(token_path, "w") as f:
            f.write(creds.to_json())
    return build("calendar", "v3", credentials=creds)


def resolve_calendar(service, cfg: dict, create: bool = True) -> str:
    name = (cfg.get("calendar_name") or "").strip()
    if not name:
        return cfg.get("calendar_id", "primary")
    page = None
    while True:
        resp = service.calendarList().list(pageToken=page).execute()
        for cal in resp.get("items", []):
            if cal.get("summary") == name:
                return cal["id"]
        page = resp.get("nextPageToken")
        if not page:
            break
    if not create:
        raise SystemExit(f"Calendar '{name}' not found.")
    cal = service.calendars().insert(body={"summary": name, "timeZone": cfg["timezone"]}).execute()
    print(f"Created calendar '{name}'.")
    return cal["id"]


def to_body(e: Event, cfg: dict) -> dict:
    tz = cfg["timezone"]
    desc = []
    if e.code != "__notice__":
        desc.append(f"Kod: {e.code}")
    if e.lecturer:
        desc.append(f"Prowadzący / temat: {e.lecturer}")
    if e.room:
        desc.append(f"Sala: {e.room}")
    desc += e.notes
    desc.append(f"Źródło: {cfg['excel_file']}, wiersz {e.row}")

    body = {
        "id": e.event_id,
        "status": "confirmed",
        "summary": e.title,
        "description": "\n".join(desc),
        "extendedProperties": {"private": {"source": cfg["source_tag"]}},
    }
    if e.color_id:
        body["colorId"] = e.color_id
    if e.room:
        body["location"] = "on-line" if "on-line" in e.room.lower() else \
            cfg["location_template"].format(room=e.room)
    if e.start is None:  # all-day (end date is exclusive in the API)
        last = e.end_date or e.date
        body["start"] = {"date": e.date.isoformat()}
        body["end"] = {"date": (last + dt.timedelta(days=1)).isoformat()}
    else:
        body["start"] = {"dateTime": dt.datetime.combine(e.date, e.start).isoformat(), "timeZone": tz}
        body["end"] = {"dateTime": dt.datetime.combine(e.date, e.end).isoformat(), "timeZone": tz}
    return body


def _retry(fn, tries=6):
    for i in range(tries):
        try:
            return fn()
        except HttpError as err:
            if err.resp.status in (403, 429, 500, 503) and i < tries - 1:
                time.sleep(2 ** i)
                continue
            raise


def upsert(service, cal_id: str, events: list[Event], cfg: dict):
    created = updated = 0
    for e in events:
        body = to_body(e, cfg)
        try:
            _retry(lambda: service.events().insert(calendarId=cal_id, body=body).execute())
            created += 1
        except HttpError as err:
            if err.resp.status == 409:  # already exists (also if previously deleted) -> update
                _retry(lambda: service.events().update(
                    calendarId=cal_id, eventId=e.event_id, body=body).execute())
                updated += 1
            else:
                raise
    print(f"Done: {created} created, {updated} updated.")


def list_tagged(service, cal_id: str, tag: str) -> list[dict]:
    out, page = [], None
    while True:
        resp = _retry(lambda: service.events().list(
            calendarId=cal_id, privateExtendedProperty=f"source={tag}",
            pageToken=page, maxResults=2500, showDeleted=False).execute())
        out += resp.get("items", [])
        page = resp.get("nextPageToken")
        if not page:
            return out


def delete_ids(service, cal_id: str, ids: list[str]):
    for eid in ids:
        _retry(lambda: service.events().delete(calendarId=cal_id, eventId=eid).execute())
    print(f"Deleted {len(ids)} event(s).")
