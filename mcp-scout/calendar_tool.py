# © 2026 Iryna Subbotina. All Rights Reserved. Proprietary — not open source. See LICENSE/NOTICE. Unauthorized copying, redistribution, or claim of authorship is prohibited.
"""Outlook Calendar integration (Microsoft Graph) — free slots and upcoming events."""
from __future__ import annotations

import logging
import re
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo

import gauth

log = logging.getLogger(__name__)

TZ = ZoneInfo("Europe/Warsaw")
UTC = timezone.utc
TZ_NAME = "Europe/Warsaw"

_FREE_STATUSES = {"free"}

_my_email_cache: str | None = None


def _my_email() -> str:
    global _my_email_cache
    if _my_email_cache is None:
        me = gauth.request("GET", "/me", params={"$select": "mail,userPrincipalName"})
        _my_email_cache = me.get("mail") or me.get("userPrincipalName") or ""
    return _my_email_cache


def _parse_graph_dt(s: str, tz=TZ) -> datetime:
    # Graph returns up to 7 fractional-second digits; datetime.fromisoformat wants <=6.
    s = re.sub(r"(\.\d{6})\d*$", r"\1", s.rstrip("Z"))
    dt = datetime.fromisoformat(s)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=tz)
    return dt


def get_free_slots(
    days: int = 3,
    slot_duration_minutes: int = 45,
    work_start_hour: int = 9,
    work_end_hour: int = 18,
    calendar_id: str = "primary",
) -> dict:
    """Find free time slots in the Outlook Calendar for scheduling interviews.

    Args:
      days: how many days ahead to look (default 3)
      slot_duration_minutes: interview duration in minutes (default 45)
      work_start_hour: working day start hour in Warsaw time (default 9)
      work_end_hour: working day end hour in Warsaw time (default 18)
      calendar_id: unused, kept for backward compatibility (always checks the signed-in user's calendar)

    Returns: {slots: [{start, end, label}], timezone, generated_at}
    """
    now = datetime.now(TZ)
    time_min = now.isoformat()
    time_max = (now + timedelta(days=days)).isoformat()

    body = {
        "schedules": [_my_email()],
        "startTime": {"dateTime": time_min, "timeZone": TZ_NAME},
        "endTime": {"dateTime": time_max, "timeZone": TZ_NAME},
        "availabilityViewInterval": 30,
    }
    result = gauth.request("POST", "/me/calendar/getSchedule", json=body)
    schedules = result.get("value", [])
    schedule_items = schedules[0].get("scheduleItems", []) if schedules else []

    busy = []
    for item in schedule_items:
        if item.get("status") in _FREE_STATUSES:
            continue
        start = _parse_graph_dt(item["start"]["dateTime"])
        end = _parse_graph_dt(item["end"]["dateTime"])
        busy.append((start, end))

    free_slots = []
    slot_delta = timedelta(minutes=slot_duration_minutes)

    for day_offset in range(days):
        day = (now + timedelta(days=day_offset)).date()
        if datetime(day.year, day.month, day.day).weekday() >= 5:
            continue

        work_start = datetime(day.year, day.month, day.day, work_start_hour, 0, tzinfo=TZ)
        work_end = datetime(day.year, day.month, day.day, work_end_hour, 0, tzinfo=TZ)

        cursor = max(work_start, now + timedelta(minutes=30))
        while cursor + slot_delta <= work_end:
            slot_end = cursor + slot_delta
            overlap = any(
                cursor < b_end and slot_end > b_start
                for b_start, b_end in busy
            )
            if not overlap:
                free_slots.append({
                    "start": cursor.strftime("%Y-%m-%dT%H:%M"),
                    "end": slot_end.strftime("%Y-%m-%dT%H:%M"),
                    "label": f"{cursor.strftime('%a %d.%m')} {cursor.strftime('%H:%M')}–{slot_end.strftime('%H:%M')}",
                })
            cursor += timedelta(minutes=30)

    log.info("calendar: found %d free slots over %d days", len(free_slots), days)
    return {
        "slots": free_slots[:20],
        "timezone": TZ_NAME,
        "slot_duration_minutes": slot_duration_minutes,
        "generated_at": now.strftime("%d.%m.%Y %H:%M"),
    }


def get_upcoming_events(
    hours: int = 48,
    calendar_id: str = "primary",
    max_results: int = 20,
) -> list[dict]:
    """Get upcoming calendar events (interviews, meetings).

    Args:
      hours: how many hours ahead to look (default 48)
      calendar_id: unused, kept for backward compatibility (always checks the signed-in user's calendar)
      max_results: max events to return (default 20)

    Returns: list of {title, start, end, start_iso, location, description, event_id}
    """
    now = datetime.now(UTC)
    time_min = now.strftime("%Y-%m-%dT%H:%M:%SZ")
    time_max = (now + timedelta(hours=hours)).strftime("%Y-%m-%dT%H:%M:%SZ")

    result = gauth.request(
        "GET",
        "/me/calendarView",
        params={
            "startDateTime": time_min,
            "endDateTime": time_max,
            "$orderby": "start/dateTime",
            "$top": max_results,
            "$select": "id,subject,start,end,location,bodyPreview",
        },
        headers={"Prefer": f'outlook.timezone="{TZ_NAME}"'},
    )

    events = []
    for e in result.get("value", []):
        try:
            start_dt = _parse_graph_dt(e["start"]["dateTime"])
            end_dt = _parse_graph_dt(e["end"]["dateTime"])
            start_label = start_dt.strftime("%a %d.%m %H:%M")
            end_label = end_dt.strftime("%H:%M")
            start_iso = start_dt.isoformat()
        except Exception:
            start_label = e.get("start", {}).get("dateTime", "")
            end_label = e.get("end", {}).get("dateTime", "")
            start_iso = start_label

        events.append({
            "event_id": e.get("id", ""),
            "title": e.get("subject") or "(без назви)",
            "start": start_label,
            "end": end_label,
            "start_iso": start_iso,
            "location": (e.get("location") or {}).get("displayName", ""),
            "description": (e.get("bodyPreview") or "")[:200],
        })

    log.info("calendar: %d upcoming events in next %dh", len(events), hours)
    return events


def get_events_starting_soon(minutes: int = 15) -> list[dict]:
    """Get events starting within the next N minutes. Used for reminders."""
    now = datetime.now(UTC)
    time_min = now.strftime("%Y-%m-%dT%H:%M:%SZ")
    time_max = (now + timedelta(minutes=minutes)).strftime("%Y-%m-%dT%H:%M:%SZ")

    result = gauth.request(
        "GET",
        "/me/calendarView",
        params={
            "startDateTime": time_min,
            "endDateTime": time_max,
            "$orderby": "start/dateTime",
            "$top": 5,
            "$select": "id,subject,start,location",
        },
        headers={"Prefer": f'outlook.timezone="{TZ_NAME}"'},
    )

    events = []
    for e in result.get("value", []):
        start_raw = (e.get("start") or {}).get("dateTime", "")
        if not start_raw:
            continue
        try:
            start_dt = _parse_graph_dt(start_raw)
            minutes_left = int((start_dt.astimezone(UTC) - now).total_seconds() / 60)
            events.append({
                "event_id": e.get("id", ""),
                "title": e.get("subject") or "(без назви)",
                "start": start_dt.strftime("%H:%M"),
                "minutes_left": minutes_left,
                "location": (e.get("location") or {}).get("displayName", ""),
            })
        except Exception:
            continue

    return events
