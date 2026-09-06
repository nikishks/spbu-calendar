import json
import re
from datetime import datetime, date, timedelta
from io import BytesIO
from pathlib import Path
from zoneinfo import ZoneInfo

import requests
import openpyxl
from icalendar import Calendar, Event

ROOT = Path(__file__).resolve().parents[1]
CONFIG = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))

API_EXCEL = "https://timetable.spbu.ru/StudentGroupEvents/ExcelWeek"
RU_CULTURE_COOKIE = "_culture"

DAYS = {
    "понедельник": 0,
    "вторник": 1,
    "среда": 2,
    "четверг": 3,
    "пятница": 4,
    "суббота": 5,
    "воскресенье": 6,
}

def norm(s):
    s = str(s or "").replace("\xa0", " ").strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s

def selected(subject):
    subjects = CONFIG.get("subjects", [])
    if not subjects:
        return True
    s = norm(subject)
    return any(norm(x) in s or s in norm(x) for x in subjects)

def get_sheet(group_id, monday):
    session = requests.Session()
    session.cookies.set(RU_CULTURE_COOKIE, "ru", domain="timetable.spbu.ru")
    r = session.get(
        API_EXCEL,
        params={"studentGroupId": group_id, "weekMonday": monday.isoformat()},
        timeout=60,
    )
    r.raise_for_status()
    wb = openpyxl.load_workbook(BytesIO(r.content), data_only=True)
    if "Расписание студенческой группы" not in wb.sheetnames:
        raise RuntimeError(
            "СПбГУ не вернул лист 'Расписание студенческой группы'. "
            f"Получены листы: {wb.sheetnames}"
        )
    return wb["Расписание студенческой группы"]

def parse_sheet(sheet):
    rows = []
    for row in sheet.iter_rows(min_row=5, max_row=300, max_col=5):
        values = []
        for c in row:
            v = c.value
            values.append(str(v).replace("\n", " ").strip() if v is not None else None)
        if any(values):
            rows.append(values)

    last_day = None
    result = []

    for row in rows:
        if row[0]:
            last_day = row[0]

        if len(row) < 5 or not row[1] or not row[2]:
            continue
        if not last_day:
            continue

        # Example: "Понедельник 07.09.2026"
        m = re.search(r"(\d{2}\.\d{2}\.\d{4})", last_day)
        if not m:
            continue

        try:
            d = datetime.strptime(m.group(1), "%d.%m.%Y").date()
        except ValueError:
            continue

        tm = str(row[1]).replace("—", "–").replace("-", "–")
        parts = [x.strip() for x in tm.split("–")]
        if len(parts) != 2:
            continue

        try:
            start = datetime.strptime(parts[0], "%H:%M").time()
            end = datetime.strptime(parts[1], "%H:%M").time()
        except ValueError:
            continue

        subject = row[2] or ""
        place = row[3] or ""
        lecturer = row[4] or ""

        if selected(subject):
            result.append((d, start, end, subject, place, lecturer))

    return result

def monday_on_or_before(d):
    return d - timedelta(days=d.weekday())

def load_schedule():
    gid = int(CONFIG["student_group_id"])
    if gid <= 0:
        raise SystemExit(
            "Укажи student_group_id в config.json. "
            "Это внутренний ID группы СПбГУ."
        )

    start = date.fromisoformat(CONFIG["start_date"])
    end = start + timedelta(days=30 * int(CONFIG.get("months_ahead", 8)))
    monday = monday_on_or_before(start)

    all_events = []
    while monday <= end:
        sheet = get_sheet(gid, monday)
        week_events = parse_sheet(sheet)
        all_events.extend(week_events)
        monday += timedelta(days=7)

    # Дедупликация
    unique = {}
    for item in all_events:
        unique[item] = item
    return sorted(unique.values(), key=lambda x: (x[0], x[1], x[3]))

def make_ics(events):
    tz_name = CONFIG.get("timezone", "Europe/Moscow")
    tz = ZoneInfo(tz_name)

    cal = Calendar()
    cal.add("prodid", "-//SPbU Apple Calendar//RU//")
    cal.add("version", "2.0")
    cal.add("X-WR-CALNAME", CONFIG.get("calendar_name", "СПбГУ"))
    cal.add("X-WR-CALDESC", CONFIG.get("calendar_description", "Расписание СПбГУ"))
    cal.add("X-WR-TIMEZONE", tz_name)
    cal.add("CALSCALE", "GREGORIAN")
    cal.add("METHOD", "PUBLISH")

    for d, start, end, subject, place, lecturer in events:
        ev = Event()
        start_dt = datetime.combine(d, start, tzinfo=tz)
        end_dt = datetime.combine(d, end, tzinfo=tz)

        uid_base = f"{d.isoformat()}-{start}-{end}-{subject}-{place}-{lecturer}"
        uid = re.sub(r"[^0-9A-Za-zА-Яа-я_.-]+", "-", uid_base).strip("-") + "@spbu-apple-calendar"

        ev.add("uid", uid)
        ev.add("dtstamp", datetime.now(tz))
        ev.add("dtstart", start_dt)
        ev.add("dtend", end_dt)
        ev.add("summary", subject)

        description = []
        if lecturer:
            description.append(f"Преподаватель: {lecturer}")
        if place:
            description.append(f"Аудитория: {place}")
        description.append("Источник: timetable.spbu.ru")
        ev.add("description", "\n".join(description))

        if place:
            ev.add("location", place)

        cal.add_component(ev)

    return cal.to_ical()

def main():
    events = load_schedule()
    out = ROOT / "docs" / "schedule.ics"
    out.write_bytes(make_ics(events))
    print(f"Создано событий: {len(events)}")
    print(f"Файл: {out}")

if __name__ == "__main__":
    main()
