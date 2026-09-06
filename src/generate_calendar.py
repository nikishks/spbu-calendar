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
CONFIG = json.loads(
    (ROOT / "config.json").read_text(encoding="utf-8")
)

API_EXCEL = "https://timetable.spbu.ru/StudentGroupEvents/ExcelWeek"
RU_CULTURE_COOKIE = "_culture"

FACULTATIVE_PREFIXES = (
    "факультатив.",
    "факультатив ",
)

ELECTIVE_PREFIX = "электив."


def norm(value):
    """
    Нормализация текста:
    - приводит к строке;
    - убирает неразрывные пробелы;
    - убирает лишние пробелы;
    - приводит к нижнему регистру.
    """
    value = str(value or "")
    value = value.replace("\xa0", " ")
    value = value.strip().lower()
    value = re.sub(r"\s+", " ", value)
    return value


def is_facultative(subject):
    """
    Определяет факультатив.
    Например:
    'Факультатив. ...'
    'Факультатив ...'
    """
    s = norm(subject)

    return any(
        s.startswith(prefix)
        for prefix in FACULTATIVE_PREFIXES
    )


def is_selected_elective(subject):
    """
    Оставляет только выбранные пользователем элективы.

    Сейчас выбран:
    - Теория игр

    Поэтому:
    'Электив. Теория игр' -> True
    'Электив. Управление конфликтами' -> False
    """

    s = norm(subject)

    if not s.startswith(ELECTIVE_PREFIX):
        return False

    elective_name = s[len(ELECTIVE_PREFIX):].strip()

    selected = CONFIG.get("selected_electives", [])

    for item in selected:
        selected_name = norm(item)

        if elective_name == selected_name:
            return True

    return False


def selected(subject):
    """
    Главный фильтр расписания.

    Оставляем:
    1. Все обычные предметы.
    2. Выбранные элективы.

    Убираем:
    1. Все остальные элективы.
    2. Все факультативы.
    """

    s = norm(subject)

    if not s:
        return False

    # Факультативы полностью исключаем.
    if is_facultative(s):
        return False

    # Если это электив —
    # оставляем только выбранный.
    if s.startswith(ELECTIVE_PREFIX):
        return is_selected_elective(s)

    # Всё остальное оставляем.
    return True


def get_sheet(group_id, monday):
    session = requests.Session()

    session.cookies.set(
        RU_CULTURE_COOKIE,
        "ru",
        domain="timetable.spbu.ru"
    )

    response = session.get(
        API_EXCEL,
        params={
            "studentGroupId": group_id,
            "weekMonday": monday.isoformat()
        },
        timeout=60,
    )

    response.raise_for_status()

    workbook = openpyxl.load_workbook(
        BytesIO(response.content),
        data_only=True
    )

    sheet_name = "Расписание студенческой группы"

    if sheet_name not in workbook.sheetnames:
        raise RuntimeError(
            "СПбГУ не вернул лист "
            "'Расписание студенческой группы'. "
            f"Получены листы: {workbook.sheetnames}"
        )

    return workbook[sheet_name]


def parse_sheet(sheet):
    rows = []

    for row in sheet.iter_rows(
        min_row=5,
        max_row=300,
        max_col=5
    ):
        values = []

        for cell in row:
            value = cell.value

            if value is not None:
                value = str(value)
                value = value.replace("\n", " ")
                value = value.strip()

            values.append(value)

        if any(values):
            rows.append(values)

    last_day = None
    result = []

    for row in rows:

        # Первый столбец содержит дату.
        # В следующих строках дата может быть пустой,
        # поэтому сохраняем последнюю найденную дату.
        if row[0]:
            last_day = row[0]

        if len(row) < 5:
            continue

        if not row[1] or not row[2]:
            continue

        if not last_day:
            continue

        # Ищем дату вида:
        # 08.09.2026
        date_match = re.search(
            r"(\d{2}\.\d{2}\.\d{4})",
            last_day
        )

        if not date_match:
            continue

        try:
            event_date = datetime.strptime(
                date_match.group(1),
                "%d.%m.%Y"
            ).date()
        except ValueError:
            continue

        # Время.
        time_string = str(row[1])

        time_string = (
            time_string
            .replace("—", "–")
            .replace("-", "–")
        )

        parts = [
            part.strip()
            for part in time_string.split("–")
        ]

        if len(parts) != 2:
            continue

        try:
            start_time = datetime.strptime(
                parts[0],
                "%H:%M"
            ).time()

            end_time = datetime.strptime(
                parts[1],
                "%H:%M"
            ).time()

        except ValueError:
            continue

        subject = row[2] or ""
        place = row[3] or ""
        lecturer = row[4] or ""

        # Главный фильтр.
        if not selected(subject):
            continue

        result.append(
            (
                event_date,
                start_time,
                end_time,
                subject,
                place,
                lecturer
            )
        )

    return result


def monday_on_or_before(current_date):
    return current_date - timedelta(
        days=current_date.weekday()
    )


def load_schedule():

    group_id = int(
        CONFIG["student_group_id"]
    )

    if group_id <= 0:
        raise SystemExit(
            "Укажи student_group_id в config.json."
        )

    start_date = date.fromisoformat(
        CONFIG["start_date"]
    )

    end_date = start_date + timedelta(
        days=30 * int(
            CONFIG.get("months_ahead", 8)
        )
    )

    monday = monday_on_or_before(
        start_date
    )

    all_events = []

    while monday <= end_date:

        print(
            f"Загрузка недели: {monday}"
        )

        sheet = get_sheet(
            group_id,
            monday
        )

        week_events = parse_sheet(sheet)

        print(
            f"Найдено подходящих занятий: "
            f"{len(week_events)}"
        )

        all_events.extend(
            week_events
        )

        monday += timedelta(days=7)

    # Убираем дубликаты.
    unique_events = {}

    for event in all_events:
        unique_events[event] = event

    events = sorted(
        unique_events.values(),
        key=lambda event: (
            event[0],
            event[1],
            event[3]
        )
    )

    return events


def make_ics(events):

    timezone_name = CONFIG.get(
        "timezone",
        "Europe/Moscow"
    )

    timezone = ZoneInfo(
        timezone_name
    )

    calendar = Calendar()

    calendar.add(
        "prodid",
        "-//SPbU Apple Calendar//RU//"
    )

    calendar.add(
        "version",
        "2.0"
    )

    calendar.add(
        "X-WR-CALNAME",
        CONFIG.get(
            "calendar_name",
            "СПбГУ"
        )
    )

    calendar.add(
        "X-WR-CALDESC",
        CONFIG.get(
            "calendar_description",
            "Расписание СПбГУ"
        )
    )

    calendar.add(
        "X-WR-TIMEZONE",
        timezone_name
    )

    calendar.add(
        "CALSCALE",
        "GREGORIAN"
    )

    calendar.add(
        "METHOD",
        "PUBLISH"
    )

    for (
        event_date,
        start_time,
        end_time,
        subject,
        place,
        lecturer
    ) in events:

        event = Event()

        start_datetime = datetime.combine(
            event_date,
            start_time,
            tzinfo=timezone
        )

        end_datetime = datetime.combine(
            event_date,
            end_time,
            tzinfo=timezone
        )

        uid_base = (
            f"{event_date.isoformat()}-"
            f"{start_time}-"
            f"{end_time}-"
            f"{subject}-"
            f"{place}-"
            f"{lecturer}"
        )

        uid = (
            re.sub(
                r"[^0-9A-Za-zА-Яа-я_.-]+",
                "-",
                uid_base
            ).strip("-")
            + "@spbu-apple-calendar"
        )

        event.add(
            "uid",
            uid
        )

        event.add(
            "dtstamp",
            datetime.now(timezone)
        )

        event.add(
            "dtstart",
            start_datetime
        )

        event.add(
            "dtend",
            end_datetime
        )

        event.add(
            "summary",
            subject
        )

        description = []

        if lecturer:
            description.append(
                f"Преподаватель: {lecturer}"
            )

        if place:
            description.append(
                f"Аудитория: {place}"
            )

        description.append(
            "Источник: timetable.spbu.ru"
        )

        event.add(
            "description",
            "\n".join(description)
        )

        if place:
            event.add(
                "location",
                place
            )

        calendar.add_component(event)

    return calendar.to_ical()


def main():

    events = load_schedule()

    output = (
        ROOT
        / "docs"
        / "schedule.ics"
    )

    output.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    output.write_bytes(
        make_ics(events)
    )

    print(
        f"Создано событий: {len(events)}"
    )

    print(
        f"Файл: {output}"
    )


if __name__ == "__main__":
    main()
