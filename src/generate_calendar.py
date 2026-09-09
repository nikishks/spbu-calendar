import json
import re
from datetime import date, datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import requests
from icalendar import Calendar, Event


# ============================================================
# CONFIG
# ============================================================

ROOT = Path(__file__).resolve().parents[1]

CONFIG = json.loads(
    (ROOT / "config.json").read_text(encoding="utf-8")
)

API_BASE = "https://timetable.spbu.ru/api/v1"


# ============================================================
# HELPERS
# ============================================================

def norm(value):
    """
    Нормализация текста для сравнений.
    """
    value = str(value or "")
    value = value.replace("\xa0", " ")
    value = value.strip().lower()
    value = re.sub(r"\s+", " ", value)
    return value


def get_value(obj, *names, default=None):
    """
    Получить значение независимо от регистра ключа.

    Например:
        IsCancelled
        isCancelled
        ISCancelled

    будут найдены одинаково.
    """
    if not isinstance(obj, dict):
        return default

    for name in names:
        if name in obj:
            return obj[name]

    lowered = {
        str(key).lower(): value
        for key, value in obj.items()
    }

    for name in names:
        key = name.lower()

        if key in lowered:
            return lowered[key]

    return default


# ============================================================
# SUBJECT FILTER
# ============================================================

def is_selected_subject(subject, api_event=None):
    """
    Обычные предметы добавляем.

    Факультативы не добавляем.

    Если событие является элективом, добавляем только
    выбранные элективы из config.json.
    """

    subject_norm = norm(subject)

    if not subject_norm:
        return False

    # --------------------------------------------
    # ФАКУЛЬТАТИВЫ
    # --------------------------------------------

    if subject_norm.startswith("факультатив"):
        return False

    # --------------------------------------------
    # ЭЛЕКТИВ
    # --------------------------------------------

    is_elective = False

    if isinstance(api_event, dict):
        is_elective = bool(
            get_value(
                api_event,
                "IsElective",
                "isElective",
                default=False,
            )
        )

    # На случай если API почему-то не выставил IsElective,
    # сохраняем поддержку старого формата названий.
    if subject_norm.startswith("электив"):
        is_elective = True

    if is_elective:
        selected = [
            norm(x)
            for x in CONFIG.get(
                "selected_electives",
                []
            )
        ]

        for elective in selected:
            if not elective:
                continue

            # Поддерживаем оба варианта:
            #
            # config: "Теория игр"
            # subject: "Теория игр"
            #
            # или:
            #
            # subject: "Электив. Теория игр"
            if (
                elective in subject_norm
                or subject_norm in elective
            ):
                return True

        return False

    return True


# ============================================================
# DATE PARSING
# ============================================================

def parse_api_datetime(value):
    """
    API СПбГУ в разные периоды мог возвращать дату
    в разных форматах.

    Поддерживаем:
        2026-09-12T10:00:00
        2026-09-12T10:00:00+03:00
        2026-09-12T10:00:00Z
        /Date(1789196400000)/
    """

    if value is None:
        return None

    # Уже datetime
    if isinstance(value, datetime):
        return value

    text = str(value).strip()

    if not text:
        return None

    # --------------------------------------------
    # Microsoft JSON date:
    # /Date(1789196400000)/
    # --------------------------------------------

    match = re.search(
        r"/Date\((-?\d+)",
        text
    )

    if match:
        try:
            timestamp_ms = int(
                match.group(1)
            )

            return datetime.fromtimestamp(
                timestamp_ms / 1000,
                tz=ZoneInfo(
                    CONFIG.get(
                        "timezone",
                        "Europe/Moscow"
                    )
                ),
            )

        except (
            ValueError,
            OverflowError
        ):
            return None

    # --------------------------------------------
    # ISO
    # --------------------------------------------

    iso_text = text

    if iso_text.endswith("Z"):
        iso_text = (
            iso_text[:-1]
            + "+00:00"
        )

    try:
        return datetime.fromisoformat(
            iso_text
        )
    except ValueError:
        pass

    # --------------------------------------------
    # Дополнительные форматы
    # --------------------------------------------

    formats = [
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%dT%H:%M",
        "%Y-%m-%d %H:%M",
        "%d.%m.%Y %H:%M",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(
                text,
                fmt
            )
        except ValueError:
            continue

    return None


def localize_datetime(value):
    """
    Привести datetime к timezone из config.json.
    """

    if value is None:
        return None

    timezone = ZoneInfo(
        CONFIG.get(
            "timezone",
            "Europe/Moscow"
        )
    )

    if value.tzinfo is None:
        return value.replace(
            tzinfo=timezone
        )

    return value.astimezone(
        timezone
    )


# ============================================================
# API EVENT DISCOVERY
# ============================================================

def looks_like_event(obj):
    """
    API может вернуть события внутри:
        Days
        DayStudyEvents
        Events
        и т.д.

    Поэтому не привязываемся жёстко к оболочке JSON.

    Событием считаем словарь, содержащий как минимум:
        Subject
        Start
        End
    """

    if not isinstance(obj, dict):
        return False

    subject = get_value(
        obj,
        "Subject"
    )

    start = get_value(
        obj,
        "Start"
    )

    end = get_value(
        obj,
        "End"
    )

    return (
        subject is not None
        and start is not None
        and end is not None
    )


def find_events_recursive(data):
    """
    Рекурсивно найти все события в JSON API.

    Это делает код устойчивым к изменению внешней
    структуры ответа API.
    """

    found = []

    if isinstance(data, dict):

        if looks_like_event(data):
            found.append(data)

        for value in data.values():
            found.extend(
                find_events_recursive(
                    value
                )
            )

    elif isinstance(data, list):

        for item in data:
            found.extend(
                find_events_recursive(
                    item
                )
            )

    return found


# ============================================================
# LOCATION / EDUCATOR
# ============================================================

def extract_location(event):
    """
    Основной текст аудитории API обычно уже отдаёт
    в LocationsDisplayText.
    """

    value = get_value(
        event,
        "LocationsDisplayText",
        "locationsDisplayText",
    )

    if value:
        return str(value).strip()

    # Fallback на EventLocations
    locations = get_value(
        event,
        "EventLocations",
        "eventLocations",
        default=[],
    )

    result = []

    if isinstance(locations, list):

        for location in locations:

            if not isinstance(
                location,
                dict
            ):
                continue

            # Пробуем распространённые поля.
            pieces = []

            for key in (
                "DisplayName",
                "LocationDisplayName",
                "Room",
                "RoomName",
                "Address",
            ):
                value = get_value(
                    location,
                    key
                )

                if (
                    value
                    and str(value).strip()
                    not in pieces
                ):
                    pieces.append(
                        str(value).strip()
                    )

            if pieces:
                result.append(
                    ", ".join(pieces)
                )

    return "; ".join(result)


def extract_educator(event):
    """
    Имя преподавателя.
    """

    value = get_value(
        event,
        "EducatorsDisplayText",
        "educatorsDisplayText",
    )

    if value:
        return str(value).strip()

    educators = get_value(
        event,
        "EducatorIds",
        "Educators",
        "educators",
        default=[],
    )

    result = []

    if isinstance(educators, list):

        for educator in educators:

            if not isinstance(
                educator,
                dict
            ):
                continue

            name = get_value(
                educator,
                "DisplayName",
                "FullName",
                "Name",
            )

            if name:
                result.append(
                    str(name).strip()
                )

    return ", ".join(result)


# ============================================================
# API
# ============================================================

def make_api_date(dt):
    """
    Формат API:
        YYYYMMDDHHMM
    """

    return dt.strftime(
        "%Y%m%d%H%M"
    )


def get_api_events(
    group_id,
    start_date,
    end_date
):
    """
    Получить события группы через официальный API СПбГУ.
    """

    timezone = ZoneInfo(
        CONFIG.get(
            "timezone",
            "Europe/Moscow"
        )
    )

    start_dt = datetime.combine(
        start_date,
        datetime.min.time(),
        tzinfo=timezone,
    )

    end_dt = datetime.combine(
        end_date,
        datetime.max.time(),
        tzinfo=timezone,
    )

    from_string = make_api_date(
        start_dt
    )

    to_string = make_api_date(
        end_dt
    )

    url = (
        f"{API_BASE}/groups/"
        f"{group_id}/events/"
        f"{from_string}/"
        f"{to_string}"
    )

    print()
    print(
        "Запрос API:"
    )
    print(
        f"  {start_date} — {end_date}"
    )

    response = requests.get(
        url,
        params={
            "timetable": "Primary"
        },
        headers={
            "Accept": "application/json",
            "User-Agent": (
                "spbu-calendar/2.0 "
                "(GitHub Actions)"
            ),
        },
        timeout=60,
    )

    response.raise_for_status()

    try:
        data = response.json()

    except ValueError as exc:
        print(
            "Ответ API не является JSON:"
        )

        print(
            response.text[:1000]
        )

        raise RuntimeError(
            "СПбГУ API вернул некорректный JSON"
        ) from exc

    events = find_events_recursive(
        data
    )

    print(
        f"API вернул событий: "
        f"{len(events)}"
    )

    return events


# ============================================================
# EVENT PARSING
# ============================================================

def parse_api_event(event):
    """
    Преобразовать событие СПбГУ во внутренний формат.

    Возвращает None, если событие:
        - отменено;
        - не подходит по фильтру;
        - имеет некорректные даты.
    """

    subject = str(
        get_value(
            event,
            "Subject",
            default="",
        )
        or ""
    ).strip()

    if not subject:
        return None

    # ========================================================
    # ГЛАВНАЯ ПРОВЕРКА:
    # ОТМЕНЁННОЕ ЗАНЯТИЕ
    # ========================================================

    is_cancelled = bool(
        get_value(
            event,
            "IsCancelled",
            "isCancelled",
            default=False,
        )
    )

    start_raw = get_value(
        event,
        "Start"
    )

    end_raw = get_value(
        event,
        "End"
    )

    start_dt = localize_datetime(
        parse_api_datetime(
            start_raw
        )
    )

    end_dt = localize_datetime(
        parse_api_datetime(
            end_raw
        )
    )

    if is_cancelled:

        if start_dt:
            event_time = (
                start_dt.strftime(
                    "%d.%m.%Y %H:%M"
                )
            )
        else:
            event_time = str(
                start_raw or ""
            )

        print(
            "ОТМЕНЕНО -> пропускаю:"
        )

        print(
            f"  {event_time} — "
            f"{subject}"
        )

        return None

    # ========================================================
    # SUBJECT FILTER
    # ========================================================

    if not is_selected_subject(
        subject,
        event
    ):
        return None

    # ========================================================
    # TIME
    # ========================================================

    if (
        start_dt is None
        or end_dt is None
    ):
        print(
            "Не удалось разобрать время:"
        )

        print(
            f"  {subject}"
        )

        print(
            f"  Start={start_raw}"
        )

        print(
            f"  End={end_raw}"
        )

        return None

    # ========================================================
    # ALL-DAY
    # ========================================================

    all_day = bool(
        get_value(
            event,
            "AllDay",
            "allDay",
            default=False,
        )
    )

    # Для нашего учебного календаря события без
    # конкретного времени пока пропускаем.
    if all_day:
        print(
            "All-day событие пропущено:"
        )

        print(
            f"  {subject}"
        )

        return None

    location = extract_location(
        event
    )

    educator = extract_educator(
        event
    )

    return {
        "start": start_dt,
        "end": end_dt,
        "subject": subject,
        "location": location,
        "educator": educator,

        "time_was_changed": bool(
            get_value(
                event,
                "TimeWasChanged",
                default=False,
            )
        ),

        "location_was_changed": bool(
            get_value(
                event,
                "LocationsWereChanged",
                default=False,
            )
        ),

        "educator_was_changed": bool(
            get_value(
                event,
                "EducatorsWereReassigned",
                default=False,
            )
        ),
    }


# ============================================================
# LOAD SCHEDULE
# ============================================================

def load_schedule():
    group_id = int(
        CONFIG[
            "student_group_id"
        ]
    )

    start_date = date.fromisoformat(
        CONFIG[
            "start_date"
        ]
    )

    months_ahead = int(
        CONFIG.get(
            "months_ahead",
            8
        )
    )

    # Как и раньше: приблизительно N месяцев.
    end_date = (
        start_date
        + timedelta(
            days=30 * months_ahead
        )
    )

    result = []

    # Запрашиваем по неделям.
    # Это надёжнее, чем один огромный API-запрос.
    current = start_date

    while current <= end_date:

        chunk_end = min(
            current
            + timedelta(days=6),
            end_date,
        )

        raw_events = get_api_events(
            group_id,
            current,
            chunk_end,
        )

        for raw_event in raw_events:

            parsed = parse_api_event(
                raw_event
            )

            if parsed is None:
                continue

            result.append(
                parsed
            )

        current = (
            chunk_end
            + timedelta(days=1)
        )

    # ========================================================
    # УДАЛЕНИЕ ДУБЛИКАТОВ
    # ========================================================

    unique = {}

    for event in result:

        key = (
            event["start"],
            event["end"],
            norm(
                event["subject"]
            ),
            norm(
                event["location"]
            ),
            norm(
                event["educator"]
            ),
        )

        unique[key] = event

    events = list(
        unique.values()
    )

    events.sort(
        key=lambda event: (
            event["start"],
            norm(
                event["subject"]
            ),
        )
    )

    return events


# ============================================================
# UID
# ============================================================

def make_uid(event):
    """
    Стабильный UID.

    Аудитория и преподаватель НЕ входят в UID:
    если они поменяются, Apple Calendar должен
    обновить существующее событие, а не создать новое.
    """

    start = event["start"]

    subject = norm(
        event["subject"]
    )

    base = (
        f"{CONFIG['student_group_id']}-"
        f"{start.strftime('%Y%m%d-%H%M')}-"
        f"{subject}"
    )

    safe = re.sub(
        r"[^0-9A-Za-zА-Яа-яЁё_.-]+",
        "-",
        base
    )

    safe = safe.strip("-")

    return (
        f"{safe}@spbu-calendar"
    )


# ============================================================
# ICS
# ============================================================

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
        "-//SPbU Calendar//RU//"
    )

    calendar.add(
        "version",
        "2.0"
    )

    calendar.add(
        "calscale",
        "GREGORIAN"
    )

    calendar.add(
        "method",
        "PUBLISH"
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

    generated_at = datetime.now(
        timezone
    )

    for item in events:

        event = Event()

        event.add(
            "uid",
            make_uid(item)
        )

        event.add(
            "dtstamp",
            generated_at
        )

        event.add(
            "last-modified",
            generated_at
        )

        event.add(
            "dtstart",
            item["start"]
        )

        event.add(
            "dtend",
            item["end"]
        )

        event.add(
            "summary",
            item["subject"]
        )

        if item["location"]:

            event.add(
                "location",
                item["location"]
            )

        description = []

        if item["educator"]:

            description.append(
                "Преподаватель: "
                + item["educator"]
            )

        if item["location"]:

            description.append(
                "Место: "
                + item["location"]
            )

        changes = []

        if item[
            "time_was_changed"
        ]:
            changes.append(
                "изменено время"
            )

        if item[
            "location_was_changed"
        ]:
            changes.append(
                "изменено место"
            )

        if item[
            "educator_was_changed"
        ]:
            changes.append(
                "изменён преподаватель"
            )

        if changes:

            description.append(
                "Изменения: "
                + ", ".join(changes)
            )

        description.append(
            "Источник: timetable.spbu.ru"
        )

        event.add(
            "description",
            "\n".join(
                description
            )
        )

        calendar.add_component(
            event
        )

    return calendar.to_ical()


# ============================================================
# MAIN
# ============================================================

def main():

    print(
        "=" * 60
    )

    print(
        "СПбГУ -> iCalendar"
    )

    print(
        "Источник: официальный API timetable.spbu.ru"
    )

    print(
        f"Группа: "
        f"{CONFIG.get('group_name', '')} "
        f"(ID {CONFIG['student_group_id']})"
    )

    print(
        "=" * 60
    )

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

    print()
    print(
        "=" * 60
    )

    print(
        f"Добавлено событий: "
        f"{len(events)}"
    )

    print(
        f"Готовый файл: "
        f"{output}"
    )

    print(
        "=" * 60
    )


if __name__ == "__main__":
    main()
