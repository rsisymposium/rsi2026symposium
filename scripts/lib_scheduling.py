"""
Shared constants + the mentor/field-aware bin-packing heuristic used to
produce an initial draft assignment of students to Day/Block/Room/slot.

Used by 02_build_schedule_workbook.py to pre-fill the Schedule sheet. The
workbook it produces is the actual source of truth after that -- rerunning
this only matters if you want to regenerate a fresh algorithmic draft (e.g.
the roster changed a lot), not for everyday hand-edits.
"""
from collections import Counter, defaultdict
from datetime import datetime, timedelta

MINUTES_PER_TALK = 16
AM_START = "09:00"
PM_START = "13:00"
LUNCH_LABEL = "Lunch Break (12:00 PM – 1:00 PM)"
# Capacities keep every section comfortably inside its window:
# AM 9:00-12:00 lunch (180 min avail, 10 talks = 160 min); PM 13:00-16:45 (225 min avail, 13 talks = 208 min).
SECTION_CAPACITY = {"AM": 10, "PM": 13}
# Once a section reaches this many talks, prefer starting a fresh section for
# the next field rather than blending; below it, let the next field spill in
# so small fields (e.g. Economics & Policy) don't strand a near-empty section.
MIN_SECTION_SIZE_BEFORE_BREAK = 5

ROOMS = ["32-124", "32-141", "32-155"]
DAYS = [
    {"label": "Day 1", "date": "2026-08-06", "weekday": "Thursday", "display": "Thursday, August 6, 2026"},
    {"label": "Day 2", "date": "2026-08-07", "weekday": "Friday", "display": "Friday, August 7, 2026"},
]
BLOCKS = ["AM", "PM"]
BLOCK_LABELS = {"AM": "Morning Session", "PM": "Afternoon Session"}

# Order fields are considered in when packing -- adjust to change which
# subjects tend to land near each other across the day.
FIELD_ORDER = [
    "Math", "Physics", "Astronomy & Planetary Science",
    "Computer Science & AI", "Engineering", "Chemistry & Materials",
    "Biology & Medicine", "Neuroscience & Psychology",
    "Earth, Ocean & Climate", "Economics & Policy", "Other",
]

FIELD_COLORS_HEX = {
    "Math": ("DBEAFE", "1D4ED8"),
    "Physics": ("EDE9FE", "6D28D9"),
    "Astronomy & Planetary Science": ("E0E7FF", "4338CA"),
    "Computer Science & AI": ("CFFAFE", "0E7490"),
    "Engineering": ("FFEDD5", "C2410C"),
    "Chemistry & Materials": ("FEF3C7", "B45309"),
    "Biology & Medicine": ("D1FAE5", "047857"),
    "Neuroscience & Psychology": ("FCE7F3", "BE185D"),
    "Earth, Ocean & Climate": ("CCFBF1", "0F766E"),
    "Economics & Policy": ("FFE4E6", "BE123C"),
    "Other": ("F1F5F9", "475569"),
}


def build_mentor_blocks(students):
    """Group students by mentor (exact match) first, so a mentor's students
    always stay together and contiguous even if their titles land in
    different fields. Each mentor-block is then filed under its *dominant*
    field (majority vote among its members) for the purposes of clustering
    sections by subject."""
    seen_mentors = {}
    order = []
    for s in students:
        key = s["mentor"]
        if key not in seen_mentors:
            seen_mentors[key] = []
            order.append(key)
        seen_mentors[key].append(s)

    field_blocks = defaultdict(list)  # field -> list of mentor-blocks (list of students)
    for key in order:
        block = seen_mentors[key]
        dominant_field = Counter(s["field"] for s in block).most_common(1)[0][0]
        field_blocks[dominant_field].append(block)
    return field_blocks


def make_sections():
    sections = []
    for day in DAYS:
        for block in BLOCKS:
            for room in ROOMS:
                sections.append({"day": day, "block": block, "room": room, "students": []})
    return sections


def pack(field_blocks):
    sections = make_sections()
    sec_idx = 0

    def current():
        return sections[sec_idx]

    def capacity(section):
        return SECTION_CAPACITY[section["block"]]

    def advance():
        nonlocal sec_idx
        if sec_idx + 1 >= len(sections):
            raise RuntimeError(
                "Ran out of sections while packing. Raise SECTION_CAPACITY, "
                "or add rooms/blocks/days."
            )
        sec_idx += 1

    for field in FIELD_ORDER:
        blocks = field_blocks.get(field, [])
        for block in blocks:
            while len(current()["students"]) + len(block) > capacity(current()):
                advance()
            current()["students"].extend(block)
        if current()["students"] and sec_idx + 1 < len(sections):
            if len(current()["students"]) >= MIN_SECTION_SIZE_BEFORE_BREAK:
                sec_idx += 1

    used = [s for s in sections if s["students"]]
    unused = [s for s in sections if not s["students"]]
    return used, unused


def slot_times(block, count):
    start = AM_START if block == "AM" else PM_START
    t = datetime.strptime(start, "%H:%M")
    times = []
    for _ in range(count):
        end = t + timedelta(minutes=MINUTES_PER_TALK)
        times.append((t.strftime("%-I:%M %p"), end.strftime("%-I:%M %p")))
        t = end
    return times
