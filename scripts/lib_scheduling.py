"""
Shared constants + placement logic for assigning students to a Day/Block/Room/slot.

The block structure below (4 blocks Thursday, 2 blocks Friday ending at
noon) comes directly from the mentor scheduling-preferences form -- it's
the real structure mentors were told about and ranked their availability
against, not an arbitrary split. See MENTOR_BLOCK_OVERRIDES for how their
responses were translated into placements.
"""
from collections import Counter, defaultdict
from datetime import datetime, timedelta

MINUTES_PER_TALK = 16

ROOMS = ["32-124", "32-141", "32-155"]

# Ordered list of blocks across both days. Each block is an independent
# room-parallel session; capacity is per room, per block.
BLOCKS = [
    {"code": "T1", "date": "2026-08-06", "weekday": "Thursday", "start": "09:00", "end": "10:30", "cap": 6},
    {"code": "T2", "date": "2026-08-06", "weekday": "Thursday", "start": "11:00", "end": "12:30", "cap": 6},
    {"code": "T3", "date": "2026-08-06", "weekday": "Thursday", "start": "13:15", "end": "14:45", "cap": 6},
    {"code": "T4", "date": "2026-08-06", "weekday": "Thursday", "start": "15:15", "end": "16:45", "cap": 6},
    {"code": "F1", "date": "2026-08-07", "weekday": "Friday", "start": "09:00", "end": "10:30", "cap": 6},
    {"code": "F2", "date": "2026-08-07", "weekday": "Friday", "start": "11:00", "end": "12:00", "cap": 4},
]
BLOCK_BY_CODE = {b["code"]: b for b in BLOCKS}

DAYS = [
    {"date": "2026-08-06", "weekday": "Thursday", "display": "Thursday, August 6, 2026",
     "block_codes": ["T1", "T2", "T3", "T4"]},
    {"date": "2026-08-07", "weekday": "Friday", "display": "Friday, August 7, 2026",
     "block_codes": ["F1", "F2"]},
]

# Breaks shown between consecutive blocks on the same day (for display only).
BREAK_AFTER = {
    "T1": "Break (10:30 AM – 11:00 AM)",
    "T2": "Lunch Break (12:30 PM – 1:15 PM)",
    "T3": "Break (2:45 PM – 3:15 PM)",
    "F1": "Break (10:30 AM – 11:00 AM)",
}

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

# --- Mentor scheduling-preference placements -------------------------------
# Derived by hand from "RSI 2026 Symposium Scheduling Preferences" form
# responses (mentor-ranked block availability + comments), cross-referenced
# against each mentor's actual student group in data/roster.csv. Where a PI
# and a co-mentor/lab member both responded with different rankings, the
# PI's (the name matching roster.csv's Mentor column) was weighted higher.
# Where a mentor was simply "Unable to attend" every option, or gave no
# preference, they're left out of this dict entirely and their students
# fall through to the field-clustering placement below instead.

# Some students' roster "mentor" field includes a co-mentor the preference
# form's respondent didn't (e.g. "Prof. Jenny Hoffman, Sarah Ziegler" vs
# "Prof. Jenny Hoffman") -- these are still the same lab / same block.
MENTOR_ALIASES = {
    "Prof. Jenny Hoffman, Sarah Ziegler": "Prof. Jenny Hoffman",
}

MENTOR_BLOCK_OVERRIDES = {
    "Dr. Prajwal T. Mohan Murthy": "T4",
    "Dr. Aaron Kesselheim": "F1",
    "Prof. Markus Buehler": "F1",
    "Hang Du": "T2",
    "Ashu Tripathi": "T2",
    "Dr. Daniel Press and Dr. Hsueh-Sheng Chiang": "T1",
    "Dr. Mrinal Shekhar": "T2",
    "Prof. Fengfeng Bei": "F2",
    "Prof. Jenny Hoffman": "F1",
    "Dr. Robbie G. Majzner": "T2",
    "Ayan Nath": "T2",
    "Prof. Paul O'Gorman": "T1",
    "Prof. James Michaelson": "F2",
    "Prof. Gil Alterovitz": "T4",
    "Prof. Timothy Miller": "T4",
    "Katharine Hesse": "T4",
    "Prof. Maggie Qi, Shu Yang": "F2",
    "Prof. Ron Weiss": "T3",
    "Prof. John D. E. Gabrieli": "T3",
    "Prof. Charles Hoffman": "F1",
    "Prof. Francis Loth and Hannah Higgins": "T2",
    "Prof. Chuchu Fan": "T3",
    "Dr. Yogesh Rathi, Dr. Sinead Marie Kelly, Dr. Lipeng Ning": "T3",
    "Dr. Howard Chen": "T2",
    "Prof. Eric Grinberg": "T1",
    "Prof. Susan Hagen, Dr. Xu Xu": "F1",
    "Dr. Sarah Bricault and Prof. Alan Pradip Jasanoff (PI)": "T3",
    "Prof. Sara Seager": "T1",
    "Dr. Ky Lowenhaupt": "T3",
    "Dr. David T. Miyamoto": "F1",
    "Prof. Koroush Shirvan": "F1",
    "Dr. Sean Carroll": "T2",
    "Dr. Artur A. Indzhykulian": "F1",
    "Prof. Min Dong": "T4",
    "Prof. Vivek Venkatachalam": "T1",
    "Dr. Cong Liu, Dr. Kamalakkannan Ravi": "F1",
    "Dr. Sudeshna Das": "T3",
    "Dr. Rosanne Di Stefano": "T2",
    "Dr. Basak Icli": "F2",
    "Prof. Vladan Vuletic": "T2",
    # Not covered by their own submission, but pulled into the early-Thursday
    # math cluster David Jerison specifically requested (see comment on his
    # response): all "core" math mentor-groups plus Eric Grinberg's 3.
    "Alexander McWeeney": "T1",
    "Ryota Inagaki": "T1",
    "Siqi Wu": "T1",
}

# Groups too large for one room-block window; split across two *consecutive*
# blocks in the same room so the mentor can attend both without moving rooms.
SPLIT_OVERRIDES = {
    "Prof. Pierre F.J Lermusiaux": [("T3", 5), ("T4", 2)],
}

# Explicit within-group ordering requests (e.g. "Mahi right before Alejandra").
STUDENT_ORDER_OVERRIDES = {
    "Prof. Vivek Venkatachalam": ["Mahi Kohli", "Alejandra Stephanie Perry", "Ferris Li"],
}

FIELD_ORDER = [
    "Math", "Physics", "Astronomy & Planetary Science",
    "Computer Science & AI", "Engineering", "Chemistry & Materials",
    "Biology & Medicine", "Neuroscience & Psychology",
    "Earth, Ocean & Climate", "Economics & Policy", "Other",
]


def group_by_mentor(students):
    """Group students by mentor (after alias normalization), preserving
    first-appearance order. Returns {canonical_mentor_key: [students]}."""
    seen = {}
    order = []
    for s in students:
        key = MENTOR_ALIASES.get(s["mentor"], s["mentor"])
        if key not in seen:
            seen[key] = []
            order.append(key)
        seen[key].append(s)
    return {k: seen[k] for k in order}


def make_room_state():
    return {b["code"]: {room: [] for room in ROOMS} for b in BLOCKS}


def place_group(state, code, students, preferred_room=None):
    """Place an intact mentor-group into the first room in `code` with
    enough remaining capacity. Returns the room used, or None if it doesn't
    fit anywhere (caller should handle/report this)."""
    cap = BLOCK_BY_CODE[code]["cap"]
    room_order = ROOMS if not preferred_room else (
        [preferred_room] + [r for r in ROOMS if r != preferred_room])
    for room in room_order:
        if len(state[code][room]) + len(students) <= cap:
            state[code][room].extend(students)
            return room
    return None


def pack(students):
    """Returns (state, warnings). state[block_code][room] -> list[student]."""
    mentor_groups = group_by_mentor(students)
    state = make_room_state()
    warnings = []
    placed_keys = set()

    # 1. Explicit splits first (order-sensitive, same room across blocks).
    for key, parts in SPLIT_OVERRIDES.items():
        group = mentor_groups.get(key)
        if not group:
            warnings.append(f"SPLIT_OVERRIDES key '{key}' not found in roster mentors")
            continue
        idx = 0
        room = None
        for code, count in parts:
            chunk = group[idx: idx + count]
            idx += count
            used = place_group(state, code, chunk, preferred_room=room)
            if used is None:
                warnings.append(f"Could not place split chunk for '{key}' in block {code}")
            else:
                room = used
        placed_keys.add(key)

    # 2. Single-block overrides (preference-driven placements).
    for key, code in MENTOR_BLOCK_OVERRIDES.items():
        group = mentor_groups.get(key)
        if not group:
            warnings.append(f"MENTOR_BLOCK_OVERRIDES key '{key}' not found in roster mentors")
            continue
        if key in STUDENT_ORDER_OVERRIDES:
            order = STUDENT_ORDER_OVERRIDES[key]
            group = sorted(group, key=lambda s: order.index(s["full_name"]) if s["full_name"] in order else 99)
        used = place_group(state, code, group)
        if used is None:
            warnings.append(f"Could not fit '{key}' ({len(group)} students) into block {code}")
        placed_keys.add(key)

    # 3. Everyone else: field-clustering fallback into remaining capacity.
    leftover_keys = [k for k in mentor_groups if k not in placed_keys]
    by_field = defaultdict(list)
    for key in leftover_keys:
        group = mentor_groups[key]
        dominant_field = Counter(s["field"] for s in group).most_common(1)[0][0]
        by_field[dominant_field].append(group)

    block_codes_in_order = [b["code"] for b in BLOCKS]
    for field in FIELD_ORDER:
        for group in by_field.get(field, []):
            placed = False
            for code in block_codes_in_order:
                if place_group(state, code, group) is not None:
                    placed = True
                    break
            if not placed:
                warnings.append(f"Could not fit fallback group (mentor={group[0]['mentor']!r}, "
                                 f"{len(group)} students) into any block")

    return state, warnings


def slot_times(start_str, count):
    t = datetime.strptime(start_str, "%H:%M")
    times = []
    for _ in range(count):
        end = t + timedelta(minutes=MINUTES_PER_TALK)
        times.append((t.strftime("%-I:%M %p"), end.strftime("%-I:%M %p")))
        t = end
    return times
