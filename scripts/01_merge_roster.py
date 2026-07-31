"""
Merge the two raw source files into one clean roster CSV.

Inputs (checked into the repo root):
  - RSI 2026 Name Pronunciations.xlsx   (name, pronunciation, email, mentor)
  - Final Presentation Title (Responses) - Form Responses 1.csv  (kerb -> title, may have edits/duplicates)

Output:
  - data/roster.csv

The FIELD_OVERRIDES dict below is a manual classification of each student's
subject area (Math / Physics / Astro / CS-AI / Bio / Neuro-Psych / Chem /
Engineering / Earth-Ocean / Econ-Policy), read off their title + mentor's
department. There's no reliable automatic way to do this, so if a student's
field looks wrong, or a new student is added, just edit the dict entry here
and rerun this script.
"""
import csv
import re
from datetime import datetime
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
XLSX_PATH = ROOT / "RSI 2026 Name Pronunciations.xlsx"
CSV_PATH = ROOT / "Final Presentation Title (Responses) - Form Responses 1.csv"
OUT_PATH = ROOT / "data" / "roster.csv"

# kerb -> subject field. Edit freely; rerun this script to regenerate data/roster.csv.
FIELD_OVERRIDES = {
    "a_sultan": "Bio", "a1wang": "Bio", "aaru0302": "CS-AI", "abihqi": "Bio",
    "advg": "CS-AI", "alej0": "CS-AI", "alexicho": "Bio", "alexren": "Bio",
    "amogh28": "Bio", "amritha9": "Bio", "annika8": "Chem", "aowang11": "Physics",
    "asperry": "Bio", "atyagi11": "Earth-Ocean", "avispute": "Earth-Ocean",
    "avtonapi": "Math", "azariart": "Physics", "bensonh": "CS-AI",
    "can_nlbt": "CS-AI", "cfyzhan": "Earth-Ocean", "chloegig": "Bio",
    "cho0707": "Earth-Ocean", "ctimothy": "Math", "cyh": "Astro",
    "dagandhi": "Econ-Policy", "dashii": "Chem", "davidw09": "Chem",
    "dguo103": "Engineering", "dhruvj07": "Physics", "dipek": "Chem",
    "dkim081": "CS-AI", "doskhan": "Physics", "dtheisen": "Chem",
    "eason": "Astro", "ebuehler": "CS-AI", "grishpai": "Math",
    "gsvinoy": "Bio", "hannahgt": "Bio", "haya_m24": "Bio", "hl888": "CS-AI",
    "iza_c397": "Bio", "jayant69": "Math", "jaydesai": "CS-AI",
    "jordanrc": "Engineering", "karlc24": "Neuro-Psych", "kmradddd": "Bio",
    "larao369": "Bio", "layth224": "CS-AI", "leobart": "CS-AI",
    "leyug": "Math", "lillieli": "Astro", "litomura": "Bio", "liukz": "Math",
    "lliu320": "Math", "lluc": "Physics", "lucichen": "Bio",
    "lwduran": "Engineering", "mahi_k19": "Bio", "mariac08": "Bio",
    "matey": "CS-AI", "mattlo": "Bio", "mbakhod": "Engineering",
    "mikolaj": "Physics", "minkyujo": "Bio", "mira_c36": "Neuro-Psych",
    "mjmvega": "Earth-Ocean", "mlele123": "Econ-Policy", "mluo928": "Math",
    "nabet": "Neuro-Psych", "namie": "CS-AI", "ngawang1": "Physics",
    "nveselin": "Math", "psubra18": "Chem", "raquelgj": "Bio",
    "rishi007": "Neuro-Psych", "riyam114": "Math", "rohanpr": "CS-AI",
    "ruchram8": "Bio", "sarmit21": "Bio", "selenem": "Neuro-Psych",
    "sgpark": "Bio", "shadenfm": "Bio", "shim112": "Physics",
    "shrivali": "Neuro-Psych", "siddhik": "Astro", "sofia937": "Bio",
    "sophiawz": "Earth-Ocean", "sowmya44": "CS-AI", "srdingli": "Bio",
    "sirius_s": "Math", "szl129": "Physics", "tarora": "CS-AI",
    "tlu27": "Math", "tohk140": "Astro", "tris": "Bio", "tsaini": "Math",
    "xindi578": "Math", "yding27": "Earth-Ocean", "yspyo": "Neuro-Psych",
    "zephyr27": "Earth-Ocean",
}

FIELD_LABELS = {
    "Math": "Math",
    "Physics": "Physics",
    "Astro": "Astronomy & Planetary Science",
    "CS-AI": "Computer Science & AI",
    "Bio": "Biology & Medicine",
    "Neuro-Psych": "Neuroscience & Psychology",
    "Chem": "Chemistry & Materials",
    "Engineering": "Engineering",
    "Earth-Ocean": "Earth, Ocean & Climate",
    "Econ-Policy": "Economics & Policy",
}


def parse_ts(s):
    try:
        return datetime.strptime(s.strip(), "%m/%d/%Y %H:%M:%S")
    except (ValueError, AttributeError):
        return datetime.min


def load_students():
    wb = openpyxl.load_workbook(XLSX_PATH, data_only=True)
    ws = wb["Sheet1"]
    rows = list(ws.iter_rows(values_only=True))
    students = {}
    for r in rows[1:]:
        if not r[4]:
            continue
        email = r[4].strip().lower()
        kerb = email.split("@")[0]
        students[kerb] = {
            "kerb": kerb,
            "first": (r[0] or "").strip(),
            "last": (r[1] or "").strip(),
            "first_pron": (r[2] or "").strip(),
            "last_pron": (r[3] or "").strip(),
            "email": email,
            "printed_name": (r[5] or f"{r[0]} {r[1]}").strip(),
            "mentor": (r[6] or "").strip(),
        }
    return students


def load_titles():
    best = {}
    with open(CSV_PATH, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        title_col = "Final Presentation Title (you can edit later)"
        for row in reader:
            kerb = row['kerb (your mit email without "", ex: limevan)'].strip().lower()
            if not kerb:
                continue
            ts = parse_ts(row["Timestamp"])
            title = re.sub(r"\s+", " ", row[title_col].strip())
            if kerb not in best or ts >= best[kerb][0]:
                best[kerb] = (ts, title)
    return {k: v[1] for k, v in best.items()}


def main():
    students = load_students()
    titles = load_titles()

    rows = []
    for kerb, s in students.items():
        title = titles.get(kerb, "(Title TBD)")
        field_code = FIELD_OVERRIDES.get(kerb)
        if field_code is None:
            print(f"WARNING: no field assigned for {kerb} ({s['printed_name']}) -- defaulting to 'Other'")
            field_code = "Other"
        field_label = FIELD_LABELS.get(field_code, "Other")
        rows.append({
            "kerb": kerb,
            "full_name": s["printed_name"],
            "first_pron": s["first_pron"],
            "last_pron": s["last_pron"],
            "email": s["email"],
            "title": title,
            "mentor": s["mentor"],
            "field": field_label,
        })
        if kerb not in titles:
            print(f"NOTE: {kerb} ({s['printed_name']}) has no submitted title yet -- using placeholder")

    rows.sort(key=lambda r: r["full_name"])

    OUT_PATH.parent.mkdir(exist_ok=True)
    with open(OUT_PATH, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=[
            "kerb", "full_name", "first_pron", "last_pron", "email",
            "title", "mentor", "field",
        ])
        writer.writeheader()
        writer.writerows(rows)

    print(f"\nWrote {len(rows)} students to {OUT_PATH}")


if __name__ == "__main__":
    main()
