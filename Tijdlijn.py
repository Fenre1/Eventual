#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Maak een interactieve, verticaal scrollbare tijdlijn (HTML) op basis van een Excelbestand
met de volgende kolommen (exact of met witruimte varianten):
  - Datum                          (formaat: dd-mm-jjjj, bijvoorbeeld 25-09-2025)
  - Starttijd                      (formaat: HH:MM 24u)
  - Eindtijd                       (formaat: HH:MM 24u)
  - Zekerheid (ja/nee)             ('ja' of 'nee' of leeg)
  - Entiteit(en) (splits op met |) (meerdere entiteiten gescheiden door |)
  - Gebeurtenis                    (meerdere regels toegestaan)

Eigenschappen van de output:
  - Events worden als blok getoond; 
  - Events met wel datum maar geen tijd worden per dag onder de tijdlijn getoond, in volgorde van voorkomen in Excel.
  - Events zonder datum én tijd komen helemaal onderaan in een aparte tabel.
  - Filteren op entiteit (checkboxes). Filters gelden voor tijdlijn, dagsecties-zonder-tijd en de onderaan-tabel.
  - Onzekere tijden (Zekerheid = 'nee') krijgen een gestippelde rand en labeltje “Onzeker”.
  - Beschrijving ondersteunt meerdere regels.

Benodigdheden:
  - Python 3.9+
  - pandas
  - openpyxl (voor .xlsx lezen)

Installeren (indien nodig):
  pip install pandas openpyxl

Gebruik:
  python build_timeline.py --input pad/naar/events.xlsx --output timeline.html
Optionele parameters:
  --sheet "Bladnaam"              # standaard: eerste blad
  --title "Mijn tijdlijn"
  --px-per-minute 0.4             # verticale schaal; 0.4 ≈ 576 px per dag
  --min-event-minutes 15          # minimale hoogte voor punt-events
  --locale-nl                     # dag-/maandnaam in NL (client-side; alleen labels)

"""

import argparse
import datetime as dt
import json
import os
import re
import sys
import math
from collections import Counter
from pathlib import Path

import pandas as pd


# ---------- Helpers ----------

REQ_COLS = [
    "Datum",
    "Starttijd",
    "Eindtijd",
    "Zekerheid (ja/nee)",
    "Entiteit(en) (splits op met |)",
    "Gebeurtenis",
]

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    mapping_candidates = {
        "Datum": ["Datum", "datum", "Date"],
        "Starttijd": ["Starttijd", "starttijd", "Start", "Start tijd", "Start time"],
        "Eindtijd": ["Eindtijd", "eindtijd", "Einde", "Eind tijd", "End time"],
        "Zekerheid (ja/nee)": [
            "Zekerheid (ja/nee)","zekerheid (ja/nee)","Zekerheid","zekerheid","Exact?"
        ],
        "Entiteit(en) (splits op met |)": [
            "Entiteit(en) (splits op met |)","Entiteit(en)","Entiteiten","Entiteit","Entities"
        ],
        "Gebeurtenis": ["Gebeurtenis","gebeurtenis","Omschrijving","Beschrijving","Event"],
    }
    mapping, used = {}, set()
    for target, variants in mapping_candidates.items():
        for v in variants:
            if v in df.columns and v not in used:
                mapping[v] = target
                used.add(v)
                break
    return df.rename(columns=mapping)

def fmt_hhmm(total_minutes):
    if total_minutes is None:
        return None
    if isinstance(total_minutes, float):
        if math.isnan(total_minutes):
            return None
        total_minutes = int(total_minutes)
    else:
        total_minutes = int(total_minutes)
    h = total_minutes // 60
    m = total_minutes % 60
    return f"{h:02d}:{m:02d}"

def parse_date(s):
    if pd.isna(s):
        return None
    if isinstance(s, (dt.datetime, dt.date, pd.Timestamp)):
        return pd.to_datetime(s, errors="coerce").date()
    ss = str(s).strip()
    if not ss:
        return None
    for fmt in ("%d-%m-%Y", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(ss[:10], fmt).date()
        except Exception:
            pass
    iso_like = bool(re.match(r"^\d{4}-\d{2}-\d{2}(?:[ T]\d{2}:\d{2}:\d{2})?$", ss))
    if iso_like:
        val = pd.to_datetime(ss, errors="coerce", dayfirst=False)
        if pd.isna(val):
            return None
        return val.date()
    val = pd.to_datetime(ss, errors="coerce", dayfirst=True)
    if pd.isna(val):
        val = pd.to_datetime(ss, errors="coerce", dayfirst=False)
    if pd.isna(val):
        return None
    return val.date()

def parse_time_to_minutes(s):
    if pd.isna(s):
        return None
    ss = str(s).strip()
    if not ss:
        return None
    ss = ss.replace(".", ":")
    m = re.match(r"^(\d{1,2}):(\d{2})(?::(\d{2}))?$", ss)
    if m:
        h = int(m.group(1)); mm = int(m.group(2))
        if h == 24 and mm == 0:
            return 24 * 60
        if 0 <= h <= 23 and 0 <= mm <= 59:
            return h * 60 + mm
        return None
    m = re.match(r"^(\d{1,2})(\d{2})$", ss)  # HHMM
    if m:
        h = int(m.group(1)); mm = int(m.group(2))
        if 0 <= h <= 23 and 0 <= mm <= 59:
            return h * 60 + mm
        return None
    return None

def parse_zekerheid(v):
    if pd.isna(v):
        return None
    s = str(v).strip().lower()
    if s in {"ja", "j", "yes", "y", "true", "1"}:
        return True
    if s in {"nee", "n", "no", "false", "0"}:
        return False
    return None

def split_entities(s):
    if pd.isna(s):
        return []
    parts = [p.strip() for p in str(s).split("|")]
    return [p for p in parts if p]

def slugify(txt):
    if not txt:
        return ""
    t = txt.lower()
    t = re.sub(r"\s+", "-", t)
    t = re.sub(r"[^a-z0-9\-._~]", "", t)
    return t


# ---------- Data voorbereiding ----------

def load_and_prepare(path, sheet_name=None):
    if not os.path.exists(path):
        print(f"Bestand niet gevonden: {path}", file=sys.stderr)
        sys.exit(1)

    df = pd.read_excel(path, sheet_name=sheet_name, dtype=str)
    df = normalize_columns(df)

    missing = [c for c in REQ_COLS if c not in df.columns]
    if missing:
        print("Ontbrekende kolommen in Excel:", ", ".join(missing), file=sys.stderr)
        sys.exit(1)

    df["_row_order"] = range(len(df))
    df["__date"] = df["Datum"].apply(parse_date)
    df["__start_mins"] = df["Starttijd"].apply(parse_time_to_minutes)
    df["__end_mins"] = df["Eindtijd"].apply(parse_time_to_minutes)
    df["__certain"] = df["Zekerheid (ja/nee)"].apply(parse_zekerheid)
    df["__entities_list"] = df["Entiteit(en) (splits op met |)"].apply(split_entities)

    timed_segments, date_only, undated = [], [], []
    entity_counter = Counter()

    next_id = 1
    for _, row in df.iterrows():
        rid = next_id; next_id += 1

        date = row["__date"]
        start = row["__start_mins"]
        end = row["__end_mins"]
        entities = row["__entities_list"] or []
        desc = row["Gebeurtenis"] if not pd.isna(row["Gebeurtenis"]) else ""
        certain = row["__certain"]
        row_order = int(row["_row_order"])

        for e in entities:
            entity_counter[e] += 1

        if date is None:
            undated.append(
                {
                    "id": rid,
                    "desc": desc,
                    "entities": entities,
                    "entity_slugs": [slugify(x) for x in entities],
                    "certain": certain,
                    "row_order": row_order,
                }
            )
            continue

        if start is None and end is None:
            date_only.append(
                {
                    "id": rid,
                    "date": date.isoformat(),
                    "date_display": date.strftime("%d-%m-%Y"),
                    "desc": desc,
                    "entities": entities,
                    "entity_slugs": [slugify(x) for x in entities],
                    "certain": certain,
                    "row_order": row_order,
                }
            )
            continue

        if start is None and end is not None:
            start = end
            end = None

        if end is not None and end < start:
            timed_segments.append(
                {
                    "master_id": rid,
                    "seg_id": f"{rid}-a",
                    "date": date.isoformat(),
                    "date_display": date.strftime("%d-%m-%Y"),
                    "start_m": start,
                    "end_m": 24 * 60,
                    "has_duration": True,
                    "start_label": fmt_hhmm(start),
                    "end_label": "24:00",
                    "desc": desc,
                    "entities": entities,
                    "entity_slugs": [slugify(x) for x in entities],
                    "certain": certain,
                    "row_order": row_order,
                    "continues_next": True,
                    "continues_prev": False,
                }
            )
            next_date = date + dt.timedelta(days=1)
            timed_segments.append(
                {
                    "master_id": rid,
                    "seg_id": f"{rid}-b",
                    "date": next_date.isoformat(),
                    "date_display": next_date.strftime("%d-%m-%Y"),
                    "start_m": 0,
                    "end_m": end,
                    "has_duration": True,
                    "start_label": "00:00",
                    "end_label": fmt_hhmm(end),
                    "desc": desc,
                    "entities": entities,
                    "entity_slugs": [slugify(x) for x in entities],
                    "certain": certain,
                    "row_order": row_order,
                    "continues_next": False,
                    "continues_prev": True,
                }
            )
        else:
            has_duration = end is not None and start is not None
            timed_segments.append(
                {
                    "master_id": rid,
                    "seg_id": f"{rid}-s",
                    "date": date.isoformat(),
                    "date_display": date.strftime("%d-%m-%Y"),
                    "start_m": start,
                    "end_m": end if has_duration else None,
                    "has_duration": has_duration,
                    "start_label": fmt_hhmm(start),
                    "end_label": fmt_hhmm(end),
                    "desc": desc,
                    "entities": entities,
                    "entity_slugs": [slugify(x) for x in entities],
                    "certain": certain,
                    "row_order": row_order,
                    "continues_next": False,
                    "continues_prev": False,
                }
            )

    timed_segments.sort(key=lambda r: (r["date"], (r["start_m"] if r["start_m"] is not None else -1), r["row_order"]))
    date_only.sort(key=lambda r: (r["date"], r["row_order"]))
    undated.sort(key=lambda r: r["row_order"])

    entities_sorted = sorted(entity_counter.items(), key=lambda kv: (-kv[1], kv[0].lower()))
    unique_entities = [name for name, _ in entities_sorted]
    return timed_segments, date_only, undated, unique_entities


# ---------- HTML Template (STACKED LAYOUT) ----------
TEMPLATE_PATH = Path("timeline_template.html")



def load_html_template() -> str:
    """Load the HTML visualisation template from disk."""
    try:
        return TEMPLATE_PATH.read_text(encoding="utf-8")
    except FileNotFoundError as exc:
        raise FileNotFoundError(
            f"HTML template not found: {TEMPLATE_PATH}"
        ) from exc


# ---------- HTML Generator ----------

def make_html(title, timed_segments, date_only, undated, unique_entities, px_per_min=0.4, min_event_minutes=15):
    # px_per_min is no longer used in the stacked layout; kept for signature compatibility
    payload = {
        "timed": timed_segments,
        "date_only": date_only,
        "undated": undated,
    }
    template = load_html_template()
    html = (
        template
        .replace("%%TITLE%%", (title or "Tijdlijn").replace("\\", "\\\\").replace("</", "<\\/"))
        .replace("%%MIN_EVENT_MIN%%", str(int(min_event_minutes)))
        .replace("%%ALL_ENTITIES_JSON%%", json.dumps(unique_entities, ensure_ascii=False))
        .replace("%%DATA%%", json.dumps(payload, ensure_ascii=False))
    )
    return html


# ---------- Notebook helper ----------

def build_timeline_notebook(
    input_path,
    sheet=None,
    title="Tijdlijn",
    px_per_minute=0.4,       # ignored in the stacked layout
    min_event_minutes=15,
    output="timeline.html",  # write to disk
    display_inline=False     # if you want to display inside a notebook, use IPython.display(HTML(html))
):
    # Convert sheet index if string
    sheet_name = sheet
    if sheet_name is not None and isinstance(sheet_name, str):
        try:
            sheet_name = int(sheet_name)
        except Exception:
            pass

    timed_segments, date_only, undated, unique_entities = load_and_prepare(
        input_path, sheet_name=sheet_name
    )
    html = make_html(
        title=title,
        timed_segments=timed_segments,
        date_only=date_only,
        undated=undated,
        unique_entities=unique_entities,
        px_per_min=px_per_minute,
        min_event_minutes=min_event_minutes,
    )

    saved_path = None
    if output:
        out_path = Path(output)
        out_path.write_text(html, encoding="utf-8")
        saved_path = str(out_path.resolve())

    if display_inline:
        from IPython.display import HTML, display
        display(HTML(html))

    return {"html": html, "saved_path": saved_path}


input_path = 'Example.xlsx'
sheet="Blad1"
title="Tijdlijn"
px_per_minute=0.4
min_event_minutes=15
output='timeline.html'
display_inline=False
build_timeline_notebook(input_path,sheet,title,px_per_minute,min_event_minutes,output,display_inline)


# ---------- CLI ----------

def main():
    parser = argparse.ArgumentParser(description="Genereer een interactieve HTML-tijdlijn vanuit Excel.")
    parser.add_argument("--input", required=True, help="Pad naar Excel (.xlsx)")
    parser.add_argument("--sheet", default=None, help="Bladnaam of index (optioneel)")
    parser.add_argument("--output", default="timeline.html", help="Uitvoer HTML-bestand (standaard: timeline.html)")
    parser.add_argument("--title", default="Tijdlijn", help="Titel in de HTML")
    parser.add_argument("--px-per-minute", type=float, default=0.4, help="Verticale schaal px/minuut (bv. 0.4 ≈ 576px per dag)")
    parser.add_argument("--min-event-minutes", type=int, default=15, help="Minimale duur (min) voor punt-events, voor zichtbaarheid en overlapindeling")
    args = parser.parse_args()

    sheet_name = args.sheet
    # Indien sheet een getal is meegegeven, converteer naar int
    if sheet_name is not None:
        try:
            sheet_name = int(sheet_name)
        except Exception:
            pass

    timed_segments, date_only, undated, unique_entities = load_and_prepare(args.input, sheet_name=sheet_name)
    html = make_html(
        args.title,
        timed_segments,
        date_only,
        undated,
        unique_entities,
        px_per_min=args.px_per_minute,
        min_event_minutes=args.min_event_minutes,
    )

    with open(args.output, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"Gereed. Bestand geschreven: {os.path.abspath(args.output)}")


if __name__ == "__main__":
    main()
