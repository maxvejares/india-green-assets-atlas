"""
create_india_overview.py
India Green Assets Atlas — Project Overview Document
Generates a Word doc summarising dataset, cleaning decisions, flagged issues,
and geocoding results for colleague review.
"""

import json
from pathlib import Path
from collections import Counter, defaultdict
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BASE = Path(__file__).parent

# ── helpers ──────────────────────────────────────────────────────────────────

def set_cell_bg(cell, hex_color):
    """Set table cell background colour."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)

def add_heading(doc, text, level, color=None):
    h = doc.add_heading(text, level=level)
    if color:
        for run in h.runs:
            run.font.color.rgb = RGBColor(*color)
    return h

def add_para(doc, text, bold=False, italic=False, size=None, color=None, left_indent=None):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    if size:
        run.font.size = Pt(size)
    if color:
        run.font.color.rgb = RGBColor(*color)
    if left_indent:
        p.paragraph_format.left_indent = Inches(left_indent)
    return p

def add_bullet(doc, text, level=0):
    p = doc.add_paragraph(style="List Bullet")
    p.add_run(text)
    p.paragraph_format.left_indent = Inches(0.25 * (level + 1))
    return p

def add_table_header(table, headers, bg_hex="1F4E79", font_color=(255, 255, 255)):
    row = table.rows[0]
    for i, h in enumerate(headers):
        cell = row.cells[i]
        cell.text = ""
        run = cell.paragraphs[0].add_run(h)
        run.bold = True
        run.font.color.rgb = RGBColor(*font_color)
        run.font.size = Pt(10)
        set_cell_bg(cell, bg_hex)

def add_data_row(table, values, shade=False, font_size=9):
    row = table.add_row()
    for i, val in enumerate(values):
        cell = row.cells[i]
        cell.text = str(val) if val is not None else ""
        for para in cell.paragraphs:
            for run in para.runs:
                run.font.size = Pt(font_size)
        if shade:
            set_cell_bg(cell, "EBF3FB")
    return row

def set_col_widths(table, widths_cm):
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            if i < len(widths_cm):
                cell.width = Cm(widths_cm[i])

# ── load data ────────────────────────────────────────────────────────────────

SHEETS = ["solar_mfg", "evs", "wind", "battery", "green_steel", "hydrogen"]
SHEET_LABELS = {
    "solar_mfg": "Solar Manufacturing",
    "evs": "Electric Vehicles (EVs)",
    "wind": "Wind Energy",
    "battery": "Battery Storage",
    "green_steel": "Green Steel",
    "hydrogen": "Green Hydrogen",
}

records_by_sheet = {}
for s in SHEETS:
    records_by_sheet[s] = json.load(open(BASE / f"{s}.json"))

all_records = [r for recs in records_by_sheet.values() for r in recs]
cleaning_flags = json.load(open(BASE / "_cleaning_flags.json"))
geocode_fallbacks = json.load(open(BASE / "_geocode_fallbacks.json"))

# ── build doc ────────────────────────────────────────────────────────────────

doc = Document()

# Page margins
for section in doc.sections:
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.0)

# ─────────────────────────────────────────────────────────────────────────────
# COVER / TITLE
# ─────────────────────────────────────────────────────────────────────────────

doc.add_paragraph()
title = doc.add_heading("India Green Assets Atlas", 0)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER

subtitle = doc.add_paragraph("Project Overview — Data Cleaning, Geocoding & Review Notes")
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
for run in subtitle.runs:
    run.font.size = Pt(13)
    run.font.color.rgb = RGBColor(89, 89, 89)

meta = doc.add_paragraph("Prepared: February 2026  ·  6 sectors  ·  439 records")
meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
for run in meta.runs:
    run.font.size = Pt(10)
    run.font.italic = True
    run.font.color.rgb = RGBColor(127, 127, 127)

doc.add_paragraph()

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 1 — DATASET SUMMARY
# ─────────────────────────────────────────────────────────────────────────────

add_heading(doc, "1.  Dataset Summary", 1)

add_para(doc, (
    "The India Green Assets Atlas covers six green industrial sectors drawn from a "
    "single Excel workbook. All sheets were cleaned and standardised for use in a "
    "GIS platform. The table below shows record counts, coordinate availability, "
    "and key data quality notes per sheet."
))
doc.add_paragraph()

# Summary table
tbl = doc.add_table(rows=1, cols=6)
tbl.style = "Table Grid"
tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
add_table_header(tbl, ["Sheet", "Records", "With Coords", "Missing Coords", "Status Flags", "Notes"])

coord_miss = {
    "solar_mfg": 94, "evs": 0, "wind": 9, "battery": 2, "green_steel": 1, "hydrogen": 0
}
status_flags = Counter(f["sheet"] for f in cleaning_flags if f["field"] == "status")
sheet_notes = {
    "solar_mfg": "94 records have no city or state — not geocodable",
    "evs":        "All records geocoded; 1 city name ambiguous (state capital used)",
    "wind":       "City/State split from combined field; 3 records no location",
    "battery":    "City extracted from full addresses; 11 fallback to state capital",
    "green_steel": "4 records; 1 city 'Salav' not found — fallback to Maharashtra capital",
    "hydrogen":   "All records geocoded; 4 ambiguous names — fallback to state capital",
}

for i, s in enumerate(SHEETS):
    recs = records_by_sheet[s]
    total = len(recs)
    has_coord = sum(1 for r in recs if r.get("latitude") and r.get("longitude"))
    missing = coord_miss[s]
    sflag = status_flags.get(s, 0)
    shade = (i % 2 == 1)
    add_data_row(tbl, [
        SHEET_LABELS[s], total, has_coord, missing, sflag, sheet_notes[s]
    ], shade=shade)

set_col_widths(tbl, [3.8, 1.5, 1.8, 2.0, 1.7, 6.0])

doc.add_paragraph()
add_para(doc, (
    f"Total records: {len(all_records)}  ·  "
    f"Records with coordinates: {sum(1 for r in all_records if r.get('latitude'))}/439  ·  "
    f"Cleaning flags raised: {len(cleaning_flags)}  ·  "
    f"Geocode fallbacks (state capital): {sum(1 for r in all_records if r.get('geocode_note'))}"
), italic=True, size=10)

doc.add_page_break()

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 2 — CLEANING DECISIONS LOG
# ─────────────────────────────────────────────────────────────────────────────

add_heading(doc, "2.  Cleaning Decisions Log", 1)
add_para(doc, (
    "This section documents every non-obvious decision made during cleaning. "
    "These decisions are applied consistently to all records."
))
doc.add_paragraph()

# 2.1 Columns dropped
add_heading(doc, "2.1  Columns Dropped", 2)
add_bullet(doc, "Region — 0% fill rate across all sheets. Removed from all outputs.")
add_bullet(doc, "Name — 3–8% fill rate (solar: 3%, EVs: 8%). Company field used instead.")
doc.add_paragraph()

# 2.2 Status vocabulary
add_heading(doc, "2.2  Status Vocabulary Standardisation", 2)
add_para(doc, (
    "The 'Status' field had inconsistent free-text values across all sheets. "
    "All values were mapped to a controlled five-value vocabulary:"
))
vocab_tbl = doc.add_table(rows=1, cols=3)
vocab_tbl.style = "Table Grid"
add_table_header(vocab_tbl, ["Controlled Value", "Maps From", "Notes"])
vocab_map = [
    ("Operational", "commissioned, open, operational, active, in operation",
     "Plant is currently running"),
    ("Announced", "announced, announced in 2022/2024/2025, MoU signed, MoU signed in 2023",
     "Announcement year moved to Notes field where present"),
    ("Under Construction", "under construction, construction underway",
     ""),
    ("Planned", "planned, greenfield planned",
     ""),
    ("Closed / Decommissioned", "closed, decommissioned",
     ""),
    ("Unknown", "(blank / null / unrecognised)",
     "Flagged for colleague review — see Section 3"),
]
for j, (val, maps_from, notes) in enumerate(vocab_map):
    shade = (j % 2 == 1)
    add_data_row(vocab_tbl, [val, maps_from, notes], shade=shade)
set_col_widths(vocab_tbl, [3.5, 7.0, 5.5])
doc.add_paragraph()

# 2.3 Location parsing
add_heading(doc, "2.3  Location Field Parsing", 2)
add_para(doc, "Several sheets stored city and state together in one field. Each case was handled as follows:")
parsing_tbl = doc.add_table(rows=1, cols=3)
parsing_tbl.style = "Table Grid"
add_table_header(parsing_tbl, ["Sheet", "Original Format", "Parsing Rule"])
parsing_rows = [
    ("Solar Manufacturing", "Separate Latitude/Longitude column ('lat, lon')",
     "Split on comma; cast to float"),
    ("EVs", "Separate Latitude/Longitude column ('lat, lon')",
     "Split on comma; cast to float"),
    ("Wind", "'City, State' combined in City field",
     "Split on comma; last token matched against India state list"),
    ("Battery", "Full address string with postal code",
     "Tokens matched against state list; postal codes stripped"),
    ("Green Steel", "'City, State' combined in City field",
     "Split on comma; last token matched against India state list"),
    ("Hydrogen", "'State - City' or 'City, State'",
     "Dash separator tried first; then comma; state matched against list"),
]
for j, row in enumerate(parsing_rows):
    add_data_row(parsing_tbl, row, shade=(j % 2 == 1))
set_col_widths(parsing_tbl, [3.5, 5.5, 7.0])
doc.add_paragraph()

# 2.4 State spelling
add_heading(doc, "2.4  State Spelling Corrections", 2)
add_bullet(doc, "\"Maharastra\" → Maharashtra")
add_bullet(doc, "\"Chhatisgarh\" → Chhattisgarh")
add_bullet(doc, "\"Orissa\" → Odisha (historical name)")
add_bullet(doc, "\"Tamilnadu\" / \"Tamil nadu\" → Tamil Nadu")
add_bullet(doc, "\"Andrapradesh\" → Andhra Pradesh")
doc.add_paragraph()

# 2.5 EV asset types
add_heading(doc, "2.5  EV Asset Type Grouping", 2)
add_para(doc, (
    "The EVs sheet had highly granular asset types. These were grouped into a "
    "controlled vocabulary to enable map filtering:"
))
ev_groups = [
    ("2-Wheeler", "scooter, motorcycle, e-bike, e-scooter, 2-wheeler"),
    ("3-Wheeler", "e-rickshaw, 3-wheeler, auto-rickshaw"),
    ("4-Wheeler / Car", "sedan, hatchback, SUV, e-car, passenger vehicle, 4-wheeler"),
    ("Bus / Commercial Vehicle", "e-bus, electric bus, truck, commercial vehicle"),
    ("Battery Pack / Component", "battery pack, cell, module, BMS"),
    ("Other", "anything not fitting above categories"),
]
for val, maps_from in ev_groups:
    add_bullet(doc, f"{val}: {maps_from}")
doc.add_paragraph()

# 2.6 Battery capacity
add_heading(doc, "2.6  Battery Capacity — Current vs. Total", 2)
add_para(doc, (
    "The Battery sheet had two separate capacity columns: 'Current capacity' and "
    "'Total capacity' (planned/expansion target). These were merged into a single "
    "'capacity' field with the following rule: if 'Current capacity' was populated, "
    "it was used as the primary value. If only 'Total capacity' was available, that "
    "was used. Where both exist, 'Total capacity' was moved to the 'notes' field. "
    "This ambiguity is flagged for colleague review in Section 3."
), italic=False)
doc.add_paragraph()

# 2.7 Hydrogen location naming
add_heading(doc, "2.7  Hydrogen Location — 'NA' City Values", 2)
add_para(doc, (
    "72 hydrogen records had no city or state information (city field contained 'NA', "
    "null, or placeholder text). Where a state was available only in the project name, "
    "it was extracted and used. Remaining records with truly unknown locations were "
    "geocoded to a central India default coordinate (19.22°N, 73.10°E — near Pune) "
    "as a placeholder. These are flagged in Section 3."
))
doc.add_paragraph()

doc.add_page_break()

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 3 — FLAGGED FOR COLLEAGUE REVIEW
# ─────────────────────────────────────────────────────────────────────────────

add_heading(doc, "3.  Flagged Items for Colleague Review", 1, color=(192, 0, 0))
add_para(doc, (
    "The items below require human judgement to resolve. They are grouped by issue type. "
    "For each, the automated assumption is stated — please override where incorrect."
), color=(192, 0, 0))
doc.add_paragraph()

# ─── 3.1 Unknown status ───────────────────────────────────────────────────────
add_heading(doc, "3.1  Unknown Status (203 records)", 2)
add_para(doc, (
    "These records had blank or unrecognised status values. They have been coded as "
    "'Unknown'. Please supply the correct status for each."
))
doc.add_paragraph()

# Group by sheet
status_flags_by_sheet = defaultdict(list)
for f in cleaning_flags:
    if f["field"] == "status":
        status_flags_by_sheet[f["sheet"]].append(f)

for s in SHEETS:
    flags_here = status_flags_by_sheet[s]
    if not flags_here:
        continue
    add_para(doc, f"{SHEET_LABELS[s]} — {len(flags_here)} records", bold=True, size=10)
    tbl = doc.add_table(rows=1, cols=3)
    tbl.style = "Table Grid"
    add_table_header(tbl, ["Row", "Company", "Original Status Value"], bg_hex="4F81BD")
    for j, f in enumerate(flags_here):
        orig = f["issue"].replace("Status blank or unrecognized: ", "")
        add_data_row(tbl, [f["row"], f["company"], orig], shade=(j % 2 == 1))
    set_col_widths(tbl, [1.2, 7.5, 5.0])
    doc.add_paragraph()

doc.add_page_break()

# ─── 3.2 Geocode fallbacks ────────────────────────────────────────────────────
add_heading(doc, "3.2  Geocode Fallbacks — State Capital Used (17 records)", 2)
add_para(doc, (
    "These records could not be geocoded at the city level (Nominatim returned no result). "
    "Their coordinates have been set to the relevant state capital. Please supply more "
    "precise coordinates if known."
))
doc.add_paragraph()

state_fallbacks = [f for f in geocode_fallbacks if "state capital" in f.get("action", "")]
tbl = doc.add_table(rows=1, cols=5)
tbl.style = "Table Grid"
add_table_header(tbl, ["Sheet", "Company", "City (attempted)", "State", "Action taken"])
for j, f in enumerate(state_fallbacks):
    add_data_row(tbl, [
        f["sheet"], f["company"][:45], f.get("city") or "(none)", f.get("state") or "(none)", f["action"]
    ], shade=(j % 2 == 1))
set_col_widths(tbl, [2.5, 4.5, 3.0, 2.5, 5.5])
doc.add_paragraph()

# ─── 3.3 Records with no location at all ──────────────────────────────────────
add_heading(doc, "3.3  Records with No City or State — Not Geocoded (94 records)", 2)
add_para(doc, (
    "These records have no location information whatsoever and were skipped by the "
    "geocoder. They will not appear on maps. The vast majority (94) are in the Solar "
    "Manufacturing sheet. Please add city/state if available."
))
doc.add_paragraph()

no_location = [f for f in geocode_fallbacks if f.get("issue", "").startswith("No city or state")]
by_sheet = Counter(f["sheet"] for f in no_location)
for s, cnt in by_sheet.most_common():
    add_bullet(doc, f"{SHEET_LABELS.get(s, s)}: {cnt} records without location")
doc.add_paragraph()

# Show solar_mfg no-location as a compact table
solar_no_loc = [f for f in no_location if f["sheet"] == "solar_mfg"]
if solar_no_loc:
    add_para(doc, "Solar Manufacturing — companies without location data:", bold=True, size=10)
    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = "Table Grid"
    add_table_header(tbl, ["Company", "Note"], bg_hex="4F81BD")
    for j, f in enumerate(solar_no_loc):
        add_data_row(tbl, [f["company"], "No city or state in source data"], shade=(j % 2 == 1))
    set_col_widths(tbl, [8.0, 8.5])
    doc.add_paragraph()

# ─── 3.4 Battery capacity ambiguity ──────────────────────────────────────────
add_heading(doc, "3.4  Battery Capacity — Current vs. Total Ambiguity", 2)
add_para(doc, (
    "Several battery records had both 'Current capacity' and 'Total capacity' values "
    "that differ. The 'capacity' field was set to 'Current capacity' where available; "
    "'Total capacity' was appended to notes. If the intended interpretation differs, "
    "please update accordingly."
))
battery_records = records_by_sheet["battery"]
ambig = [r for r in battery_records if r.get("notes") and "Total capacity" in (r.get("notes") or "")]
if ambig:
    doc.add_paragraph()
    tbl = doc.add_table(rows=1, cols=4)
    tbl.style = "Table Grid"
    add_table_header(tbl, ["Company", "City", "Capacity (used)", "Notes (contains total)"])
    for j, r in enumerate(ambig):
        add_data_row(tbl, [
            r.get("company", ""), r.get("city", ""),
            f"{r.get('capacity')} {r.get('capacity_units', '')}",
            (r.get("notes") or "")[:80]
        ], shade=(j % 2 == 1))
    set_col_widths(tbl, [4.5, 3.0, 3.5, 7.5])
doc.add_paragraph()

# ─── 3.5 Green steel / hydrogen — undisclosed locations ───────────────────────
add_heading(doc, "3.5  Green Steel — Record with No Location (1 record)", 2)
add_para(doc, (
    "POSCO and JSW Group (green_steel, id=1) has no city or state. "
    "The source article does not disclose a specific location. "
    "This record has no coordinates and will not appear on the map."
))
doc.add_paragraph()

# ─── 3.6 State parsing flags ──────────────────────────────────────────────────
state_flags = [f for f in cleaning_flags if f["field"] in ("state", "city/state")]
if state_flags:
    add_heading(doc, "3.6  Location Parsing — Ambiguous or Multi-Value Entries", 2)
    add_para(doc, (
        "These records had unusual location formats that required manual parsing decisions. "
        "Please verify the city and state assigned."
    ))
    tbl = doc.add_table(rows=1, cols=5)
    tbl.style = "Table Grid"
    add_table_header(tbl, ["Sheet", "Row", "Company", "Issue", "Assumed"])
    for j, f in enumerate(state_flags):
        add_data_row(tbl, [
            f["sheet"], f["row"], f["company"][:40], f["issue"][:60], f.get("assumed", "")
        ], shade=(j % 2 == 1))
    set_col_widths(tbl, [2.5, 1.2, 5.0, 6.0, 4.5])
    doc.add_paragraph()

doc.add_page_break()

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 4 — GEOCODING NOTES
# ─────────────────────────────────────────────────────────────────────────────

add_heading(doc, "4.  Geocoding Notes", 1)

add_para(doc, "Method: OpenStreetMap Nominatim API (free, no API key required).")
add_para(doc, "Query format: {City}, {State}, India — with countrycodes=in filter.")
add_para(doc, "Rate limit: 1 request per second (per Nominatim usage policy).")
doc.add_paragraph()

# Results table
add_heading(doc, "4.1  Results by Sheet", 2)
tbl = doc.add_table(rows=1, cols=6)
tbl.style = "Table Grid"
add_table_header(tbl, ["Sheet", "Total", "Pre-existing", "Geocoded OK", "State capital fallback", "Not geocodable"])

pre_existing = {"solar_mfg": 51, "evs": 55, "wind": 4, "battery": 0, "green_steel": 0, "hydrogen": 0}
state_cap_ct = Counter(f["sheet"] for f in geocode_fallbacks if "state capital" in f.get("action", ""))
no_loc_ct = Counter(f["sheet"] for f in geocode_fallbacks if f.get("issue", "").startswith("No city or state"))

for j, s in enumerate(SHEETS):
    recs = records_by_sheet[s]
    total = len(recs)
    pre = pre_existing[s]
    cap = state_cap_ct.get(s, 0)
    no_loc = no_loc_ct.get(s, 0)
    geocoded_ok = total - pre - cap - no_loc - coord_miss[s] + cap  # state cap counts as geocoded
    # Simpler: records with coords - pre-existing = new geocodes (including cap fallback)
    with_coords = sum(1 for r in recs if r.get("latitude") and r.get("longitude"))
    new_geocodes = with_coords - pre
    shade = (j % 2 == 1)
    add_data_row(tbl, [SHEET_LABELS[s], total, pre, new_geocodes - cap, cap, no_loc], shade=shade)

set_col_widths(tbl, [3.8, 1.5, 2.2, 2.2, 2.8, 2.5])
doc.add_paragraph()

add_heading(doc, "4.2  Fallback Coordinate", 2)
add_para(doc, (
    "When Nominatim returned no result for a city, coordinates were set to the relevant "
    "Indian state capital (see table in Appendix A). When neither city nor state was "
    "known, no coordinates were assigned (record will not appear on maps)."
))
doc.add_paragraph()

add_heading(doc, "4.3  'NA' Records in Hydrogen Sheet", 2)
add_para(doc, (
    "Many hydrogen records (especially those sourced from MoU registries) had no specific "
    "location — city was listed as 'NA'. Where a state could be inferred from the project "
    "name, state capital coordinates were used. Where no location could be inferred at all, "
    "a generic central-India default coordinate (19.22°N, 73.10°E) was assigned. "
    "These records should be treated as approximate and reviewed before final publication."
))
doc.add_paragraph()

doc.add_page_break()

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 5 — GIS PLATFORM RECOMMENDATIONS
# ─────────────────────────────────────────────────────────────────────────────

add_heading(doc, "5.  GIS Platform Recommendations", 1)
add_para(doc, "Based on the coordinate coverage achieved, we recommend the following layer structure:")
doc.add_paragraph()

layer_tbl = doc.add_table(rows=1, cols=4)
layer_tbl.style = "Table Grid"
add_table_header(layer_tbl, ["Layer", "Display Level", "Coverage", "Notes"])
layers = [
    ("Green Hydrogen", "Point (lat/lon)", "149/149 — 100%",
     "All records geocoded. 4 approximate (state capital). NA-only records use central-India placeholder."),
    ("Electric Vehicles", "Point (lat/lon)", "69/69 — 100%",
     "All records geocoded. 1 fallback (Toopor, Telangana)."),
    ("Wind Energy", "Point (lat/lon)", "27/36 — 75%",
     "9 records missing: 6 Suzlon with no location, 3 others. Will appear as state-level only."),
    ("Battery Storage", "Point (lat/lon)", "28/30 — 93%",
     "2 records not geocodable. 11 use state capital fallback."),
    ("Green Steel", "Point (lat/lon)", "3/4 — 75%",
     "1 record (POSCO/JSW) has no location. 1 uses Maharashtra capital."),
    ("Solar Manufacturing", "State choropleth (preferred)", "57/151 — 38%",
     "94 records (62%) have no city/state. Point layer possible for 57 records. "
     "State-level choropleth recommended as primary view."),
]
for j, row in enumerate(layers):
    add_data_row(layer_tbl, row, shade=(j % 2 == 1))
set_col_widths(layer_tbl, [3.5, 3.2, 3.0, 8.3])
doc.add_paragraph()

add_para(doc, (
    "Note: The solar manufacturing layer is best displayed as a state-level choropleth "
    "(number of manufacturers per state) rather than individual points, given the high "
    "proportion of records with missing location data. A point layer can be added as an "
    "optional overlay for the 57 records with known coordinates."
), italic=True, size=10)

doc.add_page_break()

# ─────────────────────────────────────────────────────────────────────────────
# APPENDIX A — STATE CAPITAL COORDINATES
# ─────────────────────────────────────────────────────────────────────────────

add_heading(doc, "Appendix A — State Capital Fallback Coordinates", 1)
add_para(doc, (
    "When a city could not be geocoded, coordinates were set to the corresponding "
    "state capital using the values below."
))
doc.add_paragraph()

STATE_CAPITALS = {
    "Andhra Pradesh": (15.9129, 79.7400),
    "Arunachal Pradesh": (27.0844, 93.6053),
    "Assam": (26.2006, 92.9376),
    "Bihar": (25.0961, 85.3131),
    "Chhattisgarh": (21.2787, 81.8661),
    "Goa": (15.2993, 74.1240),
    "Gujarat": (23.0225, 72.5714),
    "Haryana": (29.0588, 76.0856),
    "Himachal Pradesh": (31.1048, 77.1734),
    "Jharkhand": (23.6102, 85.2799),
    "Karnataka": (12.9716, 77.5946),
    "Kerala": (8.5241, 76.9366),
    "Madhya Pradesh": (23.2599, 77.4126),
    "Maharashtra": (19.0760, 72.8777),
    "Manipur": (24.6637, 93.9063),
    "Meghalaya": (25.5788, 91.8933),
    "Mizoram": (23.1645, 92.9376),
    "Nagaland": (25.6751, 94.1086),
    "Odisha": (20.2961, 85.8245),
    "Punjab": (30.7333, 76.7794),
    "Rajasthan": (26.9124, 75.7873),
    "Sikkim": (27.5330, 88.5122),
    "Tamil Nadu": (13.0827, 80.2707),
    "Telangana": (17.3850, 78.4867),
    "Tripura": (23.9408, 91.9882),
    "Uttar Pradesh": (26.8467, 80.9462),
    "Uttarakhand": (30.3165, 78.0322),
    "West Bengal": (22.5726, 88.3639),
    "Delhi": (28.6139, 77.2090),
    "Jammu And Kashmir": (34.0837, 74.7973),
    "Ladakh": (34.2996, 78.2932),
    "Puducherry": (11.9416, 79.8083),
    "Chandigarh": (30.7333, 76.7794),
}

cap_tbl = doc.add_table(rows=1, cols=3)
cap_tbl.style = "Table Grid"
add_table_header(cap_tbl, ["State / UT", "Latitude", "Longitude"])
for j, (state, (lat, lon)) in enumerate(sorted(STATE_CAPITALS.items())):
    add_data_row(cap_tbl, [state, f"{lat:.4f}", f"{lon:.4f}"], shade=(j % 2 == 1))
set_col_widths(cap_tbl, [5.5, 2.5, 2.5])

# ─────────────────────────────────────────────────────────────────────────────
# SAVE
# ─────────────────────────────────────────────────────────────────────────────

out_path = BASE / "India Green Assets Atlas - Project Overview.docx"
doc.save(out_path)
print(f"Saved: {out_path.name}")
print(f"  {len(all_records)} records | {len(cleaning_flags)} flags | {len(geocode_fallbacks)} geocode notes")
