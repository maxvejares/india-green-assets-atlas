"""
clean_india_data.py
India Green Assets Atlas — Data Cleaning Script
Reads India Industrial Atlas Excel, cleans all 6 populated sheets, outputs JSON files.
"""

import json
import re
import openpyxl
from pathlib import Path

BASE = Path(__file__).parent
EXCEL = BASE / "India Industrial Atlas. Master Sheet.xlsx"

# ── Status normalization ──────────────────────────────────────────────────────
STATUS_MAP = {
    # Operational variants
    "commissioned": "Operational",
    "open": "Operational",
    "operational": "Operational",
    "active": "Operational",
    # Announced variants
    "announced": "Announced",
    "announced in 2022": "Announced",
    "announced in 2023": "Announced",
    "announced in 2024": "Announced",
    "announced in 2025": "Announced",
    "mou signed": "Announced",
    "mou signed in 2023": "Announced",
    "mou signed in 2024": "Announced",
    # Under Construction
    "under construction": "Under Construction",
    "under-construction": "Under Construction",
    # Planned
    "planned": "Planned",
    # Closed/Decommissioned
    "closed": "Closed",
    "decommissioned": "Closed",
    # Expansion
    "expansion announced": "Announced",
    "expansion": "Announced",
}

def normalize_status(raw):
    if not raw:
        return "Unknown"
    s = str(raw).strip().lower()
    for key, val in STATUS_MAP.items():
        if key in s:
            return val
    return "Unknown"

# ── EV Asset Type normalization ───────────────────────────────────────────────
def normalize_ev_asset(raw):
    if not raw:
        return "Other"
    s = str(raw).lower()
    if any(x in s for x in ["2-wheel", "2 wheel", "e-bike", "scooter", "motorcycle", "moped"]):
        return "2-Wheeler"
    if any(x in s for x in ["3-wheel", "3 wheel", "auto", "rickshaw", "cargo"]):
        return "3-Wheeler"
    if any(x in s for x in ["bus", "truck", "commercial", "lcv", "hcv"]):
        return "Bus/Commercial"
    if any(x in s for x in ["car", "suv", "sedan", "hatchback", "4-wheel", "4 wheel", "passenger"]):
        return "4-Wheeler/Car"
    if any(x in s for x in ["battery", "cell", "pack"]):
        return "Battery Pack"
    return "Other"

# ── Location parsing ──────────────────────────────────────────────────────────
INDIA_STATES = {
    "andhra pradesh", "arunachal pradesh", "assam", "bihar", "chhattisgarh",
    "goa", "gujarat", "haryana", "himachal pradesh", "jharkhand", "karnataka",
    "kerala", "madhya pradesh", "maharashtra", "manipur", "meghalaya", "mizoram",
    "nagaland", "odisha", "punjab", "rajasthan", "sikkim", "tamil nadu",
    "telangana", "tripura", "uttar pradesh", "uttarakhand", "west bengal",
    "delhi", "jammu and kashmir", "ladakh", "puducherry", "chandigarh",
    "andaman and nicobar", "dadra and nagar haveli", "lakshadweep"
}

STATE_SPELLING_FIX = {
    "maharastra": "Maharashtra",
    "maharastra": "Maharashtra",
    "chhatisgarh": "Chhattisgarh",
    "chattisgarh": "Chhattisgarh",
    "orissa": "Odisha",
    "uttaranchal": "Uttarakhand",
    "pondicherry": "Puducherry",
    "tamilnadu": "Tamil Nadu",
    "tamil nadu": "Tamil Nadu",
}

def fix_state_spelling(state):
    if not state:
        return state
    s = state.strip()
    lower = s.lower()
    if lower in STATE_SPELLING_FIX:
        return STATE_SPELLING_FIX[lower]
    return s.title()

def parse_city_state_combined(raw):
    """Parse 'City, State' or 'State - City' or 'City, State, PostalCode' formats."""
    if not raw:
        return None, None
    s = str(raw).strip()

    # Format: "State - City" (hydrogen pattern)
    if " - " in s:
        parts = s.split(" - ", 1)
        # Check which side looks like a state
        left_lower = parts[0].strip().lower()
        if left_lower in INDIA_STATES or any(st in left_lower for st in INDIA_STATES):
            state = fix_state_spelling(parts[0].strip())
            city = parts[1].strip()
            # Remove postal codes from city
            city = re.sub(r',?\s*\d{6}', '', city).strip()
            return city, state
        else:
            city = parts[0].strip()
            state = fix_state_spelling(parts[1].strip())
            return city, state

    # Format: "City, State" or "City, District, State, PostalCode"
    parts = [p.strip() for p in s.split(",")]
    # Strip postal codes
    parts = [p for p in parts if not re.match(r'^\d{5,6}$', p)]
    # Strip known industrial area keywords mixed in
    parts = [p for p in parts if p]

    if len(parts) == 1:
        # Could be just state name or just city
        lower = parts[0].lower()
        if lower in INDIA_STATES:
            return None, fix_state_spelling(parts[0])
        return parts[0], None
    elif len(parts) == 2:
        # "City, State"
        state_lower = parts[1].lower()
        if any(st in state_lower for st in INDIA_STATES) or state_lower in INDIA_STATES:
            return parts[0], fix_state_spelling(parts[1])
        # Maybe it's "State, City" — check first part
        state_lower0 = parts[0].lower()
        if any(st in state_lower0 for st in INDIA_STATES) or state_lower0 in INDIA_STATES:
            return parts[1], fix_state_spelling(parts[0])
        return parts[0], parts[1]
    elif len(parts) >= 3:
        # Try to find which part is a known state
        for i, p in enumerate(parts):
            if p.lower() in INDIA_STATES or any(st in p.lower() for st in INDIA_STATES):
                state = fix_state_spelling(p)
                # City is usually the first non-state part
                city_parts = [parts[j] for j in range(len(parts)) if j != i]
                city = city_parts[0] if city_parts else parts[0]
                return city, state
        # Fallback: first is city, last is state
        return parts[0], fix_state_spelling(parts[-1])

    return s, None

def parse_coordinates(raw):
    """Parse 'lat, lon' string into (lat, lon) floats."""
    if not raw:
        return None, None
    s = str(raw).strip()
    parts = s.split(",")
    if len(parts) == 2:
        try:
            lat = float(parts[0].strip())
            lon = float(parts[1].strip())
            if -90 <= lat <= 90 and -180 <= lon <= 180:
                return lat, lon
        except ValueError:
            pass
    return None, None

def clean_float(val):
    if val is None:
        return None
    try:
        return float(str(val).replace(",", "").strip())
    except (ValueError, TypeError):
        return None

def v(val):
    """Return None if empty/whitespace, else stripped string."""
    if val is None:
        return None
    s = str(val).strip()
    return s if s else None

# ── Sheet readers ─────────────────────────────────────────────────────────────

def read_sheet(wb, name):
    ws = wb[name]
    headers = [cell.value for cell in ws[1]]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        d = dict(zip(headers, row))
        if any(v(val) for key, val in d.items() if key and 'region' not in str(key).lower()):
            rows.append(d)
    return rows

def clean_solar_mfg(wb):
    rows = read_sheet(wb, "solar mfg")
    flags = []
    out = []
    for i, r in enumerate(rows):
        # Coordinates
        coord_raw = v(r.get("Coordinates")) or v(r.get("Coordinates (lat, long)"))
        lat, lon = parse_coordinates(coord_raw)
        if lat is None:
            lat_raw = v(r.get("Lattitude")) or v(r.get("Latitude"))
            lon_raw = v(r.get("Longitude"))
            lat = clean_float(lat_raw)
            lon = clean_float(lon_raw)

        city = v(r.get("City"))
        state = fix_state_spelling(v(r.get("State/Province")) or v(r.get("State")) or "")

        status_raw = v(r.get("Status"))
        status = normalize_status(status_raw)
        if status == "Unknown":
            flags.append({"sheet": "solar_mfg", "row": i+2, "company": v(r.get("Company")), "field": "status", "issue": f"Status blank or unrecognized: '{status_raw}'", "assumed": "Unknown"})

        cap = clean_float(r.get("Capacity (2023)") or r.get("Capacity"))
        cap27 = clean_float(r.get("Capacity (2027)"))

        rec = {
            "id": i + 1,
            "sheet": "solar_mfg",
            "company": v(r.get("Company")),
            "city": city,
            "state": state if state else None,
            "asset_type": v(r.get("Asset Type")),
            "capacity": cap,
            "capacity_2027": cap27,
            "capacity_units": v(r.get("Capacity Units")),
            "status": status,
            "operational_date": v(r.get("Operational date")),
            "capital_investment": v(r.get("Capital Investment (BRL)") or r.get("Capital Investment")),
            "source": v(r.get("Source")),
            "notes": v(r.get("Note") or r.get("Notes")),
            "latitude": lat,
            "longitude": lon,
            "geocoded": False,
        }
        out.append(rec)
    return out, flags

def clean_evs(wb):
    rows = read_sheet(wb, "EVs")
    flags = []
    out = []
    for i, r in enumerate(rows):
        coord_raw = v(r.get("Coordinates")) or v(r.get("Coordinates (lat, long)"))
        lat, lon = parse_coordinates(coord_raw)
        if lat is None:
            lat = clean_float(v(r.get("Lattitude")) or v(r.get("Latitude")))
            lon = clean_float(v(r.get("Longitude")))

        status_raw = v(r.get("Status"))
        status = normalize_status(status_raw)
        if status == "Unknown":
            flags.append({"sheet": "evs", "row": i+2, "company": v(r.get("Company")), "field": "status", "issue": f"Status blank: '{status_raw}'", "assumed": "Unknown — please verify"})

        asset_raw = v(r.get("Asset Type"))
        asset_clean = normalize_ev_asset(asset_raw)

        rec = {
            "id": i + 1,
            "sheet": "evs",
            "company": v(r.get("Company")),
            "city": v(r.get("City")),
            "state": fix_state_spelling(v(r.get("State")) or ""),
            "asset_type": asset_clean,
            "asset_type_original": asset_raw,
            "capacity": clean_float(r.get("Capacity")),
            "capacity_units": "vehicles/year",
            "status": status,
            "operational_date": v(r.get("Operational date")),
            "capital_investment": v(r.get("Capital Investment (BRL)") or r.get("Capital Investment")),
            "fdi_source_country": v(r.get("Source Country of Investment (If FDI)")),
            "direct_employment": clean_float(r.get("Direct Employment")),
            "source": v(r.get("Source")),
            "notes": v(r.get("Notes")),
            "latitude": lat,
            "longitude": lon,
            "geocoded": False,
        }
        out.append(rec)
    return out, flags

def clean_wind(wb):
    rows = read_sheet(wb, "wind")
    flags = []
    out = []
    for i, r in enumerate(rows):
        lat = clean_float(v(r.get("Lattitude")) or v(r.get("Latitude")))
        lon = clean_float(v(r.get("Longitude")))

        # City field contains "City, State"
        city_raw = v(r.get("City"))
        city, state = parse_city_state_combined(city_raw)
        if not state:
            flags.append({"sheet": "wind", "row": i+2, "company": v(r.get("Company")), "field": "state", "issue": f"Could not extract state from city field: '{city_raw}'", "assumed": "None — needs manual review"})

        status_raw = v(r.get("Status"))
        status = normalize_status(status_raw)
        if status == "Unknown":
            flags.append({"sheet": "wind", "row": i+2, "company": v(r.get("Company")), "field": "status", "issue": f"Status blank: '{status_raw}'", "assumed": "Unknown"})

        rec = {
            "id": i + 1,
            "sheet": "wind",
            "company": v(r.get("Company")),
            "city": city,
            "state": state,
            "asset_type": v(r.get("Asset Type")),
            "capacity": clean_float(r.get("Capacity")),
            "capacity_units": v(r.get("Capacity Units")),
            "status": status,
            "year_established": v(r.get("Year established / operational")),
            "capital_investment": v(r.get("Capital Investment (BRL)") or r.get("Capital Investment")),
            "fdi_source_country": v(r.get("Source Country of Investment (If FDI)")),
            "direct_employment": clean_float(r.get("Direct Employment")),
            "source": v(r.get("Source")),
            "notes": v(r.get("Notes")),
            "latitude": lat,
            "longitude": lon,
            "geocoded": False,
        }
        out.append(rec)
    return out, flags

def clean_battery(wb):
    rows = read_sheet(wb, "battery")
    flags = []
    out = []
    for i, r in enumerate(rows):
        lat = clean_float(v(r.get("Lattitude")) or v(r.get("Latitude")))
        lon = clean_float(v(r.get("Longitude")))

        city_raw = v(r.get("City"))
        city, state = parse_city_state_combined(city_raw)
        if not state:
            flags.append({"sheet": "battery", "row": i+2, "company": v(r.get("Company")), "field": "state", "issue": f"Could not extract state from city field: '{city_raw}'", "assumed": "None"})

        status_raw = v(r.get("Status"))
        status = normalize_status(status_raw)
        if status == "Unknown":
            flags.append({"sheet": "battery", "row": i+2, "company": v(r.get("Company")), "field": "status", "issue": f"Status blank: '{status_raw}'", "assumed": "Unknown"})

        # Capacity: prefer total_capacity, fall back to current_capacity; flag ambiguity
        cur_cap = clean_float(r.get("Current capacity"))
        tot_cap = clean_float(r.get("Total capacity"))
        if cur_cap and tot_cap and cur_cap != tot_cap:
            flags.append({"sheet": "battery", "row": i+2, "company": v(r.get("Company")), "field": "capacity", "issue": f"Both current ({cur_cap}) and total ({tot_cap}) capacity exist", "assumed": f"Used total capacity ({tot_cap})"})
        capacity = tot_cap if tot_cap else cur_cap

        rec = {
            "id": i + 1,
            "sheet": "battery",
            "company": v(r.get("Company")),
            "city": city,
            "state": state,
            "asset_type": v(r.get("Asset Type")),
            "capacity": capacity,
            "capacity_current": cur_cap,
            "capacity_total": tot_cap,
            "capacity_units": v(r.get("Capacity Units")),
            "status": status,
            "operational_date": v(r.get("Operational date")),
            "capital_investment": v(r.get("Capital Investment")),
            "fdi_source_country": v(r.get("Source Country of Investment (If FDI)")),
            "direct_employment": clean_float(r.get("Direct Employment")),
            "source": v(r.get("Source")),
            "notes": v(r.get("Notes")),
            "latitude": lat,
            "longitude": lon,
            "geocoded": False,
        }
        out.append(rec)
    return out, flags

def clean_green_steel(wb):
    rows = read_sheet(wb, "green steel and aluminum")
    flags = []
    out = []
    for i, r in enumerate(rows):
        lat = clean_float(v(r.get("Lattitude")) or v(r.get("Latitude")))
        lon = clean_float(v(r.get("Longitude")))

        city_raw = v(r.get("City"))
        if city_raw and city_raw.lower() == "undisclosed":
            city, state = None, None
            flags.append({"sheet": "green_steel", "row": i+2, "company": v(r.get("Company")), "field": "city/state", "issue": "Location listed as 'Undisclosed'", "assumed": "No coordinates assigned"})
        else:
            city, state = parse_city_state_combined(city_raw)

        status_raw = v(r.get("Status"))
        # Preserve year from status like "Announced in 2022"
        year_match = re.search(r'\b(20\d\d)\b', str(status_raw or ""))
        status = normalize_status(status_raw)
        notes = v(r.get("Notes")) or ""
        if year_match:
            notes = (notes + f" [Announcement year: {year_match.group(1)}]").strip()

        rec = {
            "id": i + 1,
            "sheet": "green_steel",
            "company": v(r.get("Company")),
            "city": city,
            "state": state,
            "asset_type": v(r.get("Asset Type")),
            "capacity": clean_float(r.get("Capacity")),
            "capacity_units": v(r.get("Capacity Units")),
            "status": status,
            "operational_date": v(r.get("Operational Date") or r.get("Operational date")),
            "capital_investment": v(r.get("Capital Investment")),
            "fdi_source_country": v(r.get("Source Country of Investment (If FDI)")),
            "direct_employment": clean_float(r.get("Direct Employment")),
            "source": v(r.get("Source")),
            "notes": notes if notes else None,
            "latitude": lat,
            "longitude": lon,
            "geocoded": False,
        }
        out.append(rec)
    return out, flags

def clean_hydrogen(wb):
    rows = read_sheet(wb, "hydrogen")
    flags = []
    out = []
    for i, r in enumerate(rows):
        lat = clean_float(v(r.get("Lattitude")) or v(r.get("Latitude")))
        lon = clean_float(v(r.get("Longitude")))

        city_raw = v(r.get("City"))
        city, state = parse_city_state_combined(city_raw)
        if not state:
            flags.append({"sheet": "hydrogen", "row": i+2, "company": v(r.get("Company")), "field": "state", "issue": f"Could not extract state from: '{city_raw}'", "assumed": "None"})

        status_raw = v(r.get("Status"))
        year_match = re.search(r'\b(20\d\d)\b', str(status_raw or ""))
        status = normalize_status(status_raw)
        notes = v(r.get("Notes")) or ""
        if year_match and "announced" in str(status_raw or "").lower():
            notes = (notes + f" [Announcement year: {year_match.group(1)}]").strip()

        rec = {
            "id": i + 1,
            "sheet": "hydrogen",
            "company": v(r.get("Company")),
            "city": city,
            "state": state,
            "asset_type": v(r.get("Asset Type")),
            "capacity": clean_float(r.get("Capacity")),
            "capacity_units": v(r.get("Capacity Units")),
            "status": status,
            "operational_date": v(r.get("Operational date")),
            "capital_investment": v(r.get("Capital Investment")),
            "fdi_source_country": v(r.get("Source Country of Investment (If FDI)")),
            "direct_employment": clean_float(r.get("Direct Employment")),
            "source": v(r.get("Source")),
            "notes": notes if notes else None,
            "latitude": lat,
            "longitude": lon,
            "geocoded": False,
        }
        out.append(rec)
    return out, flags

# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print(f"Loading: {EXCEL}")
    wb = openpyxl.load_workbook(EXCEL, data_only=True)

    all_flags = []
    sheets = {
        "solar_mfg": clean_solar_mfg,
        "evs": clean_evs,
        "wind": clean_wind,
        "battery": clean_battery,
        "green_steel": clean_green_steel,
        "hydrogen": clean_hydrogen,
    }

    results = {}
    for name, fn in sheets.items():
        print(f"  Cleaning {name}...", end=" ")
        data, flags = fn(wb)
        results[name] = data
        all_flags.extend(flags)

        out_path = BASE / f"{name}.json"
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        print(f"{len(data)} records → {out_path.name}")

    # Save flags for use by overview doc script
    flags_path = BASE / "_cleaning_flags.json"
    with open(flags_path, "w", encoding="utf-8") as f:
        json.dump(all_flags, f, indent=2, ensure_ascii=False)

    print(f"\nTotal records: {sum(len(v) for v in results.values())}")
    print(f"Total flags: {len(all_flags)}")
    print(f"Flags saved to: {flags_path.name}")

if __name__ == "__main__":
    main()
