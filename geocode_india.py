"""
geocode_india.py
India Green Assets Atlas — Geocoding Script
Adds lat/lon to all records missing coordinates using OpenStreetMap Nominatim.
Rate limited to 1 request/second per Nominatim usage policy.
"""

import json
import time
import urllib.request
import urllib.parse
from pathlib import Path

BASE = Path(__file__).parent

# Indian state capitals as fallback coordinates
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

geocode_cache = {}
fallback_records = []

def nominatim_geocode(city, state):
    """Query OSM Nominatim. Returns (lat, lon) or (None, None)."""
    query = ", ".join(filter(None, [city, state, "India"]))
    if query in geocode_cache:
        return geocode_cache[query]

    params = urllib.parse.urlencode({
        "q": query,
        "format": "json",
        "limit": 1,
        "countrycodes": "in",
    })
    url = f"https://nominatim.openstreetmap.org/search?{params}"
    headers = {"User-Agent": "IndiaGreenAssetsAtlas/1.0"}

    try:
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read())
        if data:
            lat = float(data[0]["lat"])
            lon = float(data[0]["lon"])
            geocode_cache[query] = (lat, lon)
            return lat, lon
    except Exception as e:
        print(f"    [ERROR] {query}: {e}")

    geocode_cache[query] = (None, None)
    return None, None

def geocode_sheet(records, sheet_name):
    need_geocode = [r for r in records if not r.get("latitude") or not r.get("longitude")]
    already_have = len(records) - len(need_geocode)
    print(f"  {sheet_name}: {already_have} already have coords, geocoding {len(need_geocode)}...")

    for i, rec in enumerate(need_geocode):
        city = rec.get("city")
        state = rec.get("state")
        company = rec.get("company", "?")

        if not city and not state:
            print(f"    [{i+1}/{len(need_geocode)}] SKIP (no city/state): {company}")
            fallback_records.append({
                "sheet": sheet_name, "company": company,
                "city": city, "state": state,
                "issue": "No city or state — cannot geocode",
                "action": "Coordinates left null"
            })
            continue

        lat, lon = nominatim_geocode(city, state)
        time.sleep(1.1)  # Nominatim rate limit: 1 req/sec

        if lat and lon:
            rec["latitude"] = lat
            rec["longitude"] = lon
            rec["geocoded"] = True
            print(f"    [{i+1}/{len(need_geocode)}] ✓ {company} ({city}, {state}) → {lat:.4f}, {lon:.4f}")
        else:
            # Fallback: state capital
            cap = STATE_CAPITALS.get(state or "")
            if cap:
                rec["latitude"] = cap[0]
                rec["longitude"] = cap[1]
                rec["geocoded"] = True
                rec["geocode_note"] = f"Geocoded to {state} capital (city '{city}' not found)"
                print(f"    [{i+1}/{len(need_geocode)}] ↩ {company}: '{city}' not found → {state} capital")
                fallback_records.append({
                    "sheet": sheet_name, "company": company,
                    "city": city, "state": state,
                    "issue": f"City '{city}' not found by Nominatim",
                    "action": f"Coordinates set to {state} state capital ({cap[0]:.4f}, {cap[1]:.4f})"
                })
            else:
                print(f"    [{i+1}/{len(need_geocode)}] ✗ {company}: no city, no state capital fallback")
                fallback_records.append({
                    "sheet": sheet_name, "company": company,
                    "city": city, "state": state,
                    "issue": "City not found and state unknown — no fallback available",
                    "action": "Coordinates left null"
                })

    return records

def main():
    sheets = ["solar_mfg", "evs", "wind", "battery", "green_steel", "hydrogen"]

    for sheet_name in sheets:
        path = BASE / f"{sheet_name}.json"
        with open(path) as f:
            records = json.load(f)

        records = geocode_sheet(records, sheet_name)

        with open(path, "w", encoding="utf-8") as f:
            json.dump(records, f, indent=2, ensure_ascii=False)

    # Save fallback log
    fallback_path = BASE / "_geocode_fallbacks.json"
    with open(fallback_path, "w", encoding="utf-8") as f:
        json.dump(fallback_records, f, indent=2, ensure_ascii=False)

    # Summary
    total = sum(1 for s in sheets for r in json.load(open(BASE / f"{s}.json")) if r.get("latitude"))
    total_all = sum(len(json.load(open(BASE / f"{s}.json"))) for s in sheets)
    print(f"\nGeocoding complete.")
    print(f"Records with coordinates: {total}/{total_all}")
    print(f"Fallback/flagged records: {len(fallback_records)}")
    print(f"Fallbacks saved to: {fallback_path.name}")

if __name__ == "__main__":
    main()
