[README.md](https://github.com/user-attachments/files/25534497/README.md)
# India Green Assets Atlas

An interactive GIS platform mapping green industrial assets across India — solar manufacturing, electric vehicles, wind energy, battery storage, green steel, and green hydrogen.

Built by the **Net Zero Industrial Policy Lab (NZIPL)** at Johns Hopkins University

---

## Overview

The atlas combines geocoded facility-level data with state-level choropleth visualisation, enabling both site-specific and regional analysis of India's green industrial landscape.

**439 records across 6 sectors · 333 mapped facilities · 28 states covered**

| Sector | Records | Mapped |
|---|---|---|
| Solar Manufacturing | 151 | 57 |
| Green Hydrogen | 149 | 149 |
| Electric Vehicles | 69 | 69 |
| Wind Energy | 36 | 27 |
| Battery Storage | 30 | 28 |
| Green Steel | 4 | 3 |

> **Note on solar coverage:** 94 of 151 solar records had no city or state in the source data, so they could not be geocoded. The remaining 57 have verified coordinates.

---

## Usage

Open `index.html` directly in a browser — no server or installation needed.

```
open index.html
```

All data is embedded as JavaScript files so the platform works from the local filesystem (`file://`).

---

## Features

- **Point markers** — facility-level locations with company popups (status, capacity, investment, FDI source)
- **State choropleth** — heatmap showing asset density per state, selectable by sector via dropdown
- **Three view modes** — Points + States / Points only / States only
- **Layer toggles** — enable/disable any of the 6 sectors independently
- **Status filter** — filter across all layers by Operational / Announced / Under Construction / Planned / Unknown
- **Search** — full-text search across companies, cities, and states
- **Cluster/expand** — markers cluster at low zoom and expand at zoom 9+

---

## Data

### Sources
Raw data compiled from public sources into `India Industrial Atlas. Master Sheet.xlsx` (7 sheets).

### Cleaning
Script: `clean_india_data.py`

- Standardised `Status` to a controlled vocabulary: Operational, Announced, Under Construction, Planned, Unknown
- Split combined `Coordinates` and `City, State` fields into separate columns
- Normalised EV asset types, battery capacity fields, and hydrogen status variants
- Dropped near-empty columns (`Name`, `Region`)

Full cleaning decisions and flagged records are documented in `India Green Assets Atlas - Project Overview.docx`.

### Geocoding
Script: `geocode_india.py`

- Used OpenStreetMap Nominatim API (`{city}, {state}, India`) at 1 req/sec
- Records with pre-existing coordinates were kept as-is
- Unresolved city names fell back to state capital coordinates (17 records flagged)
- 123 geocode notes saved to `_geocode_fallbacks.json`

### Output files

| File | Description |
|---|---|
| `solar_mfg.json` | Solar manufacturing facilities |
| `evs.json` | EV manufacturing plants |
| `wind.json` | Wind component manufacturers |
| `battery.json` | Battery storage facilities |
| `green_steel.json` | Green steel plants |
| `hydrogen.json` | Green hydrogen projects |
| `india_state_counts.js` | State-level asset counts per sector |
| `india_states_embedded.js` | India state polygons (GeoJSON, embedded as JS) |
| `india_embedded_data.js` | All 6 datasets embedded as JS variables |
| `_cleaning_flags.json` | 281 cleaning decisions logged |
| `_geocode_fallbacks.json` | 123 geocoding fallback/skip records |

---

## File Structure

```
india-green-assets-atlas/
├── index.html                    # GIS platform (open this)
├── india_embedded_data.js        # All 6 datasets as JS variables
├── india_state_counts.js         # State-level counts + maxima
├── india_states_embedded.js      # India states GeoJSON (7 MB)
├── solar_mfg.json
├── evs.json
├── wind.json
├── battery.json
├── green_steel.json
├── hydrogen.json
├── clean_india_data.py           # Data cleaning script
├── geocode_india.py              # Geocoding script
├── create_india_overview.py      # Word doc generation script
├── _cleaning_flags.json          # Cleaning audit log
├── _geocode_fallbacks.json       # Geocoding audit log
├── India Green Assets Atlas - Project Overview.docx
└── India Industrial Atlas. Master Sheet.xlsx
```

---

## Limitations & Known Issues

- **Solar data gap:** 94 solar records have no location in the source data and do not appear on the map
- **State-capital fallbacks:** 17 records with unresolved city names are geocoded to their state capital — marked with a warning in popups
- **EV status:** 64 EV records have unknown status in the source data
- **Green steel:** Only 4 records; one has undisclosed location

---

## Related Work

- [Mexico Green Industrial Policy Observatory](https://github.com/NZIPL) — subnational policy tracking for Mexico

---

*Net Zero Industrial Policy Lab · Johns Hopkins University · 2026*
