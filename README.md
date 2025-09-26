# Ally Building Manager (Web Utility)

Lightweight admin utility for creating and visualizing building polygon footprints stored in Firestore for the main ALLY application.

## Files

- [ally-webdata/index.html](ally-webdata/index.html)
- [ally-webdata/style.css](ally-webdata/style.css)
- [ally-webdata/scripts/add-polygondata.js](ally-webdata/scripts/add-polygondata.js)

## Features

- Add/update building documents by name (document ID = building name)
- Define polygon via 3–4 latitude/longitude pairs (minimum 3)
- Automatic color + level metadata
- Renders polygons on an interactive Leaflet map
- Auto-fit map to all stored buildings
- Anonymous Firebase Auth session bootstrap

## Data Schema (Firestore)

Collection: `buildings`  
Document ID: `<buildingName>`

Fields:
- `color`: Hex string (e.g. `#6366F1`)
- `level`: Integer (floor or logical grouping)
- `polygons`: Array of GeoPoints (ordered vertices; first and last need not repeat)

## Coordinate Tips

- Order points either clockwise or counter‑clockwise for cleaner rendering.
- Avoid self-intersections.
- You can add more inputs by extending the form & script (currently 4 fields; only non-empty are used).
  
## Deployment

This is a static bundle—host on any static host (Firebase Hosting, S3, etc.). Ensure Firestore rules align with exposure risk.

## Disclaimer

Intended for internal administrative use supporting the ALLY app’s building geofence & spatial context features.

---
