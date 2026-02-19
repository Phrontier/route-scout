# Instrument Training Route Planner

Browser app for instrument students to generate and compare multiple training routes from a home airport.

## Flight Types

- `single`: fly to 1 airport and return (no stop), adjustable max leg distance up to 150 NM (default 100)
- `double`: fly to 1 airport, stop, and return, adjustable max leg distance up to 350 NM
- `triple`: route through 2 airports, with final return leg constrained by selected max leg distance

## Route Output

The app generates and ranks multiple route options (top 10). Selecting an option expands details inline on that card. For each option it shows:

- route length
- airfields in the route
- forecast surface winds (Open-Meteo 10m)
- likely runway in use with dimensions
- approach types likely available for that runway
- other/circling-capable approaches on non-primary runways
- runway mini view (no-scroll quick layout) plus full airport diagram link
- interactive quick route map (Leaflet + OpenStreetMap)

## Scoring (1-100)

Airport training score is weighted toward approach training value over perfect wind alignment:

- approach type diversity (max 45)
- approach quantity (max 20)
- wind/runway suitability (max 20)
- runway infrastructure (max 10)
- traffic/complexity value (max 5)

Route score is based on stop airport training scores plus route-level diversity and distance fit.
TACAN procedures remain in backend parsing but are excluded from T-6 scoring/output by default.

## Filters

- airport must have at least one runway `>= 4000 ft`
- contract fuel is **not** hard-filtered (manual verification required)

## Data Sources

- airports/runways: OurAirports CSV (with local fallback dataset if fetch fails)
- winds: Open-Meteo forecast/archive APIs, using `wind_speed_10m` in knots
- approaches: FAA d-TPP Metafile XML (official source, with AviationAPI fallback)
- airport diagrams/radar minimums: AviationAPI charts page parsing (General + Approach sections)

## Cloudflare Pages Deployment

This app uses a same-origin proxy at `/proxy` (implemented in `functions/proxy.js`) to fetch FAA/Aviation data without browser CORS proxy services.  
If deployed on Cloudflare Pages, ensure `functions/proxy.js` is included in your project.

## Local Dev Notes

- `python3 -m http.server 8080` serves static files only (no `/proxy` route).
- On localhost, the app now avoids `/proxy` for AviationAPI chart HTML and uses localhost-safe fallback fetch candidates.
- For full parity with Cloudflare Functions locally, run Pages dev (`wrangler pages dev .`) instead of static-only hosting.

## Versioning

The app uses semantic versioning with patch bumps for small bug fixes: `v0.0.x`.

- Update `/Users/conwaybolt/Documents/Codex/Instruments/version.json` `version` value for each deploy.
- Match the same patch value in asset query params inside `/Users/conwaybolt/Documents/Codex/Instruments/index.html` (for `style.css`, `app.js`, Bootstrap files) to force cache refresh on Cloudflare.
- The header ticker renders the live deployed value from `version.json`.

## Run

```bash
cd /Users/conwaybolt/Documents/Codex/Instruments
python3 -m http.server 8080
```

Open [http://localhost:8080](http://localhost:8080)
