# Baja Roosterfish Guide — Build Specification v29
**Last updated:** May 2026  
**File:** `baja_roosterfish_guide.html`  
**Target size:** ~490–520KB (includes inlined Leaflet JS + CSS)  
**Google Drive folder:** `Baja_Roosterfish_Guide_2026` (ID: `1F8zPDaDUvRzadlNSRHNtWNVjJZU4Q76t`)

---

## CRITICAL: JavaScript Architecture

The JS must be structured in this exact order or the map and buttons will break:

```
1. <style> Leaflet CSS inlined </style>
2. <style> Our custom CSS </style>
3. [page body / HTML content]
4. <script> Leaflet JS inlined (exports to globalThis.leaflet) </script>
5. <script>
     // L alias — MUST run before icon fix and map init
     if (typeof leaflet !== 'undefined' && typeof L === 'undefined') { window.L = leaflet; }
     // Icon fix — uses L, so must be AFTER alias
     delete L.Icon.Default.prototype._getIconUrl;
     L.Icon.Default.mergeOptions({ iconUrl: '...base64...', ... });
   </script>
6. <script>
     // GLOBAL functions — must be global (not in closure) for onclick= attrs
     function toggleLayer(key, btn) { ... }   // uses map, layerMap globals
     function flyToBeach(num) { ... }          // uses map, beaches globals
     function scrollToDiag(diagId) { ... }
     
     // Global variable declarations — assigned inside window.load
     var map, layerMap, layerBeaches, layerShops, layerRoads,
         layerZones, layerFood, layerAirbnb, beaches;
     
     // Map init — inside window.load so Leaflet is guaranteed ready
     window.addEventListener('load', function() {
       // ═══ LEAFLET MAP INITIALIZATION ═══
       map = L.map('map', { tap:true, tapTolerance:15, touchZoom:true,
                            dragging:true, scrollWheelZoom:false, zoomControl:true });
       // NOTE: NO 'var' on map, layerBeaches, layerMap, beaches etc —
       // they must write to the global scope declared above
       beaches = [ ... ];
       layerBeaches = L.layerGroup(); ...
       layerMap = { beaches: layerBeaches, ... };
       // invalidateSize for mobile
       setTimeout(function(){ map.invalidateSize(true); }, 300);
     }); // end window load
   </script>
```

### Why this structure is required
- **Leaflet inlined:** Prevents cross-origin "Script error" in Chrome when opening file locally
- **L alias:** npm Leaflet exports to `globalThis.leaflet` not `window.L`
- **Global functions:** `onclick="toggleLayer(...)"` requires global scope; closures break this
- **Global var decls + no-var assignments:** `toggleLayer` reads `map`/`layerMap` at call time — they must be in global scope
- **window.load wrapper:** Guarantees Leaflet JS is fully parsed before map init runs (fixes iOS Safari race condition)

---

## Viewport Meta (required for mobile)
```html
<meta name="viewport" content="width=device-width, initial-scale=1.0">
```
Do NOT add `maximum-scale=1.0` — it blocks map pinch-zoom.

---

## Inlining Leaflet (to avoid CDN issues)

```bash
npm install leaflet@1.9.4
# JS: node_modules/leaflet/dist/leaflet.js  (~147KB)
# CSS: node_modules/leaflet/dist/leaflet.css (~14KB)
```

CSS image URLs must be replaced with base64 data URIs:
```python
import base64
# Replace: url(images/marker-icon.png)
# With:    url(data:image/png;base64,<base64 of file>)
# Files: marker-icon.png, marker-icon-2x.png, marker-shadow.png, layers.png, layers-2x.png
```

---

## Document Structure (section order)

1. `<head>` — viewport, Google Fonts CDN, inlined Leaflet CSS, custom CSS
2. Legend bar (`.legend-bar`)
3. Moon bar — June 2026 trip week (`.moon-bar`, `class="moon-days"`)
4. Interactive Leaflet map (`id="map"`) + layer toggle buttons
5. `<!-- BEACH CARDS -->` — 12 beach cards, Zones 1–4
6. Tackle shops (4 cards)
7. ATV/UTV rentals (6 cards)
8. Beach rankings (`id="beach-rankings"`)
9. Pacific explainer (`.pe-sec`)
10. Quick reference card (`.qref`)
11. `<!-- ATV DIRECTIONS SECTION -->` (`id="atv-directions"`)
    - Vehicle guide table (Car vs. ATV)
    - Dir overview boxes (4)
    - 6 route cards, each with Google Maps iframe + nav buttons
12. Casting diagrams — 5 SVGs (`.dgrid`)
13. Lure guide (`.lure-grid`) — 8 lure cards
14. Rigging guide (`.rig-grid`) — 5 scenarios
15. Food/restaurants (`.food-grid`) — 9 cards
16. `<!-- JUNE CONDITIONS CALENDAR -->` (`id="june-conditions"`)
17. `<!-- SURF CONDITIONS PER BEACH -->` (`id="surf-conditions"`)
18. `<!-- PANGA FISHING SECTION -->` (`id="panga-guide"`)
19. `<!-- NORTH CORRIDOR ROOSTERFISH -->` (`id="north-corridor"`)
20. `<footer>`
21. Inlined Leaflet JS + L alias + icon fix
22. Global functions + map init script

---

## CSS Variables
```css
--sand: #f5efe6;  --deep-sea: #0d3b4f;  --cortez: #1a6b8a;
--rooster-red: #c0392b;  --gold: #d4a017;  --dusk: #8b6914;
--foam: #e8f4f8;  --text: #3a2e22;  --muted: #6b5c4e;
--border: #e0d5c5;  --atv: #8b4500;
--green: #1a7a5e;  --blue: #2d5986;
```

## Fonts (Google Fonts CDN)
- Playfair Display (headers)
- IBM Plex Mono (labels, mono)
- Libre Baskerville (body)

---

## Beach Data (12 beaches)
See `baja_fishing_data.json` for full beach array. Key structure:
```js
{n:'01', name:'El Cardonal Beach', lat:23.8439, lng:-109.7441,
 zone:1, species:'...', time:'Dawn 6–9:30am ★★★★★',
 tide:'...', id:'beach-01', icon:fishIcon}
```
Beach card IDs: `id="beach-1"` through `id="beach-12"` (NOT zero-padded)
JS beach array uses zero-padded n: `n:'01'` through `n:'12'`

---

## Map Configuration
- Center: `[23.3, -109.65]`, zoom 8
- Layer groups: `layerBeaches`, `layerShops`, `layerRoads`, `layerZones`, `layerFood`, `layerAirbnb`
- `layerMap` object maps string keys to layer groups for `toggleLayer()`
- All 6 layers added to map on init; buttons toggle them

## Layer Button CSS
```css
.layer-btn { ... }
.layer-btn.active, .layer-btn.on { background:rgba(255,255,255,.9); color:var(--deep-sea); }
```
Buttons call: `onclick="toggleLayer('beaches', this)"`  ← key first, this second

---

## ATV Route Navigation Blocks
Each of the 6 route cards contains:
1. Google Maps iframe embed: `maps.google.com/maps?saddr=LAT,LNG&daddr=LAT,LNG&output=embed&dirflg=d`
2. Google Maps directions button: `maps.google.com/maps/dir/?api=1&origin=...&destination=...`
3. Waze button: `waze.com/ul?ll=LAT,LNG&navigate=yes`
4. Apple Maps button: `maps.apple.com/?saddr=...&daddr=...&dirflg=d`

Routes:
- Beach 03 town: origin 23.6855,-109.7010 → dest 23.6823,-109.6870
- Beach 01 El Cardonal: origin 23.6823,-109.6967 → dest 23.8439,-109.7441
- Beach 02 Punta Pescadero: origin 23.6823,-109.6967 → dest 23.7957,-109.7005 (via 23.8439,-109.7441)
- South run: origin 23.6823,-109.6870 → dest 23.6300,-109.6700
- Beach 04 La Ribera: origin 23.6823,-109.6967 → dest 23.5997,-109.5751
- Beach 05 Punta Colorada: origin 23.6823,-109.6967 → dest 23.5939,-109.5623

---

## June 2026 Calendar
- Grid starts June 7 (Sunday = column 7), 6 spacer cells before it
- 24 real day cells (June 7–30) + 5 trailing spacers = 35 total
- Trip dates June 15–22 marked with gold border + TRIP badge
- Moon phases: Last Qtr Jun 8, New Moon Jun 15, First Qtr Jun 21, Full Moon Jun 29

## Panga Pricing (2025–2026)
- Basic panga: ~$250–320/day + bait + tip + 16% tax
- Super panga (23–26ft): ~$480–550/day
- Deluxe cruiser: ~$900–1000/day
- Independent captain direct: ~$250–350 all-in
- Hotel Palmas de Cortez dock: walk-in booking, largest fleet
- East Cape Tackle Shop: best insider booking method
- Ramon Almanza direct: +52 624 166 7216 / ramon_losamigos@hotmail.com

## North Corridor Spots
- Ensenada de Muertos: 45 min north, grande roosters June–July
- Punta Arena de la Ventana: 55 min, excellent shore fishing
- La Ventana / El Sargento: 75 min, active panga fleet
- Cerralvo Island: offshore, world record 114lb roosterfish
- Las Arenas / La Paz: 90 min, Tailhunter International

---

## Build Method
Use Python section-by-section append — NOT heredoc (causes truncation):
```python
with open('baja_roosterfish_guide.html', 'a') as f:
    f.write(section_content)
```
Verify file size after each major section.

## Rebuild Checklist
1. Load this BUILD_SPEC.md
2. Load baja_fishing_data.json for beach/species data
3. Install leaflet@1.9.4 via npm, inline JS+CSS with base64 images
4. Build HTML section by section per document structure above
5. Apply JS architecture exactly as specified (global vars, window.load wrapper)
6. Verify: all 12 beach cards, 5 diagrams, 8 lures, 9 restaurants, 6 route nav blocks
7. Test layer toggles, flyToBeach, scrollToDiag in browser console
