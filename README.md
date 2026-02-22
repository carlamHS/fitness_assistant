# NG Fitness Assistant (Google Sheets + Apps Script)

Starter web app to:
- Log gym training details (`training`, `mass`, `reps`, `sets`, `notes`)
- Log cardio/time-based sessions (`duration`, optional `intensity`, optional `speed`)
- Log daily protein intake by food + grams
- Save each food's protein profile (`protein per 100g`) once, then reuse it automatically
- Switch quick input mode (`Gym Work` or `Food Intake`) on mobile
- Mobile-first navigation tabs (`Log`, `Trends`, `History`) with bottom app bar
- Use top `Profile` selector so multiple people can use the same app
- Set dynamic text size (`Auto`, `Standard`, `Large`, `Extra Large`, `Huge`) for readability
- Auto-apply a distinct color theme per profile
- Edit/delete existing workout and protein entries from the UI
- Show weekly trend charts (protein totals and workout load)
- Filter weekly workout trend by selected gym work from dropdown

## Files
- `Code.gs`: backend functions and Google Sheets persistence
- `Index.html`: web app UI
- `appsscript.json`: Apps Script manifest

## Google Sheets schema
The script auto-creates 3 sheets:
- `WorkoutLog`: `CreatedAt`, `Date`, `ProfileName`, `Training`, `WorkoutType`, `WeightKg`, `Reps`, `Sets`, `DurationMin`, `Intensity`, `SpeedKph`, `Notes`
- `FoodCatalog`: `CreatedAt`, `FoodName`, `ProteinPer100g`
- `ProteinLog`: `CreatedAt`, `Date`, `ProfileName`, `FoodName`, `IntakeGrams`, `ProteinPer100g`, `ProteinGrams`, `Notes`
- `Profiles`: `CreatedAt`, `ProfileName`

## Setup (manual in Google Sheets editor)
1. Create a new Google Sheet.
2. Open `Extensions > Apps Script`.
3. Replace default files with:
   - contents of `Code.gs`
   - contents of `Index.html`
   - contents of `appsscript.json` (Project Settings > Show "appsscript.json")
4. In Apps Script, run `setupSheets` once and authorize access.
5. Run `initApp` once (optional quick check).

## Deploy as web app
1. In Apps Script: `Deploy > New deployment`.
2. Type: `Web app`.
3. Execute as: `Me`.
4. Who has access: `Anyone with the link` (or your preferred restriction).
5. Deploy and open the web app URL.

## Mobile app icon (GitHub Pages hosting)
Because Apps Script web apps cannot directly serve PNG icon files, host icon assets on GitHub Pages.

1. In GitHub repo settings, enable Pages:
   - `Settings > Pages`
   - Source: `Deploy from a branch`
   - Branch: `main`, Folder: `/(root)`
2. Push icon assets and manifest in this repo:
   - `assets/icons/icon-180.png`
   - `assets/icons/icon-192.png`
   - `assets/icons/icon-512.png`
   - `assets/icons/favicon-32.png`
   - `manifest.webmanifest`
3. `Index.html` already points to:
   - `https://carlamHS.github.io/ng_fitness_assistant/assets/icons/...`
   - `https://carlamHS.github.io/ng_fitness_assistant/manifest.webmanifest`
4. After re-deploying Apps Script, remove old Home Screen shortcut (if any), then add again.

To regenerate icons later, run:
- `python3 tools/generate_icons.py`

## Usage notes
- Workout logging supports 2 modes:
  - `Strength`: reps + sets required, mass optional.
  - `Cardio`: duration required, intensity/speed optional.
- Profile selector filters all data views and saves new entries to that profile.
- `Auto` text size now applies `Huge` scaling on mobile-like devices by default.
- Input picker buttons are mobile-friendly and show one form at a time on small screens.
- Mobile bottom tabs reduce scrolling by showing one major screen at a time.
- Workout table supports `Edit` and `Delete` per row.
- Existing old workout rows are still supported (backward compatible parsing).
- Protein logging:
  - If food exists in `FoodCatalog`, only food + grams is needed.
  - If food is new, provide `protein per 100g` once; it is then saved.
- Protein table supports `Edit` and `Delete` per row.
- Weekly charts show last 7 days ending on selected review date:
  - protein total per day
  - workout load per day (strength volume + cardio effort score)
  - workout chart can be filtered to one gym work for more valid comparison
- Formula used per protein entry:
  - `ProteinGrams = IntakeGrams * ProteinPer100g / 100`

## Optional: use clasp locally
If you prefer local development:
1. Install: `npm i -g @google/clasp`
2. Login: `clasp login`
3. Create or clone your Apps Script project.
4. Copy these files into your clasp workspace.
5. Push with: `clasp push`
