# Calorimeter Comparison Dashboard

Compares Walton compressor calorimeter test reports (PDF) across compressor IDs, grouped by RPM test tier.

## Folder structure expected

```
MainFolder/
  ID-7338.01/
    ID-7338.01 (Without Vitrek) (25-06-2026) 1400 RPM.pdf
    ID-7338.01 (Without Vitrek) (25-06-2026) 2000 RPM.pdf
    ...
  ID-7338.02/
    ID-7338.02 ... 1400 RPM.pdf
    ...
```

Each subfolder = one compressor sample. Each PDF = one RPM test point. The dashboard
reads the actual Sample ID and RPM straight out of each report's content (not just
the filename), so it will still group correctly even if filenames vary.

## Setup

```bash
cd calorimeter_dashboard
python3 -m venv venv
source venv/bin/activate        # venv\Scripts\activate on Windows
pip install -r requirements.txt
python app.py
```

Then open **http://localhost:5000** in your browser.

## Using it

1. Click **"Select main folder…"** in the left sidebar and choose your top-level
   folder (the one containing `ID-xxxx.xx` subfolders). Chrome/Edge will ask you
   to confirm folder access — this is normal for directory uploads.
2. All PDFs are parsed automatically. KPI cards and RPM-tier comparison tables
   appear on the right.
3. Use the **Parameters** checklist on the left to show/hide columns in the
   comparison tables (and in exports).
4. Use **Export XLSX** / **Export PDF** in the top bar to download the current
   comparison (respecting your parameter selection) — XLSX has one sheet per
   RPM tier plus a summary sheet; PDF is a landscape multi-table report.
5. Toggle dark/light mode with the sun icon, top right.

## Notes

- Everything runs locally — no data leaves your machine.
- The compressor's actual RPM (as measured) is shown as a data column;
  reports are *bucketed* into RPM tiers (1400, 2000, 2400, 3000, 3600, 4000, …)
  based on the *rated/target* RPM in each report so tests at nominally the same
  point are compared together even if the achieved RPM differs slightly.
- Want this packaged as a desktop app (Electron) or deployed to Render.com,
  like the earlier motor-test dashboards? Just ask.
