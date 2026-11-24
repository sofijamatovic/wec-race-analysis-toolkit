# WEC Race Analysis Toolkit

Automated post-race reporting on **public FIA WEC timing data** (2026 season):
a Python/SQL data pipeline, a one-click **Excel/VBA post-race report** and a
**Power BI season KPI dashboard**.

> Status: work in progress — data pipeline done, Excel report and Power BI next.

## Architecture

```
Al Kamel results manifest (JSON)
        │  src/ingest.py      download race lap-analysis CSVs
        ▼
SQLite  data/wec.db  ── races, laps
        │  src/transform.py   stints, clean laps, degradation, pit stops
        ▼
        ├── excel/WEC_Race_Report.xlsm   (VBA, one-click post-race report)  [next]
        └── powerbi/WEC_Season.pbix      (season KPIs)                     [next]
```

## Data source

Official timing documents published by Al Kamel Systems for the FIA WEC
(`fiawec.alkamelsystems.com`). For each race the pipeline uses the final-hour
`23_Analysis_Race_Hour N.CSV`: one row per lap with driver, lap and sector times,
pit-lane flag, pit time and track flag.

Races loaded by default: Imola, Spa-Francorchamps, **Le Mans**, Fuji
(edit `RACES` in `src/config.py`, or run with `--all`).

## Quick start (Windows)

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
cd src
python ingest.py       # downloads CSVs to data/raw and builds data/wec.db
python transform.py    # computes stints and pit stops
```

Offline test with the included sample:

```bash
python ingest.py --file ../data/sample/lemans_toyota_sample.csv --round 3 --event "LE MANS"
python transform.py
```

## Definitions

| Metric | Definition |
|---|---|
| Stint | Laps from the out-lap after a stop up to and including the next in-lap |
| Clean lap | Green flag, not in/out lap, not lap 1, ≤ 107% of the stint median |
| Consistency | Standard deviation of clean lap times in a stint (s) |
| Degradation | Slope of clean lap time vs lap-in-stint (s/lap). Net of fuel burn, so tyre wear is partly masked |
| Pit stop time | Time in the pit lane, reported on the out-lap |

## Database

See [`sql/schema.sql`](sql/schema.sql): `races`, `laps`, `stints`, `pit_stops`.

## Roadmap

- [x] Data pipeline (manifest → CSV → SQLite)
- [x] Stints, clean laps, degradation, pit stops
- [ ] Excel/VBA post-race report (one click → filled sheets, charts, PDF export)
- [ ] Power BI season dashboard
- [ ] User guide for the Excel tool
- [ ] Optional: LLM-generated race summary

Related earlier work: [f1-strategy](https://github.com/sofijamatovic/f1-strategy) (FastF1).
