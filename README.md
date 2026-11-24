# WEC Race Analysis Toolkit

Automated post-race reporting on **public FIA WEC timing data** (2026 season):
a Python/SQL data pipeline, a one-click **Excel/VBA post-race report** and a
**Power BI season KPI dashboard**.

> Status: data pipeline and Excel report done; Power BI next.

## Architecture

```
Al Kamel results manifest (JSON)
        │  src/ingest.py      download race lap-analysis CSVs
        ▼
SQLite  data/wec.db  ── races, laps
        │  src/transform.py   stints, clean laps, degradation, pit stops
        │  sql/views.sql      KPI definitions (car, driver, stint, pit-stop level)
        │  src/export.py      views → data/export/*.csv
        ▼
        ├── excel/WEC_Race_Report.xlsm   one-click post-race report (VBA)
        └── powerbi/                     season KPIs                 [next]
```

## Data source

Official timing documents published by Al Kamel Systems for the FIA WEC
(`fiawec.alkamelsystems.com`). For each race the pipeline uses the final-hour
`23_Analysis_Race_Hour N.CSV`: one row per lap with driver, lap and sector times,
pit-lane flag, pit time and track flag.

Races loaded by default: Imola, Spa-Francorchamps, **Le Mans**, Fuji
(edit `RACES` in `src/config.py`, or run `src/ingest.py --all`).

## Quick start (Windows)

```bash
py -V:3.14 -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python src/run_all.py      # download → SQLite → stints/pit stops → CSV export
```

## Excel post-race report — user guide

The report works for **any car in any loaded race**, not only Toyota.

**First-time setup** (the VBA source is kept in `excel/WecReport.bas` so it is version-controlled):

1. Open Excel → blank workbook → *Save As* `excel/WEC_Race_Report.xlsm` (*Excel Macro-Enabled Workbook*).
2. `Alt+F11` → *File → Import File…* → `excel/WecReport.bas`. Close the VBA editor.
3. `Alt+F8` → `SetupWorkbook` → *Run*. This builds the Control sheet, buttons and report sheets.

**Every report:**

1. **Load data** – imports `data/export/*.csv` (re-run after `run_all.py`).
2. Pick a **race** and a **car** in the yellow drop-downs (the car list follows the race).
3. **Generate report** – fills five sheets:

| Sheet | Contents |
|---|---|
| Summary | Result, pace and pit-lane KPI cards; driver table |
| Stints | Stint table with colour scales; avg clean lap and degradation per stint |
| Pit Stops | Every stop, driver change, flag; pit-lane time chart |
| Lap Times | Clean laps by driver; scatter chart |
| Class Comparison | All cars in the class; gap-to-best chart, selected car and team-mates highlighted |

4. **Export PDF** – saves all report sheets to `excel/reports/<race>_<car>.pdf`.

If you change `excel/WecReport.bas`, re-import it (remove the old module first).

## Definitions

| Metric | Definition |
|---|---|
| Stint | Laps from the out-lap after a stop up to and including the next in-lap |
| Clean lap | Green flag, not in/out lap, not lap 1, ≤ 107% of the stint median |
| Consistency | Standard deviation of clean lap times in a stint (s), averaged over stints |
| Degradation | Slope of clean lap time vs lap-in-stint (s/lap). Net of fuel burn, so tyre wear is partly masked |
| Pace gap | Car's average clean lap minus the best average in its class (cars with ≥ 20 clean laps) |
| Pit stop time | Time in the pit lane, reported on the out-lap |
| Position | Approximated from laps completed, then race time; official classification can differ |

All KPIs are defined in SQL in [`sql/views.sql`](sql/views.sql).

## Validation

- Toyota #7, Le Mans 2026: 381 laps, 33 stints, 32 pit stops — matches official timing.
- Across all loaded races: stints − pit stops = number of cars (every car has one more stint than stops).

## Roadmap

- [x] Data pipeline (manifest → CSV → SQLite)
- [x] Stints, clean laps, degradation, pit stops
- [x] KPI views and CSV export
- [x] Excel/VBA post-race report (one click → filled sheets, charts, PDF export)
- [ ] Power BI season dashboard
- [ ] Optional: LLM-generated race summary

Related earlier work: [f1-strategy](https://github.com/sofijamatovic/f1-strategy) (FastF1).
