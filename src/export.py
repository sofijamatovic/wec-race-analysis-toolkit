"""Create the report views and export them as CSV files for Excel and Power BI.

Usage:
    python src/export.py

Output: data/export/{races,car_summary,stints,pit_stops,driver_summary,laps}.csv
"""
from __future__ import annotations

import sqlite3

import pandas as pd

import config

# file name -> SQL query. The ORDER BY matters: the Excel car drop-down
# expects all cars of one race to be in consecutive rows of car_summary.
EXPORTS = {
    "races":          "SELECT * FROM v_races ORDER BY round",
    "car_summary":    "SELECT * FROM v_car_summary ORDER BY round, pos_overall, car_label",
    "stints":         "SELECT * FROM v_stints ORDER BY round, car_label, stint_no",
    "pit_stops":      "SELECT * FROM v_pit_stops ORDER BY round, car_label, stop_no",
    "driver_summary": "SELECT * FROM v_driver_summary ORDER BY round, car_label, first_lap",
    "laps":           "SELECT * FROM v_laps ORDER BY round, car_label, lap_number",
}


def main() -> None:
    config.DATA_EXPORT.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(config.DB_PATH)
    con.executescript(config.VIEWS_PATH.read_text(encoding="utf-8"))

    for name, query in EXPORTS.items():
        df = pd.read_sql(query, con)
        path = config.DATA_EXPORT / f"{name}.csv"
        df.to_csv(path, index=False, encoding="utf-8", float_format="%.4f")
        print(f"  exported {name + '.csv':<20} {len(df):>6} rows")

    con.close()
    print(f"CSV files: {config.DATA_EXPORT}")


if __name__ == "__main__":
    main()
