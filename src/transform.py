"""Derive stints, clean laps, degradation and pit stops from the raw laps table.

Rules (documented so the numbers can be checked):
- A stint runs from the out-lap after a pit stop to the next in-lap (or race end).
  The in-lap (car crosses the line in the pit lane, 'B' in the CSV) is the LAST lap of a stint.
- A lap is "clean" if: green flag (GF), not an in-lap, not an out-lap, not the
  race's first lap, and lap time <= CLEAN_LAP_FACTOR x median lap time of its stint.
- Degradation = slope (s/lap) of a straight line through clean lap times vs lap
  number in the stint. Positive = getting slower. Note: fuel burn makes the car
  faster during a stint, so this is "net" degradation (tyres minus fuel effect).
- Pit stop time = PIT_TIME reported on the out-lap (time in the pit lane).
"""
from __future__ import annotations

import sqlite3

import numpy as np
import pandas as pd

import config


def assign_stints(laps: pd.DataFrame) -> pd.DataFrame:
    laps = laps.sort_values(["race_id", "car_number", "lap_number"]).copy()
    # A new stint starts on the lap after an in-lap.
    prev_in = laps.groupby(["race_id", "car_number"])["is_in_lap"].shift(1, fill_value=0)
    laps["stint_no"] = prev_in.groupby([laps["race_id"], laps["car_number"]]).cumsum() + 1
    laps["lap_in_stint"] = laps.groupby(["race_id", "car_number", "stint_no"]).cumcount() + 1

    candidate = (
        (laps["flag"] == "GF")
        & (laps["is_in_lap"] == 0)
        & (laps["is_out_lap"] == 0)
        & (laps["lap_number"] > 1)
        & laps["lap_time_s"].notna()
    )
    stint_median = (
        laps[candidate].groupby(["race_id", "car_number", "stint_no"])["lap_time_s"].median()
        .rename("stint_median")
    )
    laps = laps.join(stint_median, on=["race_id", "car_number", "stint_no"])
    laps["is_clean"] = (
        candidate & (laps["lap_time_s"] <= laps["stint_median"] * config.CLEAN_LAP_FACTOR)
    ).astype(int)
    return laps.drop(columns="stint_median")


def deg_slope(g: pd.DataFrame) -> float | None:
    c = g[g["is_clean"] == 1]
    if len(c) < config.MIN_LAPS_FOR_DEG:
        return None
    slope, _ = np.polyfit(c["lap_in_stint"], c["lap_time_s"], 1)
    return round(float(slope), 4)


def build_stints(laps: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for (race_id, car, stint), g in laps.groupby(["race_id", "car_number", "stint_no"]):
        clean = g.loc[g["is_clean"] == 1, "lap_time_s"]
        rows.append({
            "race_id": race_id,
            "car_number": car,
            "stint_no": int(stint),
            # driver who did most laps in the stint
            "driver_name": g["driver_name"].mode().iat[0],
            "first_lap": int(g["lap_number"].min()),
            "last_lap": int(g["lap_number"].max()),
            "laps": len(g),
            "clean_laps": int(clean.size),
            "best_lap_s": g["lap_time_s"].min(),
            "avg_clean_lap_s": round(clean.mean(), 3) if clean.size else None,
            "std_clean_lap_s": round(clean.std(), 3) if clean.size > 1 else None,
            "deg_s_per_lap": deg_slope(g),
        })
    return pd.DataFrame(rows)


def build_pit_stops(laps: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for (race_id, car), g in laps.groupby(["race_id", "car_number"]):
        g = g.sort_values("lap_number").reset_index(drop=True)
        stop_no = 0
        for i in g.index[g["is_in_lap"] == 1]:
            if i + 1 >= len(g):
                continue  # in-lap at the finish (no stop recorded)
            stop_no += 1
            in_lap, out_lap = g.loc[i], g.loc[i + 1]
            rows.append({
                "race_id": race_id,
                "car_number": car,
                "stop_no": stop_no,
                "in_lap": int(in_lap["lap_number"]),
                "pit_time_s": out_lap["pit_time_s"],
                "driver_in": in_lap["driver_name"],
                "driver_out": out_lap["driver_name"],
                "driver_change": int(in_lap["driver_name"] != out_lap["driver_name"]),
                "under_flag": in_lap["flag"],
            })
    return pd.DataFrame(rows)


def main() -> None:
    con = sqlite3.connect(config.DB_PATH)
    laps = pd.read_sql("SELECT * FROM laps", con)
    laps = assign_stints(laps)
    stints = build_stints(laps)
    pits = build_pit_stops(laps)

    with con:
        con.execute("DELETE FROM stints")
        con.execute("DELETE FROM pit_stops")
        con.executemany(
            "UPDATE laps SET stint_no=?, lap_in_stint=?, is_clean=? "
            "WHERE race_id=? AND car_number=? AND lap_number=?",
            laps[["stint_no", "lap_in_stint", "is_clean", "race_id", "car_number", "lap_number"]]
            .astype(object).itertuples(index=False, name=None),
        )
        stints.to_sql("stints", con, if_exists="append", index=False)
        pits.to_sql("pit_stops", con, if_exists="append", index=False)
    con.close()
    print(f"{len(stints)} stints, {len(pits)} pit stops, "
          f"{int(laps['is_clean'].sum())}/{len(laps)} clean laps")


if __name__ == "__main__":
    main()
