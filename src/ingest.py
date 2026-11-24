"""Download official WEC lap-analysis CSVs and load them into SQLite.

Usage:
    python src/ingest.py              # races listed in config.RACES
    python src/ingest.py --all        # every race of the season
    python src/ingest.py --file path/to/23_Analysis_Race_Hour 24.CSV --round 3 --event "LE MANS"
"""
from __future__ import annotations

import argparse
import re
import sqlite3
from datetime import datetime
from io import StringIO
from pathlib import Path
from urllib.parse import quote

import pandas as pd
import requests

import config


# --------------------------------------------------------------------------- #
# 1. Find the files in the Al Kamel manifest
# --------------------------------------------------------------------------- #
def _walk(node):
    """Yield every leaf (file) in the manifest tree."""
    if isinstance(node, dict):
        if node.get("url"):
            yield node
        for child in node.get("children") or []:
            yield from _walk(child)
    elif isinstance(node, list):
        for child in node:
            yield from _walk(child)


def find_race_analysis_files(manifest: dict) -> list[dict]:
    """Return the final-hour race analysis CSV for each event.

    Path looks like:
    wec/results/15_2026/03_LE MANS/657_FIA WEC/202606131600_Race/24_Hour 24/23_Analysis_Race_Hour 24.CSV
    The last hour's file contains every lap of the race.
    """
    pattern = re.compile(
        r"wec/results/\d+_\d{4}/(?P<round>\d+)_(?P<event>[^/]+)/\d+_FIA WEC/"
        r"(?P<start>\d{12})_Race/\d+_Hour (?P<hour>\d+)/23_Analysis_Race_Hour \d+\.CSV$",
        re.IGNORECASE,
    )
    best: dict[str, dict] = {}
    for leaf in _walk(manifest.get("folders", manifest)):
        m = pattern.search(leaf["url"])
        if not m:
            continue
        info = {
            "round": int(m["round"]),
            "event": m["event"],
            "start_time": datetime.strptime(m["start"], "%Y%m%d%H%M").strftime("%Y-%m-%d %H:%M"),
            "hour": int(m["hour"]),
            "url": config.FILES_BASE + quote(leaf["url"]),
        }
        if info["event"] not in best or info["hour"] > best[info["event"]]["hour"]:
            best[info["event"]] = info
    return sorted(best.values(), key=lambda x: x["round"])


def download(url: str, dest: Path) -> Path:
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.exists():
        print(f"  cached  {dest.name}")
        return dest
    r = requests.get(url, timeout=60)
    r.raise_for_status()
    dest.write_bytes(r.content)
    print(f"  saved   {dest.name} ({len(r.content) / 1e6:.1f} MB)")
    return dest


# --------------------------------------------------------------------------- #
# 2. Parse and clean one CSV
# --------------------------------------------------------------------------- #
def to_seconds(value) -> float | None:
    """'3:29.610' -> 209.61, '0:01:04.554' -> 64.554, '' -> None."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = str(value).strip()
    if not s:
        return None
    total = 0.0
    for part in s.split(":"):
        total = total * 60 + float(part)
    return total


def parse_analysis_csv(path: Path) -> pd.DataFrame:
    text = path.read_bytes().decode("utf-8-sig", errors="replace")
    df = pd.read_csv(StringIO(text), sep=";", dtype=str, keep_default_na=False)
    df.columns = [c.strip() for c in df.columns]
    df = df.loc[:, [c for c in df.columns if c and not c.startswith("Unnamed")]]
    df = df.apply(lambda col: col.str.strip())

    out = pd.DataFrame({
        "car_number": df["NUMBER"],
        "lap_number": pd.to_numeric(df["LAP_NUMBER"], errors="coerce").astype("Int64"),
        "driver_number": pd.to_numeric(df["DRIVER_NUMBER"], errors="coerce").astype("Int64"),
        "driver_name": df["DRIVER_NAME"],
        "lap_time_s": df["LAP_TIME"].map(to_seconds),
        "s1_s": pd.to_numeric(df["S1_SECONDS"], errors="coerce"),
        "s2_s": pd.to_numeric(df["S2_SECONDS"], errors="coerce"),
        "s3_s": pd.to_numeric(df["S3_SECONDS"], errors="coerce"),
        "kph": pd.to_numeric(df["KPH"], errors="coerce"),
        "top_speed": pd.to_numeric(df["TOP_SPEED"], errors="coerce"),
        "elapsed_s": df["ELAPSED"].map(to_seconds),
        "is_in_lap": (df["CROSSING_FINISH_LINE_IN_PIT"] == "B").astype(int),
        "pit_time_s": df["PIT_TIME"].map(to_seconds),
        "flag": df["FLAG_AT_FL"],
        "class": df["CLASS"],
        "team": df["TEAM"],
        "manufacturer": df["MANUFACTURER"],
    })
    out["is_out_lap"] = out["pit_time_s"].notna().astype(int)
    out = out.dropna(subset=["lap_number"])
    return out.sort_values(["car_number", "lap_number"]).reset_index(drop=True)


# --------------------------------------------------------------------------- #
# 3. Load into SQLite
# --------------------------------------------------------------------------- #
def init_db(con: sqlite3.Connection) -> None:
    con.executescript(config.SCHEMA_PATH.read_text(encoding="utf-8"))


def load_race(con: sqlite3.Connection, info: dict, laps: pd.DataFrame) -> int:
    cur = con.execute(
        "INSERT INTO races (season, round, event, start_time, source_url, loaded_at) "
        "VALUES (?, ?, ?, ?, ?, ?)",
        (config.SEASON, info["round"], info["event"], info.get("start_time"),
         info["url"], datetime.now().isoformat(timespec="seconds")),
    )
    race_id = cur.lastrowid
    laps = laps.assign(race_id=race_id)
    laps.to_sql("laps", con, if_exists="append", index=False)
    return race_id


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--all", action="store_true", help="load every race of the season")
    ap.add_argument("--file", type=Path, help="load one local CSV instead of downloading")
    ap.add_argument("--round", type=int, default=0)
    ap.add_argument("--event", default="LOCAL FILE")
    args = ap.parse_args()

    con = sqlite3.connect(config.DB_PATH)
    init_db(con)

    if args.file:
        races = [{"round": args.round, "event": args.event, "url": str(args.file), "path": args.file}]
    else:
        print(f"Reading manifest {config.MANIFEST_URL}")
        manifest = requests.get(config.MANIFEST_URL, timeout=60).json()
        races = find_race_analysis_files(manifest)
        if not args.all and config.RACES:
            races = [r for r in races if r["event"] in config.RACES]
        for r in races:
            fname = f"{config.SEASON}_R{r['round']:02d}_{r['event'].replace(' ', '_')}_race_analysis.csv"
            r["path"] = download(r["url"], config.DATA_RAW / fname)

    for r in races:
        laps = parse_analysis_csv(r["path"])
        race_id = load_race(con, r, laps)
        print(f"  loaded  R{r['round']:02d} {r['event']:<22} {len(laps):>6} laps, "
              f"{laps['car_number'].nunique()} cars  (race_id={race_id})")

    con.commit()
    con.close()
    print(f"Database: {config.DB_PATH}")


if __name__ == "__main__":
    main()
