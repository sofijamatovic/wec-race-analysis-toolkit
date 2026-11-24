"""Project-wide settings."""
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA_RAW = ROOT / "data" / "raw"
DB_PATH = ROOT / "data" / "wec.db"
SCHEMA_PATH = ROOT / "sql" / "schema.sql"
VIEWS_PATH = ROOT / "sql" / "views.sql"
DATA_EXPORT = ROOT / "data" / "export"

# Al Kamel Systems publishes all official FIA WEC timing documents here.
# The results website (fiawec.alkamelsystems.com) reads this JSON manifest.
FILES_BASE = "https://files.alkamelcloud.com/"
SEASON = 2026
MANIFEST_URL = f"{FILES_BASE}wec/manifests/results_15_{SEASON}.json"

# Races to load. Names must match the event folder names in the manifest
# (e.g. "03_LE MANS" -> "LE MANS"). Set to None to load the whole season.
RACES = ["IMOLA", "SPA FRANCORCHAMPS", "LE MANS", "FUJI SPEEDWAY"]

# Clean-lap rule: a lap counts as "clean" if it is green-flag, not an in/out lap,
# and no slower than this factor x the median lap of its stint.
CLEAN_LAP_FACTOR = 1.07
# Minimum clean laps in a stint before we fit a degradation slope.
MIN_LAPS_FOR_DEG = 5
