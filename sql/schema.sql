-- WEC Race Analysis Toolkit - SQLite schema
-- Raw lap data comes from Al Kamel "23_Analysis_Race_Hour N.CSV" files.

DROP TABLE IF EXISTS pit_stops;
DROP TABLE IF EXISTS stints;
DROP TABLE IF EXISTS laps;
DROP TABLE IF EXISTS races;

CREATE TABLE races (
    race_id      INTEGER PRIMARY KEY,
    season       INTEGER NOT NULL,
    round        INTEGER NOT NULL,
    event        TEXT    NOT NULL,      -- e.g. 'LE MANS'
    start_time   TEXT,                  -- from folder name, e.g. '2026-06-13 16:00'
    source_url   TEXT    NOT NULL,
    loaded_at    TEXT    NOT NULL,
    UNIQUE (season, round)
);

CREATE TABLE laps (
    race_id        INTEGER NOT NULL REFERENCES races(race_id),
    car_number     TEXT    NOT NULL,    -- text on purpose: '007' and '7' are different cars
    lap_number     INTEGER NOT NULL,
    driver_number  INTEGER,
    driver_name    TEXT,
    lap_time_s     REAL,
    s1_s           REAL,
    s2_s           REAL,
    s3_s           REAL,
    kph            REAL,
    top_speed      REAL,
    elapsed_s      REAL,                -- race time at the end of the lap
    is_in_lap      INTEGER NOT NULL,    -- 1 = car crossed the line in the pit lane
    is_out_lap     INTEGER NOT NULL,    -- 1 = lap that contains the pit stop time
    pit_time_s     REAL,                -- time spent in the pit lane (on the out-lap)
    flag           TEXT,                -- GF green, FCY full course yellow, SF safety car, FF finish
    class          TEXT,                -- HYPERCAR / LMGT3
    team           TEXT,
    manufacturer   TEXT,
    stint_no       INTEGER,             -- filled by transform step
    lap_in_stint   INTEGER,
    is_clean       INTEGER,             -- filled by transform step
    PRIMARY KEY (race_id, car_number, lap_number)
);

CREATE TABLE stints (
    race_id          INTEGER NOT NULL REFERENCES races(race_id),
    car_number       TEXT    NOT NULL,
    stint_no         INTEGER NOT NULL,
    driver_name      TEXT,
    first_lap        INTEGER,
    last_lap         INTEGER,
    laps             INTEGER,
    clean_laps       INTEGER,
    best_lap_s       REAL,
    avg_clean_lap_s  REAL,
    std_clean_lap_s  REAL,               -- consistency
    deg_s_per_lap    REAL,               -- slope of clean lap time vs lap in stint
    PRIMARY KEY (race_id, car_number, stint_no)
);

CREATE TABLE pit_stops (
    race_id        INTEGER NOT NULL REFERENCES races(race_id),
    car_number     TEXT    NOT NULL,
    stop_no        INTEGER NOT NULL,
    in_lap         INTEGER,
    pit_time_s     REAL,
    driver_in      TEXT,
    driver_out     TEXT,
    driver_change  INTEGER,
    under_flag     TEXT,                 -- flag on the in-lap (stop under FCY/SC is 'cheaper')
    PRIMARY KEY (race_id, car_number, stop_no)
);

CREATE INDEX idx_laps_car ON laps (race_id, car_number);
CREATE INDEX idx_laps_team ON laps (team);
