-- WEC Race Analysis Toolkit - report views
-- Every KPI used in the Excel report and Power BI is defined here, in one place.
-- Run by src/export.py, which writes each view to data/export/<name>.csv.

DROP VIEW IF EXISTS v_races;
DROP VIEW IF EXISTS v_cars;
DROP VIEW IF EXISTS v_car_summary;
DROP VIEW IF EXISTS v_stints;
DROP VIEW IF EXISTS v_pit_stops;
DROP VIEW IF EXISTS v_driver_summary;
DROP VIEW IF EXISTS v_laps;

CREATE VIEW v_races AS
SELECT race_id, round, event, start_time
FROM races;

-- One row per car per race. car_label ('#7 Toyota Racing') is the key the Excel tool uses.
CREATE VIEW v_cars AS
SELECT l.race_id,
       r.round,
       r.event,
       l.car_number,
       '#' || l.car_number                       AS car,
       '#' || l.car_number || ' ' || MAX(l.team) AS car_label,
       MAX(l.team)                               AS team,
       MAX(l.class)                              AS class,
       MAX(l.manufacturer)                       AS manufacturer
FROM laps l
JOIN races r ON r.race_id = l.race_id
GROUP BY l.race_id, l.car_number;

-- Car-level KPIs.
-- Positions are approximated from laps completed, then race time at the last lap.
-- Official classification can differ (penalties, non-classified cars).
CREATE VIEW v_car_summary AS
WITH lap_stats AS (
    SELECT race_id, car_number,
           COUNT(*)                                          AS laps,
           MAX(elapsed_s)                                    AS race_time_s,
           MIN(lap_time_s)                                   AS best_lap_s,
           AVG(CASE WHEN is_clean = 1 THEN lap_time_s END)   AS avg_clean_lap_s,
           SUM(is_clean)                                     AS clean_laps
    FROM laps
    GROUP BY race_id, car_number
),
stint_stats AS (
    SELECT race_id, car_number,
           COUNT(*)             AS stints,
           AVG(std_clean_lap_s) AS consistency_s,
           AVG(deg_s_per_lap)   AS avg_deg_s_per_lap
    FROM stints
    GROUP BY race_id, car_number
),
pit_stats AS (
    SELECT race_id, car_number,
           COUNT(*)                                               AS pit_stops,
           SUM(pit_time_s)                                        AS total_pit_s,
           AVG(pit_time_s)                                        AS avg_pit_s,
           AVG(CASE WHEN driver_change = 1 THEN pit_time_s END)   AS avg_pit_driver_change_s,
           AVG(CASE WHEN driver_change = 0 THEN pit_time_s END)   AS avg_pit_no_change_s,
           SUM(CASE WHEN under_flag IN ('FCY', 'SF') THEN 1 ELSE 0 END) AS stops_under_caution
    FROM pit_stops
    GROUP BY race_id, car_number
)
SELECT c.round, c.event, c.car, c.car_label, c.team, c.class, c.manufacturer,
       RANK() OVER (PARTITION BY c.race_id ORDER BY ls.laps DESC, ls.race_time_s)          AS pos_overall,
       RANK() OVER (PARTITION BY c.race_id, c.class ORDER BY ls.laps DESC, ls.race_time_s) AS pos_class,
       ls.laps,
       ROUND(ls.race_time_s, 3)     AS race_time_s,
       ls.best_lap_s,
       ROUND(ls.avg_clean_lap_s, 3) AS avg_clean_lap_s,
       ls.clean_laps,
       -- gap to the best average clean lap in the class (cars with >= 20 clean laps set the reference)
       ROUND(ls.avg_clean_lap_s
             - MIN(CASE WHEN ls.clean_laps >= 20 THEN ls.avg_clean_lap_s END)
               OVER (PARTITION BY c.race_id, c.class), 3) AS pace_gap_s,
       COALESCE(ss.stints, 0)               AS stints,
       ROUND(ss.consistency_s, 3)           AS consistency_s,
       ROUND(ss.avg_deg_s_per_lap, 4)       AS avg_deg_s_per_lap,
       COALESCE(ps.pit_stops, 0)            AS pit_stops,
       ROUND(ps.total_pit_s, 1)             AS total_pit_s,
       ROUND(ps.avg_pit_s, 1)               AS avg_pit_s,
       ROUND(ps.avg_pit_driver_change_s, 1) AS avg_pit_driver_change_s,
       ROUND(ps.avg_pit_no_change_s, 1)     AS avg_pit_no_change_s,
       COALESCE(ps.stops_under_caution, 0)  AS stops_under_caution
FROM v_cars c
JOIN lap_stats ls        ON ls.race_id = c.race_id AND ls.car_number = c.car_number
LEFT JOIN stint_stats ss ON ss.race_id = c.race_id AND ss.car_number = c.car_number
LEFT JOIN pit_stats ps   ON ps.race_id = c.race_id AND ps.car_number = c.car_number;

CREATE VIEW v_stints AS
SELECT c.round, c.event, c.car_label,
       s.stint_no, s.driver_name, s.first_lap, s.last_lap, s.laps, s.clean_laps,
       s.best_lap_s, s.avg_clean_lap_s, s.std_clean_lap_s, s.deg_s_per_lap
FROM stints s
JOIN v_cars c ON c.race_id = s.race_id AND c.car_number = s.car_number;

CREATE VIEW v_pit_stops AS
SELECT c.round, c.event, c.car_label,
       p.stop_no, p.in_lap, p.driver_in, p.driver_out,
       CASE WHEN p.driver_change = 1 THEN 'Yes' ELSE 'No' END AS driver_change,
       p.under_flag,
       p.pit_time_s
FROM pit_stops p
JOIN v_cars c ON c.race_id = p.race_id AND c.car_number = p.car_number;

CREATE VIEW v_driver_summary AS
WITH dl AS (
    SELECT race_id, car_number, driver_name,
           MIN(lap_number)                                 AS first_lap,
           COUNT(*)                                        AS laps,
           SUM(is_clean)                                   AS clean_laps,
           MIN(lap_time_s)                                 AS best_lap_s,
           AVG(CASE WHEN is_clean = 1 THEN lap_time_s END) AS avg_clean_lap_s
    FROM laps
    GROUP BY race_id, car_number, driver_name
),
ds AS (
    SELECT race_id, car_number, driver_name,
           COUNT(*)             AS stints,
           AVG(std_clean_lap_s) AS consistency_s,
           AVG(deg_s_per_lap)   AS avg_deg_s_per_lap
    FROM stints
    GROUP BY race_id, car_number, driver_name
)
SELECT c.round, c.event, c.car_label,
       dl.driver_name, dl.first_lap, dl.laps, dl.clean_laps,
       COALESCE(ds.stints, 0)         AS stints,
       dl.best_lap_s,
       ROUND(dl.avg_clean_lap_s, 3)   AS avg_clean_lap_s,
       ROUND(ds.consistency_s, 3)     AS consistency_s,
       ROUND(ds.avg_deg_s_per_lap, 4) AS avg_deg_s_per_lap
FROM dl
JOIN v_cars c  ON c.race_id = dl.race_id AND c.car_number = dl.car_number
LEFT JOIN ds   ON ds.race_id = dl.race_id AND ds.car_number = dl.car_number
              AND ds.driver_name = dl.driver_name;

CREATE VIEW v_laps AS
SELECT c.round, c.event, c.car_label,
       l.lap_number, l.driver_name, l.stint_no, l.lap_time_s, l.is_clean, l.flag
FROM laps l
JOIN v_cars c ON c.race_id = l.race_id AND c.car_number = l.car_number;
