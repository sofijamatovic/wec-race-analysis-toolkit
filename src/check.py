import sqlite3

con = sqlite3.connect("../data/wec.db")

stints = con.execute("""
    SELECT COUNT(*), SUM(laps)
    FROM stints JOIN races USING (race_id)
    WHERE event = 'LE MANS' AND car_number = '7'
""").fetchone()

pits = con.execute("""
    SELECT COUNT(*), ROUND(AVG(pit_time_s), 1)
    FROM pit_stops JOIN races USING (race_id)
    WHERE event = 'LE MANS' AND car_number = '7'
""").fetchone()

print(f"Toyota #7 Le Mans: {stints[0]} stintova, {stints[1]} krugova")
print(f"Toyota #7 Le Mans: {pits[0]} pit stopova, prosek {pits[1]} s")