"""Run the whole data pipeline in one go.

    python src/run_all.py

1. ingest.py    download race timing CSVs and load laps into SQLite
2. transform.py compute stints, clean laps, degradation, pit stops
3. export.py    build report views and write CSV files for Excel / Power BI
"""
import export
import ingest
import transform

if __name__ == "__main__":
    print("== 1/3 ingest ==")
    ingest.main()
    print("== 2/3 transform ==")
    transform.main()
    print("== 3/3 export ==")
    export.main()
