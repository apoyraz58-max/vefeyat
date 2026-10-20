"""Wait for the project-owned local PostgreSQL database without logging secrets."""

import json
import pathlib
import sys
import time

import psycopg

root = pathlib.Path(__file__).resolve().parents[1]
for attempt in range(120):
    try:
        config = json.loads((root / ".local/dev-database.json").read_text())
        with psycopg.connect(
            dbname=config["NAME"],
            user=config["USER"],
            password=config["PASSWORD"],
            host=config["HOST"],
            port=config["PORT"],
            connect_timeout=2,
        ) as connection:
            assert connection.execute("SELECT 1").fetchone() == (1,)
        print("PostgreSQL bağlantısı doğrulandı.")
        sys.exit(0)
    except (OSError, ValueError, KeyError, psycopg.Error):
        time.sleep(0.5)
print("PostgreSQL hazır değil; .local/db.log dosyasını inceleyin.", file=sys.stderr)
sys.exit(1)
