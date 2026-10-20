#!/bin/sh
set -eu

# The workbook is mounted read-only; application data never enters the image.
if [ ! -r /data/people.xlsx ]; then
    printf 'Kaynak Excel /data/people.xlsx konumunda okunabilir olmalıdır.\n' >&2
    exit 1
fi

python - <<'PY'
import os
import sys
import time

import psycopg

for attempt in range(60):
    try:
        with psycopg.connect(os.environ["DATABASE_URL"], connect_timeout=2) as connection:
            connection.execute("SELECT 1")
        break
    except psycopg.Error:
        time.sleep(1)
else:
    print("PostgreSQL hazır değil; db servisinin durumunu inceleyin.", file=sys.stderr)
    sys.exit(1)
PY

python manage.py migrate --noinput
python manage.py collectstatic --noinput
python manage.py import_people /data/people.xlsx
exec gunicorn vefeyat.wsgi:application \
    --bind 0.0.0.0:8000 \
    --workers 2 \
    --threads 2 \
    --timeout 60 \
    --access-logfile - \
    --error-logfile -
