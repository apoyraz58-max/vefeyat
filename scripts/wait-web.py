"""Verify that Django responds successfully after startup."""

import sys
import time
import urllib.error
import urllib.request

for attempt in range(60):
    try:
        with urllib.request.urlopen("http://127.0.0.1:8000/api/health/", timeout=2) as response:
            if response.status == 200:
                print("Uygulama sağlık kontrolü başarılı.")
                sys.exit(0)
    except (OSError, urllib.error.URLError):
        time.sleep(0.5)
print("Uygulama hazır değil; .local/web.log dosyasını inceleyin.", file=sys.stderr)
sys.exit(1)
