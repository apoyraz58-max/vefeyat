import json
import os
import secrets
from pathlib import Path
from urllib.parse import unquote, urlparse

BASE_DIR = Path(__file__).resolve().parent.parent
DEBUG = os.environ.get("DJANGO_DEBUG", "1") == "1"
SECRET_KEY = os.environ.get("DJANGO_SECRET_KEY")
if not SECRET_KEY:
    if not DEBUG:
        raise RuntimeError("Yayında DJANGO_SECRET_KEY ayarlanmalıdır.")
    local = BASE_DIR / ".local"
    local.mkdir(exist_ok=True, mode=0o700)
    key_file = local / "django-secret"
    try:
        with key_file.open("x", encoding="utf-8") as stream:
            key_file.chmod(0o600)
            stream.write(secrets.token_urlsafe(64))
    except FileExistsError:
        pass
    SECRET_KEY = key_file.read_text(encoding="utf-8")

ALLOWED_HOSTS = [host.strip() for host in os.environ.get("DJANGO_ALLOWED_HOSTS", "localhost,127.0.0.1,[::1],testserver").split(",") if host.strip()]
INSTALLED_APPS = [
    "django.contrib.admin", "django.contrib.auth", "django.contrib.contenttypes",
    "django.contrib.sessions", "django.contrib.messages", "django.contrib.staticfiles",
    "django.contrib.postgres", "catalog",
]
MIDDLEWARE = [
    "django.middleware.security.SecurityMiddleware", "django.contrib.sessions.middleware.SessionMiddleware",
    "django.middleware.common.CommonMiddleware", "django.middleware.csrf.CsrfViewMiddleware",
    "django.contrib.auth.middleware.AuthenticationMiddleware", "django.contrib.messages.middleware.MessageMiddleware",
    "vefeyat.middleware.FramePolicyMiddleware",
]
# Word task panes run inside an Office frame. Do not add XFrameOptionsMiddleware
# globally: the public task pane must be embeddable; admin gets DENY in its view.
ROOT_URLCONF = "vefeyat.urls"
TEMPLATES = [{
    "BACKEND": "django.template.backends.django.DjangoTemplates",
    "DIRS": [BASE_DIR / "templates"], "APP_DIRS": True,
    "OPTIONS": {"context_processors": [
        "django.template.context_processors.request", "django.contrib.auth.context_processors.auth",
        "django.contrib.messages.context_processors.messages",
    ]},
}]
WSGI_APPLICATION = "vefeyat.wsgi.application"

database_url = os.environ.get("DATABASE_URL")
if database_url:
    parsed = urlparse(database_url)
    if parsed.scheme not in {"postgres", "postgresql"}:
        raise RuntimeError("DATABASE_URL PostgreSQL bağlantısı olmalıdır.")
    database = {
        "NAME": unquote(parsed.path.lstrip("/")), "USER": unquote(parsed.username or ""),
        "PASSWORD": unquote(parsed.password or ""), "HOST": parsed.hostname or "localhost",
        "PORT": parsed.port or 5432,
    }
    from urllib.parse import parse_qs
    options = parse_qs(parsed.query)
    if "sslmode" in options:
        database["OPTIONS"] = {"sslmode": options["sslmode"][0]}
elif (BASE_DIR / ".local/dev-database.json").exists():
    database = json.loads((BASE_DIR / ".local/dev-database.json").read_text())
else:
    database = {"NAME": "vefeyat", "USER": "vefeyat", "HOST": "127.0.0.1", "PORT": 55432}
DATABASES = {"default": {"ENGINE": "django.db.backends.postgresql", **database}}
AUTH_PASSWORD_VALIDATORS = [
    {"NAME": "django.contrib.auth.password_validation.UserAttributeSimilarityValidator"},
    {"NAME": "django.contrib.auth.password_validation.MinimumLengthValidator"},
    {"NAME": "django.contrib.auth.password_validation.CommonPasswordValidator"},
    {"NAME": "django.contrib.auth.password_validation.NumericPasswordValidator"},
]
LANGUAGE_CODE = "tr"
TIME_ZONE = "Europe/Istanbul"
USE_I18N = True
USE_TZ = True
STATIC_URL = "/static/"
STATICFILES_DIRS = [BASE_DIR / "static"]
STATIC_ROOT = BASE_DIR / ".local/static"
DEFAULT_AUTO_FIELD = "django.db.models.BigAutoField"
DATA_UPLOAD_MAX_MEMORY_SIZE = 32_768
CSRF_TRUSTED_ORIGINS = [v.strip() for v in os.environ.get("CSRF_TRUSTED_ORIGINS", "").split(",") if v.strip()]
SESSION_COOKIE_HTTPONLY = True
SESSION_COOKIE_SAMESITE = "Lax"
SESSION_COOKIE_SECURE = not DEBUG
CSRF_COOKIE_SECURE = not DEBUG
SECURE_CONTENT_TYPE_NOSNIFF = True
SECURE_REFERRER_POLICY = "same-origin"
if not DEBUG:
    SECURE_SSL_REDIRECT = True
    SECURE_HSTS_SECONDS = 3600
# Set only behind a proxy you control that strips incoming forwarding headers.
if os.environ.get("TRUST_HTTPS_PROXY") == "1":
    SECURE_PROXY_SSL_HEADER = ("HTTP_X_FORWARDED_PROTO", "https")

WORD_BASE_URL = os.environ.get("WORD_BASE_URL", "https://localhost:8443").rstrip("/")
