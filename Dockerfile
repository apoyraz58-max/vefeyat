FROM python:3.12-slim-bookworm

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

WORKDIR /app
COPY requirements.txt ./requirements.txt
RUN python -m pip install --no-cache-dir --require-hashes -r requirements.txt \
    && groupadd --gid 10001 appuser \
    && useradd --uid 10001 --gid 10001 --no-create-home --home-dir /app appuser \
    && mkdir -p /app/.local/static \
    && chown -R appuser:appuser /app

COPY --chown=appuser:appuser . /app/
RUN chmod 755 /app/scripts/container-start.sh
USER 10001:10001
EXPOSE 8000
CMD ["/bin/sh", "/app/scripts/container-start.sh"]
