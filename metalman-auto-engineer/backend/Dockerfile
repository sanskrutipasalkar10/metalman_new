FROM python:3.10-slim

# System libraries required by opencv-python-headless (libgl1/libglib2.0-0)
# and cadquery/OCP (libsm6/libxext6/libxrender1) on Debian, plus a real font
# so PIL's ImageFont has something better than the bitmap default on Linux
# (arial.ttf only exists on Windows; the code already falls back gracefully,
# this just makes the fallback look decent).
RUN apt-get update && apt-get install -y --no-install-recommends \
    libgl1 \
    libglib2.0-0 \
    libsm6 \
    libxext6 \
    libxrender1 \
    libfontconfig1 \
    fonts-dejavu-core \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Persistent data (uploads/outputs/sqlite db) - main.py and api/routes.py
# already auto-detect this path and switch to it when present.
RUN mkdir -p /data/uploads /data/outputs
VOLUME ["/data"]

EXPOSE 8000

# IMPORTANT: keep --workers 1. Task status (TASKS dict in api/routes.py)
# lives in per-process memory and is only persisted to tasks_db.json on
# write, not re-read live - multiple worker processes would each have their
# own inconsistent view of task state and break /api/status polling.
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "1"]
