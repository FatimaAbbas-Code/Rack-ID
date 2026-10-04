# Picked up automatically by `gunicorn app:app` (run from the repo root).
# Importing a 70+ line factory invoice uploads dozens of photos to R2, which
# can take longer than gunicorn's default 30s worker timeout.
timeout = 180
