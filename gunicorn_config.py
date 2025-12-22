# gunicorn_config.py
import os

bind = f"0.0.0.0:{os.getenv('PORT', '8080')}"
workers = int(os.getenv("WEB_CONCURRENCY", "1"))  # keep 1 on 512MB
timeout = int(os.getenv("GUNICORN_TIMEOUT", "120"))
