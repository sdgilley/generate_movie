Minimal FastAPI + Celery scaffold (MVP)

Run locally (development):

1. Start Redis (e.g. `docker run -p 6379:6379 redis:7`) or use docker-compose
2. Install Python deps: `pip install -r webapp/requirements.txt`
3. Start worker: `celery -A webapp.celery_app.celery worker --loglevel=info`
4. Start web server: `uvicorn webapp.app:app --reload`

API endpoints:
- POST /upload (multipart form) fields: `ppt` (file), `speech_key`, `endpoint`, `voice_name`
  - returns: {"job_id": "..."}
- GET /status/{job_id}
- GET /result/{job_id} -> returns mp4 file when ready

Notes:
- This MVP keeps per-job files in temporary directories; results are returned as local FileResponse.
- For production, add authentication, storage (S3/Blob), and secret encryption for user-provided keys.
