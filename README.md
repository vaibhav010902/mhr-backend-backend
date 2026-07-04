# MHR Platform — Backend

This folder contains the FastAPI backend for the MHR Analytics Platform.

## Quick local run

1. Create a virtual environment and install dependencies:

```bash
python -m venv .venv
source .venv/bin/activate  # or .venv\Scripts\activate on Windows
pip install -r requirements.txt
```

2. Create an `.env` file from the example and update values:

```bash
cp .env.example .env
# then edit .env and fill in MONGO_URI etc.
```

3. Run the app:

```bash
uvicorn app.main:app --host 0.0.0.0 --port 8000
```

4. Test the health endpoint:

```bash
curl -sS http://127.0.0.1:8000/api/health
# expect: {"status":"ok"}
```

## Render (Docker) deployment

1. In Render, create a new Web Service → Docker:
   - Repository: select your backend repo
   - Root Directory: `backend` (or leave as repo root if Dockerfile is in root)
   - Dockerfile Path: `Dockerfile`
   - Branch: `main`

2. Add Environment variables in Render Service settings:

```
APP_NAME=MHR Analytics Platform
MONGO_URI=your_mongodb_connection_string
DATABASE_NAME=mhr_platform
ALLOWED_ORIGINS=https://your-frontend-domain
```

3. Deploy and watch logs. Health endpoint on Render should be:

```
https://<your-render-url>/api/health
```

## Frontend

Set `VITE_API_BASE` in your frontend hosting provider to point to the Render URL:

```
VITE_API_BASE=https://<your-render-url>
```

## Troubleshooting

- If the build fails on pip metadata for packages that require Rust, prefer Docker deploy (this repo includes a Dockerfile).
- Check that the service root points to the folder containing `Dockerfile` or that Dockerfile path is correct.
- Confirm `MONGO_URI` is valid and IP whitelist in Atlas allows Render's outbound IPs (or set 0.0.0.0/0 for testing).

## Files added by automation
- `.env.example` contains the environment keys used by `app/core/config.py`.
