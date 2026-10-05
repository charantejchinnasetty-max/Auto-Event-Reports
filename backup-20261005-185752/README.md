# Attendance Reports - Dr. Bhatia Medical Coaching Institute

Upload the Excel exported from **Class Events > View Students > Export** and get the
daily + cumulative attendance report (batch-wise), plus unique-student counts.
Download it as an image (PNG, for WhatsApp/email) or as Excel.

## Run locally
    python -m venv .venv && source .venv/bin/activate     # Windows: .venv\Scripts\activate
    pip install -r requirements.txt
    cd backend && uvicorn main:app --reload
    # open http://localhost:8000

## Protect it (recommended before hosting - it handles student data)
    export APP_PASSWORD='choose-a-strong-password'   # optional: APP_USER (default: admin)

## Docker
    docker build -t attendance-reports .
    docker run -p 8000:8000 -e APP_PASSWORD=choose-one attendance-reports

## Notes
- Files are processed in memory and never saved on the server.
- Do not commit exports to GitHub (.gitignore already blocks .xlsx/.png).
- Batch names that are the same batch written differently: map them in `BATCH_ALIASES` in `backend/report_engine.py`.
