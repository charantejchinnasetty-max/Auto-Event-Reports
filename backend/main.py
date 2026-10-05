"""Attendance report web app: upload the class-event Excel export, get daily + cumulative reports."""
import io, os, secrets
from pathlib import Path

from fastapi import Depends, FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse, Response
from fastapi.security import HTTPBasic, HTTPBasicCredentials

import report_engine as eng

app = FastAPI(title="Attendance Reports")
security = HTTPBasic(auto_error=False)
FRONTEND = Path(__file__).resolve().parent.parent / "frontend" / "index.html"
MAX_MB = 20


def auth(cred: HTTPBasicCredentials = Depends(security)):
    """Optional login: set APP_PASSWORD (and APP_USER) in the environment to protect the app."""
    pw = os.getenv("APP_PASSWORD")
    if not pw:
        return
    ok = cred and secrets.compare_digest(cred.password, pw) \
        and secrets.compare_digest(cred.username, os.getenv("APP_USER", "admin"))
    if not ok:
        raise HTTPException(401, "Login required", headers={"WWW-Authenticate": "Basic"})


async def read_export(file: UploadFile):
    data = await file.read(MAX_MB * 1024 * 1024 + 1)
    if len(data) > MAX_MB * 1024 * 1024:
        raise HTTPException(413, f"File too large (max {MAX_MB} MB)")
    try:                                   # processed in memory only - nothing is saved to disk
        df = eng.load(io.BytesIO(data))
    except KeyError as e:
        raise HTTPException(400, f"This doesn't look like a class-event export. Missing: {e}")
    except ValueError as e:
        raise HTTPException(400, str(e))
    except Exception as e:
        raise HTTPException(400, f"Could not read the file ({type(e).__name__}). Upload the .xlsx exported from View Students.")
    missing = [c for c in eng.REQUIRED if c not in df.columns]
    if missing or df.empty:
        raise HTTPException(400, f"Missing columns: {missing}" if missing else "The file has no rows.")
    return df


@app.exception_handler(Exception)
async def unexpected(request, exc):
    """Always answer with JSON so the page can show the real reason."""
    print("ERROR:", repr(exc))
    return JSONResponse({"detail": f"Server error: {type(exc).__name__}: {exc}"}, status_code=500)


@app.get("/", dependencies=[Depends(auth)])
def home():
    return FileResponse(FRONTEND)


@app.post("/api/analyze", dependencies=[Depends(auth)])
async def analyze(file: UploadFile = File(...)):
    return eng.summarize(await read_export(file))


@app.post("/api/image", dependencies=[Depends(auth)])
async def image(file: UploadFile = File(...), event_name: str = Form("Class Event")):
    buf = io.BytesIO()
    eng.build_image(await read_export(file), buf, event_name)
    return Response(buf.getvalue(), media_type="image/png",
                    headers={"Content-Disposition": 'attachment; filename="attendance_report.png"'})


@app.post("/api/excel", dependencies=[Depends(auth)])
async def excel(file: UploadFile = File(...), event_name: str = Form("Class Event")):
    buf = io.BytesIO()
    eng.build(await read_export(file), buf, event_name)
    return Response(buf.getvalue(),
                    media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": 'attachment; filename="attendance_report.xlsx"'})
