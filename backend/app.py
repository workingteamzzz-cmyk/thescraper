"""Social Scraper — local backend (FastAPI).

Runs one scraping job at a time in a background thread. The web UI polls /api/status
for log lines and progress, so there is no browser/request timeout as in Colab.
"""
import importlib.util
import os
import re
import subprocess
import sys
import threading
import time
import traceback

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ENV_FILE = os.path.join(ROOT, ".env")

try:
    from dotenv import load_dotenv
    load_dotenv(ENV_FILE)
except ImportError:
    pass

from . import core, runner  # noqa: E402  (core reads .env on import)

STATIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")
IMPORT_DIR = os.path.join(core.OUTPUT_DIR, "imports")


# ---------------- Log capture: every print() during a job goes into the job log ----------------
_PICTO = re.compile("[\U0001F000-\U0001FAFF\u2600-\u27BF\u2B00-\u2BFF\u23E9-\u23FA\u2139\uFE0F]+ ?")

class Job:
    def __init__(self, params):
        self.params = params
        self.lines = []
        self._partial = ""
        self.lock = threading.Lock()
        self.progress = runner.Progress()
        self.result = {}
        self.running = True
        self.error = None
        self.started = time.time()
        self.finished = None

    def write(self, s):
        s = _PICTO.sub("", s)
        with self.lock:
            buf = self._partial + s
            parts = buf.split("\n")
            self._partial = parts.pop()
            self.lines.extend(parts)

    def flush_partial(self):
        with self.lock:
            if self._partial:
                self.lines.append(self._partial)
                self._partial = ""


JOB = {"current": None}


class _Tee:
    """Replaces sys.stdout: still prints to the terminal and copies into the active job log."""
    def __init__(self, orig):
        self.orig = orig

    def write(self, s):
        try:
            self.orig.write(s)
        except Exception:
            pass
        job = JOB["current"]
        if job is not None and job.running:
            job.write(s)
        return len(s)

    def flush(self):
        try:
            self.orig.flush()
        except Exception:
            pass

    def __getattr__(self, name):
        return getattr(self.orig, name)


sys.stdout = _Tee(sys.stdout)


class _KeepAwake:
    """Stop the laptop from sleeping while a job runs (sleep would drop the connections)."""
    def __enter__(self):
        self.proc = None
        try:
            if sys.platform == "win32":
                import ctypes
                # ES_CONTINUOUS | ES_SYSTEM_REQUIRED
                ctypes.windll.kernel32.SetThreadExecutionState(0x80000000 | 0x00000001)
            elif sys.platform == "darwin":
                self.proc = subprocess.Popen(["caffeinate", "-i", "-w", str(os.getpid())])
        except Exception:
            pass
        return self

    def __exit__(self, *exc):
        try:
            if sys.platform == "win32":
                import ctypes
                ctypes.windll.kernel32.SetThreadExecutionState(0x80000000)
            elif self.proc:
                self.proc.terminate()
        except Exception:
            pass


def _worker(job):
    with _KeepAwake():
        _run(job)


def _run(job):
    try:
        runner.run_job(job.params, job.progress, job.result)
    except core.StopRequested:
        print("\nStopped by user.")
    except Exception as e:
        job.error = f"{type(e).__name__}: {e}"
        print("\n❌ Error: " + job.error)
        print(traceback.format_exc())
    finally:
        job.flush_partial()
        if core.STOP["flag"]:
            job.progress.bar_style = "warning"
        elif job.error:
            job.progress.bar_style = "danger"
        job.finished = time.time()
        job.running = False
        core.STOP["flag"] = False


# ---------------- App ----------------
app = FastAPI(title="Social Scraper (local)")


def _mask(k):
    return (k[:4] + "…" + k[-4:]) if len(k) > 10 else ("set" if k else "")


@app.get("/api/config")
def get_config():
    return {
        "modes": [core.LINK_MODE, core.PROFILE_MODE, core.MIXED_MODE, "ALL platforms"]
                 + list(core.PLATFORMS.keys()) + [core.IG_KW_PLATFORM, core.FB_PLATFORM, core.TW_PLATFORM],
        "link_mode": core.LINK_MODE, "profile_mode": core.PROFILE_MODE, "mixed_mode": core.MIXED_MODE,
        "columns": runner.ALL_COLS,
        "defaults": {**runner.DEFAULTS, "columns": runner.ALL_COLS},
        "api_key": _mask(core.API_KEY),
        "twitter_key": _mask(core.TWITTER_API_KEY),
        "output_dir": core.OUTPUT_DIR,
        "ai_available": all(importlib.util.find_spec(m) for m in ("transformers", "torch")),
    }


def _write_env(updates):
    lines = []
    if os.path.exists(ENV_FILE):
        with open(ENV_FILE, encoding="utf-8") as f:
            lines = f.read().splitlines()
    for key, val in updates.items():
        new = f"{key}={val}"
        for i, l in enumerate(lines):
            if re.match(rf"\s*{key}\s*=", l):
                lines[i] = new
                break
        else:
            lines.append(new)
    with open(ENV_FILE, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


@app.post("/api/settings")
def save_settings(body: dict):
    api_key = (body.get("api_key") or "").strip()
    tw_key = (body.get("twitter_key") or "").strip()
    updates = {}
    if api_key:
        core.set_api_keys(api_key=api_key)
        updates["SCRAPECREATORS_API_KEY"] = api_key
    if tw_key:
        core.set_api_keys(twitter_key=tw_key)
        updates["TWITTER_API_KEY"] = tw_key
    if updates and body.get("remember", True):
        _write_env(updates)
    return {"api_key": _mask(core.API_KEY), "twitter_key": _mask(core.TWITTER_API_KEY)}


@app.get("/api/credits")
def credits():
    if not core.API_KEY:
        return {"error": "API key not set"}
    return core.check_credits()


@app.post("/api/run")
def run(body: dict):
    cur = JOB["current"]
    if cur is not None and cur.running:
        raise HTTPException(409, "A scrape is already running. Wait for it to finish or press Stop.")
    core.STOP["flag"] = False
    job = Job(body or {})
    JOB["current"] = job
    threading.Thread(target=_worker, args=(job,), daemon=True).start()
    return {"ok": True}


@app.post("/api/stop")
def stop():
    job = JOB["current"]
    if job is None or not job.running:
        return {"ok": False}
    core.STOP["flag"] = True
    print("\nStop requested \u2014 finishing up and saving the data collected so far...")
    return {"ok": True}


@app.get("/api/status")
def status(since: int = 0):
    job = JOB["current"]
    if job is None:
        return {"running": False, "lines": [], "next": 0, "progress": None, "result": {}}
    with job.lock:
        lines = job.lines[since:]
        nxt = len(job.lines)
        partial = job._partial
    return {
        "running": job.running,
        "stopping": core.STOP["flag"],
        "lines": lines,
        "partial": partial,
        "next": nxt,
        "progress": job.progress.to_dict(),
        "result": job.result if not job.running else {},
        "error": job.error,
        "elapsed": round((job.finished or time.time()) - job.started),
    }


def _safe_name(name):
    name = os.path.basename(name)
    if not name or name.startswith("."):
        raise HTTPException(400, "invalid file name")
    return name


@app.get("/api/files")
def list_files():
    os.makedirs(core.OUTPUT_DIR, exist_ok=True)
    out = []
    for n in os.listdir(core.OUTPUT_DIR):
        p = os.path.join(core.OUTPUT_DIR, n)
        if os.path.isfile(p) and n.lower().endswith((".xlsx", ".csv")):
            st = os.stat(p)
            out.append({"name": n, "size": st.st_size, "mtime": st.st_mtime})
    out.sort(key=lambda x: -x["mtime"])
    return out


@app.get("/api/download/{name}")
def download(name: str):
    p = os.path.join(core.OUTPUT_DIR, _safe_name(name))
    if not os.path.isfile(p):
        raise HTTPException(404, "file not found")
    return FileResponse(p, filename=os.path.basename(p))


@app.delete("/api/files/{name}")
def delete_file(name: str):
    p = os.path.join(core.OUTPUT_DIR, _safe_name(name))
    if os.path.isfile(p):
        os.remove(p)
    return {"ok": True}


@app.post("/api/upload")
async def upload(file: UploadFile = File(...)):
    """Upload an .xlsx/.csv for the 'Import file(s)' box; returns its local path."""
    name = _safe_name(file.filename or "")
    if not name.lower().endswith((".xlsx", ".xls", ".csv")):
        raise HTTPException(400, "only .xlsx / .xls / .csv files")
    os.makedirs(IMPORT_DIR, exist_ok=True)
    p = os.path.join(IMPORT_DIR, name)
    with open(p, "wb") as f:
        f.write(await file.read())
    return {"path": p}


@app.get("/")
def index():
    return FileResponse(os.path.join(STATIC_DIR, "index.html"))


app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


@app.exception_handler(Exception)
async def _err(request, exc):
    return JSONResponse({"detail": f"{type(exc).__name__}: {exc}"}, status_code=500)
