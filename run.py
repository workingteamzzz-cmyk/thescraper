"""Run Social Scraper on your laptop: python run.py  (opens the browser automatically)."""
import os
import sys
import threading
import webbrowser

HOST = os.environ.get("HOST", "127.0.0.1")
PORT = int(os.environ.get("PORT", "8765"))


def main():
    try:
        import uvicorn
    except ImportError:
        sys.exit("Dependencies are missing. Run first:  pip install -r requirements.txt")
    url = f"http://{'localhost' if HOST in ('127.0.0.1', '0.0.0.0') else HOST}:{PORT}"
    print(f"\n  Social Scraper is running at {url}  (close this window or press Ctrl+C to quit)\n")
    if "--no-browser" not in sys.argv:
        threading.Timer(1.5, lambda: webbrowser.open(url)).start()
    uvicorn.run("backend.app:app", host=HOST, port=PORT, log_level="warning")


if __name__ == "__main__":
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    main()
