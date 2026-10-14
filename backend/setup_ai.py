"""One-time setup for the AI analysis mode: installs the packages and downloads the models.

Run by the start scripts on first launch (python -m backend.setup_ai). Safe to run again;
if anything fails the app still works and falls back to the lexicon.
"""
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MARKER = os.path.join(ROOT, ".venv", ".ai_ready")
TRIED = os.path.join(ROOT, ".venv", ".ai_tried")   # start scripts skip setup once attempted


def _pip(*args):
    return subprocess.run([sys.executable, "-m", "pip", "install", "--disable-pip-version-check", *args]).returncode == 0


def _have_packages():
    try:
        import torch  # noqa: F401
        import transformers  # noqa: F401
        return True
    except ImportError:
        return False


def main():
    os.makedirs(os.path.dirname(TRIED), exist_ok=True)
    open(TRIED, "w").close()
    print("Setting up AI sentiment & emotion analysis (one time only).")
    print("This downloads ~2-3 GB and can take 5-15 minutes depending on your internet.\n")
    if not _have_packages():
        print("[1/2] Installing AI packages...")
        if sys.platform != "darwin":
            # CPU-only torch is much smaller than the default CUDA build on Windows/Linux.
            _pip("--index-url", "https://download.pytorch.org/whl/cpu", "torch") or _pip("torch")
        if not _pip("-r", os.path.join(ROOT, "requirements-ai.txt")) or not _have_packages():
            print("\nCould not install the AI packages. The app will use the lexicon analysis.")
            print("You can retry later with install-ai-windows.bat / install-ai-mac.command.")
            return 1
    print("[2/2] Downloading AI models...")
    from backend import core
    if not core._load_ai_models():
        print("\nCould not download the AI models (check your internet). The app will use the lexicon")
        print("analysis for now and try the models again on the next scrape.")
        return 1
    os.makedirs(os.path.dirname(MARKER), exist_ok=True)
    open(MARKER, "w").close()
    print("\nAI analysis is ready.")
    return 0


if __name__ == "__main__":
    os.chdir(ROOT)
    sys.exit(main())
