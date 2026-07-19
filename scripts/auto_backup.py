#!/usr/bin/env python3
"""
Auto git backup on file change.
Usage: python3 auto_backup.py [--debounce SECONDS] [--remote NAME] [--branch NAME]
"""

import os
import sys
import time
import subprocess
import threading
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler


# Project root is two levels up from this script (scripts/auto_backup.py -> scripts/ -> project root)
PROJECT_ROOT = Path(__file__).parent.parent.resolve()
# Defaults, can be overridden by env or CLI
DEBOUNCE_SECONDS = int(os.environ.get("DEBOUNCE_SECONDS", "10"))
GIT_REMOTE = os.environ.get("GIT_REMOTE", "origin")
GIT_BRANCH = os.environ.get("GIT_BRANCH", "main")

EXCLUDE_PATTERNS = [
    ".git",
    "__pycache__",
    "*.pyc",
    ".venv",
    "venv",
    "env",
    "*.webp", "*.jpg", "*.jpeg", "*.png", "*.pdf",
    "*.tiff", "*.bmp", "*.heic",
    "*.log",
    "*.webm", "*.mp4", "*.mov",
    "node_modules",
    ".pytest_cache",
    ".coverage",
    "htmlcov",
    "dist",
    "build",
    "*.egg-info",
    "OCRCola_ToOCR",  # project-specific: data folder
    "ocr_results",
    "checkpoints",
    "models",
    "data",
]


class DebouncedHandler(FileSystemEventHandler):
    def __init__(self):
        self.timer = None
        self.pending = False
        self.last_event = 0

    def _should_exclude(self, path: Path) -> bool:
        try:
            rel = path.relative_to(PROJECT_ROOT)
        except ValueError:
            # Outside project root, exclude
            return True
        rel_str = str(rel)
        for pattern in EXCLUDE_PATTERNS:
            if pattern.startswith("*."):
                if rel_str.endswith(pattern[1:]):
                    return True
            elif pattern in rel_str:
                return True
        return False

    def on_any_event(self, event):
        if event.is_directory:
            return
        src = Path(event.src_path)
        if self._should_exclude(src):
            return
        self.last_event = time.time()
        if not self.pending:
            self.pending = True
            self.timer = threading.Timer(DEBOUNCE_SECONDS, self.commit_and_push)
            self.timer.start()

    def commit_and_push(self):
        self.pending = False
        if time.time() - self.last_event < DEBOUNCE_SECONDS:
            self.timer = threading.Timer(DEBOUNCE_SECONDS, self.commit_and_push)
            self.timer.start()
            return

        print(f"\n[{time.strftime('%H:%M:%S')}] Change detected, committing...")
        try:
            subprocess.run(["git", "add", "-A"], cwd=PROJECT_ROOT, check=True, capture_output=True)
            result = subprocess.run(["git", "status", "--porcelain"], cwd=PROJECT_ROOT, capture_output=True, text=True)
            if not result.stdout.strip():
                print("  No changes to commit")
                return

            msg = f"auto: backup {time.strftime('%Y-%m-%d %H:%M:%S')}"
            subprocess.run(["git", "commit", "-m", msg], cwd=PROJECT_ROOT, check=True, capture_output=True)
            print(f"  ✅ Committed: {msg}")
            subprocess.run(["git", "push", GIT_REMOTE, GIT_BRANCH], cwd=PROJECT_ROOT, check=True, capture_output=True)
            print(f"  ✅ Pushed to {GIT_REMOTE}/{GIT_BRANCH}")
        except subprocess.CalledProcessError as e:
            print(f"  ❌ Error: {e.stderr.decode() if e.stderr else e}")


def main():
    import argparse
    global DEBOUNCE_SECONDS, GIT_REMOTE, GIT_BRANCH
    parser = argparse.ArgumentParser()
    parser.add_argument("--debounce", type=int, default=DEBOUNCE_SECONDS)
    parser.add_argument("--remote", default=GIT_REMOTE)
    parser.add_argument("--branch", default=GIT_BRANCH)
    args = parser.parse_args()

    DEBOUNCE_SECONDS = args.debounce
    GIT_REMOTE = args.remote
    GIT_BRANCH = args.branch

    if not (PROJECT_ROOT / ".git").exists():
        print("❌ Not a git repository")
        sys.exit(1)

    # Verify remote exists
    result = subprocess.run(["git", "remote"], cwd=PROJECT_ROOT, capture_output=True, text=True)
    if GIT_REMOTE not in result.stdout:
        print(f"❌ Remote '{GIT_REMOTE}' not found. Run: git remote add {GIT_REMOTE} <url>")
        sys.exit(1)

    print(f"🔍 Watching: {PROJECT_ROOT}")
    print(f"⏱️  Debounce: {DEBOUNCE_SECONDS}s")
    print(f"🚀 Push to: {GIT_REMOTE}/{GIT_BRANCH}")
    print("Press Ctrl+C to stop\n")

    handler = DebouncedHandler()
    observer = Observer()
    observer.schedule(handler, str(PROJECT_ROOT), recursive=True)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n🛑 Stopping...")
        observer.stop()
    observer.join()


if __name__ == "__main__":
    main()