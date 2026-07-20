#!/usr/bin/env bash
set -euo pipefail
PROJECT_ROOT="${1:-$(pwd)}"
REMOTE="${2:-origin}"
BRANCH="${3:-main}"
DEBOUNCE="${4:-5}"
cd "$PROJECT_ROOT"
[[ ! -d .git ]] && { echo "❌ Not a git repo"; exit 1; }

# Exclude patterns - TestAll specific
EXCLUDE_REGEX='(\.git|__pycache__|\.pyc$|\.venv|venv|env|\.webp$|\.jpg$|\.jpeg$|\.png$|\.pdf$|\.tiff$|\.bmp$|\.heic$|data_cache|backtest_results|backtest_logs|logs|\.log$|\.db$|\.sqlite$|node_modules|\.pytest_cache|dist|build|\.egg-info|ocr_results|E:|checkpoints|models|data)'

echo "🔍 Watching: $PROJECT_ROOT"
echo "⏱️  Debounce: ${DEBOUNCE}s"
echo "🚀 Push to: $REMOTE/$BRANCH"

last_event=0
timer_pid=0

commit_and_push() {
    if ! git diff --quiet || ! git diff --cached --quiet || [[ -n "$(git ls-files --others --exclude-standard)" ]]; then
        msg="auto: backup $(date '+%Y-%m-%d %H:%M:%S')"
        git add -A
        git commit -m "$msg"
        git push "$REMOTE" "$BRANCH"
        echo "  ✅ $msg"
    else
        echo "  No changes to commit"
    fi
}

inotifywait -m -r -e modify,create,delete,move \
    --exclude "$EXCLUDE_REGEX" \
    --format '%w%f %e' "$PROJECT_ROOT" | while read -r file event; do

    now=$(date +%s)
    last_event=$now
    [[ $timer_pid -ne 0 ]] && kill "$timer_pid" 2>/dev/null

    ( sleep "$DEBOUNCE"
      [[ $(date +%s) -ge $((last_event + DEBOUNCE)) ]] && commit_and_push
    ) &
    timer_pid=$!
done
