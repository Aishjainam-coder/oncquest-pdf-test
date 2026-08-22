#!/usr/bin/env bash
# Removes everything the isolated setup created. Project source is untouched.
set -u
cd "$(dirname "$0")"

pkill -f "\.venv-oncquest/bin/streamlit run app\.py" 2>/dev/null && echo "stopped running app"

if [ -d .venv-oncquest ]; then
  # Remove Playwright browsers this venv downloaded (~95 MB, in ~/Library/Caches/ms-playwright)
  .venv-oncquest/bin/playwright uninstall --all 2>/dev/null || true
  rm -rf .venv-oncquest && echo "removed .venv-oncquest"
fi

rm -rf __pycache__ && echo "removed __pycache__"
echo "done."
