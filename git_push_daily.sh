#!/bin/bash
# 每日執行 generate_tsmc_html_report.py 後，執行此腳本上傳 Git
# 用法：bash git_push_daily.sh

cd "$(dirname "$0")"

TODAY=$(python3 -c "from datetime import date; print(date.today().isoformat())")
HTML="TSMC_日報_${TODAY}.html"

echo "=== TSMC Daily Git Push ==="
echo "Date  : ${TODAY}"
echo "Report: ${HTML}"
echo ""

# 確認 HTML 檔案存在
if [ ! -f "${HTML}" ]; then
  echo "ERROR: ${HTML} not found. Did you run generate_tsmc_html_report.py first?"
  exit 1
fi

# Stage 3 files
git add generate_tsmc_html_report.py "${HTML}" manifest.json

# Show what will be committed
echo "--- Staged files ---"
git status --short

# Commit
git commit -m "daily: ${TODAY} 日報"

# Push
git push origin master

echo ""
echo "Done. GitHub Pages will update in ~1 minute."
echo "URL: https://tomwu229622.github.io/tsmc-daily-report/"
