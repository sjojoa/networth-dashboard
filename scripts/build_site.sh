#!/bin/sh
# Build the files Cloudflare serves into dist/.
#
# Rebuilds from the workbook when Python is available; otherwise uses the
# committed NetWorth_Dashboard.html, which is always rebuilt before a commit.
# Either way it's the SAMPLE workbook in this repo — never commit a real one,
# because whatever is here gets published.
set -e
cd "$(dirname "$0")/.."
if command -v python3 >/dev/null 2>&1; then
  python3 build_dashboard.py NetWorthTracker.xlsx
fi
rm -rf dist
mkdir -p dist
cp NetWorth_Dashboard.html dist/index.html
echo "dist/ ready: $(wc -c < dist/index.html | tr -d ' ') bytes"
