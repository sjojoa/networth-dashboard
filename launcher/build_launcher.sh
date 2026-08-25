#!/bin/bash
# Compiles the double-click launcher into "Net Worth Dashboard.app" at the repo
# root. macOS only; everyone else can just run build_dashboard.py.
set -euo pipefail
here="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
root="$(dirname "$here")"
rm -rf "$root/Net Worth Dashboard.app"
osacompile -o "$root/Net Worth Dashboard.app" "$here/NetWorthDashboard.applescript"
echo "Built: $root/Net Worth Dashboard.app"
