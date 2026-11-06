#!/bin/bash
set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"
command -v conda >/dev/null || { echo "Install conda or use the Python/venv instructions in README.md" >&2; exit 1; }
conda env update --file "$SCRIPT_DIR/environment.yml"
conda run -n pyexcel-forecast python -m xlwings addin install
echo "Configure xlwings to use pyexcel-forecast and import $SCRIPT_DIR/pyexcel_macros.bas"
