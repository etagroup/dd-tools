#!/usr/bin/env bash
set -euo pipefail

if [ ! -d "venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
fi

source venv/bin/activate
echo "Installing dependencies..."
pip install --quiet pandas openpyxl numpy

echo ""
echo "Setup complete. Activate with:"
echo "  source venv/bin/activate"
