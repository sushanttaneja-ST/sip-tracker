#!/bin/bash
# SIP Tracker — startup script
# Run this every morning: bash run.sh

cd "$(dirname "$0")"

# Install dependencies if needed
if ! python3 -c "import flask" 2>/dev/null; then
  echo "Installing dependencies..."
  pip3 install -r requirements.txt -q
fi

echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "  📈  SIP Tracker Dashboard"
echo "  Open → http://localhost:5001"
echo "  Stop → Ctrl+C"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""

open http://localhost:5001
python3 app.py
