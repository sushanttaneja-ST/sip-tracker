#!/bin/bash
cd "/Users/sushant/iCloud Drive (Archive)/Desktop/Sushant/CLAUDE CODE/MUTUAL FUND TRACKER"
python3 app.py &
sleep 2
open http://localhost:5001
echo ""
echo "Dashboard is running at http://localhost:5001"
echo "Close this window to stop the server."
wait
