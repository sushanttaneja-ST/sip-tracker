#!/bin/bash
echo "Stopping SIP Tracker server on port 5001..."
lsof -ti tcp:5001 | xargs kill -9 2>/dev/null && echo "Server stopped." || echo "No server was running."
sleep 1
