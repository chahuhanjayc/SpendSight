#!/bin/bash
echo "Setting up SpendSight for Oracle Cloud..."
python3 -m pip install -r requirements.txt
echo "Starting Gunicorn server on port 8080..."
# JSON-file persistence is not safe across multiple worker processes.
gunicorn -w 1 -b 0.0.0.0:8080 app:app
