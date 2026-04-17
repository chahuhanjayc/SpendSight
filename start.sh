#!/bin/bash

echo ""
echo "================================================"
echo "           SpendSight — Starting Up"
echo "================================================"
echo ""

# Check for Python 3
if ! command -v python3 &> /dev/null; then
    echo "  ERROR: Python 3 is not installed."
    echo ""
    echo "  Please install it first:"
    echo "  https://www.python.org/downloads/"
    echo ""
    echo "  Or if you have Homebrew:"
    echo "  brew install python3"
    echo ""
    exit 1
fi

echo "  Python found: $(python3 --version)"
echo ""

# Install requirements
echo "  Checking requirements..."
python3 -m pip install -r requirements.txt --quiet
echo "  Requirements ready."
echo ""

# Run the app
echo "  Open http://localhost:5000 in your browser"
echo "  Press Ctrl+C to stop"
echo "================================================"
echo ""
python3 app.py
