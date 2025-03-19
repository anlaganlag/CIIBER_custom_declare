#!/bin/bash

# Stop on errors
set -e

# Get directory of this script
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

echo "Changing to project directory: $SCRIPT_DIR"
cd "$SCRIPT_DIR"

echo "Setting up Python virtual environment..."
python3 -m venv venv

echo "Activating virtual environment..."
source /Users/ciiber/Documents/code/CIIBER-shipping-list/venv/bin/activate

echo "Installing requirements..."
if [ -f requirements.txt ]; then
    pip install -r requirements.txt
else
    echo "Error: requirements.txt not found!"
    exit 1
fi

echo "Starting Streamlit application..."
streamlit run  /Users/ciiber/Documents/code/CIIBER-shipping-list/app.py  # Replace with your main app filename if different 
