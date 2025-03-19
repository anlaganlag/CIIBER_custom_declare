#!/bin/bash

# Stop on errors
set -e

# Get directory of this script
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

echo "Changing to project directory: $SCRIPT_DIR"
cd "$SCRIPT_DIR"

echo "Setting up Python virtual environment..."
python3 -m venv venv

echo "Installing requirements..."
if [ -f requirements.txt ]; then
    "$SCRIPT_DIR/venv/bin/pip" install -r requirements.txt  # 明确使用虚拟环境的pip
else
    echo "Error: requirements.txt not found!"
    exit 1
fi

echo "Activating virtual environment..."
source /Users/ciiber/Documents/code/CIIBER_custom_declare/venv/bin/activate

echo "Starting Streamlit application..."
streamlit run /Users/ciiber/Documents/code/CIIBER_custom_declare/streamlit_app.py
