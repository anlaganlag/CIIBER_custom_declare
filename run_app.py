#!/usr/bin/env python
"""
Excel Converter Streamlit App Runner

This script launches the Streamlit app for the Excel converter.
It provides a simple way to start the UI without command-line arguments.

Usage:
    python run_app.py
    ./run_app.py (if executable)
"""

import streamlit as st
import subprocess
import sys
import os

def main():
    """Launch the Streamlit app"""
    # Get the directory of this script
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Form the path to the streamlit_app.py file
    streamlit_app_path = os.path.join(current_dir, 'streamlit_app.py')
    
    # Check if the file exists
    if not os.path.exists(streamlit_app_path):
        print(f"Error: Could not find {streamlit_app_path}")
        sys.exit(1)
    
    # Launch the Streamlit app
    print("Starting Excel Converter App...")
    try:
        # Run directly with streamlit
        subprocess.run([sys.executable, "-m", "streamlit", "run", streamlit_app_path],
                      check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error running Streamlit app: {e}")
        sys.exit(1)
    except KeyboardInterrupt:
        print("\nStreamlit app stopped by user.")
        sys.exit(0)

if __name__ == "__main__":
    main() 