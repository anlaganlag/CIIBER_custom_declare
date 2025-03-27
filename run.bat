@echo off
setlocal enabledelayedexpansion

echo Setting up Python virtual environment...
python -m venv venv

echo Installing requirements...
if exist requirements.txt (
    .\venv\Scripts\pip install -r requirements.txt
) else (
    echo Error: requirements.txt not found!
    exit /b 1
)

echo Activating virtual environment...
call .\venv\Scripts\activate.bat

echo Starting Streamlit application...
streamlit run streamlit_app.py 