@echo off
cd /d "%~dp0"
echo Starting Job Readiness Dashboard...
echo.
echo When ready, open your browser at: http://localhost:8501
echo Press Ctrl+C to stop the dashboard.
echo.
streamlit run app.py
if errorlevel 1 (
    echo.
    echo ERROR: Dashboard failed to start. See message above.
    pause
)
