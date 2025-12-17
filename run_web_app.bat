@echo off
echo ========================================
echo  Slide Translator Web Interface
echo  Project C
echo ========================================
echo.
echo Starting web server...
echo Browser will open automatically at http://localhost:8501
echo.
echo Press Ctrl+C to stop the server
echo.

uv run streamlit run streamlit_app.py
