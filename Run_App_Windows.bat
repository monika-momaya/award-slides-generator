@echo off
title M2M Programme Planner
echo.
echo  ================================
echo    M2M Programme Planner v4
echo  ================================
echo.
python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo  ERROR: Python not found.
    echo  Install from https://www.python.org — tick "Add Python to PATH"
    pause & exit /b 1
)
echo  Checking dependencies...
python -m pip install streamlit openpyxl pandas python-docx --quiet
echo  Starting app at http://localhost:8501
echo.
python -m streamlit run m2m_app.py --server.headless false
pause
