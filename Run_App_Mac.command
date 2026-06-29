#!/bin/bash
cd "$(dirname "$0")"
echo ""
echo "================================"
echo "  M2M Programme Planner v4"
echo "================================"
echo ""
if ! command -v python3 &>/dev/null; then
    echo "❌  Python not found. Install from https://www.python.org"
    read -p "Press Enter to close..."; exit 1
fi
echo "  Checking dependencies..."
python3 -m pip install streamlit openpyxl pandas python-docx --quiet
echo "  Starting app at http://localhost:8501"
echo ""
python3 -m streamlit run m2m_app.py --server.headless false
