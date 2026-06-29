M2M PROGRAMME PLANNER v4
=========================

CHANGES IN v4:
  • Word doc completely redesigned to match BTS 2023 MC Copy format:
      - Title centred in maroon
      - Date/venue in grey below title
      - Logo centred above title (if uploaded)
      - Clean table: Timings | : | Programme Details — NO colour fills
      - MC Script on page 2 with bilingual side-by-side layout
  • Excel: no serial number, Time Slot first, clean 2-column layout
  • Logo embedded top-right in Excel; centred in Word

LOCAL:
  Windows → double-click Run_App_Windows.bat
  Mac     → double-click Run_App_Mac.command

STREAMLIT CLOUD (free, shareable):
  Upload m2m_app.py + requirements.txt to GitHub
  Deploy at share.streamlit.io
  NOTE: Word export requires Node.js on the server.
  Streamlit Cloud includes Node.js — works out of the box.

IMPORTANT FOR WORD DOWNLOAD:
  The Word export uses Node.js (docx library) for professional output.
  Node.js is pre-installed on Streamlit Cloud.
  For local use: Node.js must be installed (https://nodejs.org)
  Download "LTS" version — one-time install.
