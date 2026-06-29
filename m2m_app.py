"""
M2M Programme Planner — Streamlit Web App v2
=============================================
New in v2:
  • Logo upload → displayed beside event details
  • File renamed to "M2M Programme for <<Event Name>>"
  • Word doc (.docx) download option
  • Excel download (existing)

Run locally:   streamlit run m2m_app_v2.py
Deploy:        Push to GitHub → share.streamlit.io
"""

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io, base64
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="M2M Programme Planner",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { background-color: #FAFAFA; }
    .stApp { font-family: 'Arial', sans-serif; }
    .title-banner {
        background: linear-gradient(135deg, #7B1B1B 0%, #A52828 100%);
        color: white; padding: 1.2rem 2rem; border-radius: 10px;
        margin-bottom: 1.5rem; box-shadow: 0 4px 12px rgba(123,27,27,0.3);
        display: flex; align-items: center; gap: 1.5rem;
    }
    .title-banner-text h1 { color: white; margin: 0; font-size: 1.8rem; }
    .title-banner-text p  { color: #F5E6C8; margin: 0.3rem 0 0 0; font-size: 0.95rem; }
    .title-banner-logo img { max-height: 70px; border-radius: 6px; }
    .section-header {
        background: #C9A84C; color: white; padding: 0.5rem 1rem;
        border-radius: 6px; font-weight: bold; margin: 1rem 0 0.5rem 0; font-size: 0.95rem;
    }
    .time-pill {
        background: #F2DADA; color: #7B1B1B; padding: 0.2rem 0.6rem;
        border-radius: 20px; font-weight: bold; font-size: 0.85rem; font-family: monospace;
    }
    .script-box {
        background: #F8F9FA; border-left: 4px solid #7B1B1B;
        padding: 0.8rem 1rem; border-radius: 0 6px 6px 0;
        margin: 0.3rem 0; font-size: 0.88rem; line-height: 1.6;
    }
    .script-box-kn {
        background: #FFF8EE; border-left: 4px solid #C9A84C;
        padding: 0.8rem 1rem; border-radius: 0 6px 6px 0;
        margin: 0.3rem 0; font-size: 0.88rem; line-height: 1.6;
    }
    .summary-card {
        background: #2C2C2C; color: #C9A84C; padding: 1rem 1.5rem;
        border-radius: 8px; text-align: center; font-weight: bold;
    }
    .event-info-panel {
        background: white; border: 1px solid #F5E6C8; border-radius: 8px;
        padding: 1rem 1.5rem; margin-bottom: 1rem;
        display: flex; align-items: center; gap: 1.5rem;
    }
    .stDownloadButton > button {
        background: #7B1B1B !important; color: white !important;
        border: none !important; border-radius: 6px !important;
        font-weight: bold !important;
    }
    .stDownloadButton > button:hover { background: #A52828 !important; }
    .dl-word > button { background: #1F4E79 !important; }
    .dl-word > button:hover { background: #2E75B6 !important; }
</style>
""", unsafe_allow_html=True)

# ── Script library ────────────────────────────────────────────────────────────
SCRIPTS = [
    (["welcome","dais","dignitar"],
     "Respected dignitaries, honoured guests, and friends — a very warm welcome to you all.\n"
     "We now request our distinguished guests to kindly proceed to the dais and take their seats.\n"
     "[MC reads out names one by one as each dignitary is escorted to the stage]",
     "ಗೌರವಾನ್ವಿತ ಅತಿಥಿಗಳೇ, ಗಣ್ಯ ಮಹನೀಯರೇ ಮತ್ತು ಆತ್ಮೀಯ ಬಂಧುಗಳೇ — ನಿಮಗೆ ಹೃತ್ಪೂರ್ವಕ ಸ್ವಾಗತ.\n"
     "ನಾವು ಗಣ್ಯ ಅತಿಥಿಗಳನ್ನು ವೇದಿಕೆಯಲ್ಲಿ ಆಸೀನರಾಗಲು ವಿನಂತಿಸುತ್ತೇವೆ.\n"
     "[ಪ್ರತಿಯೊಬ್ಬ ಅತಿಥಿಯ ಹೆಸರನ್ನು ಓದಿ ಅವರನ್ನು ಸ್ವಾಗತಿಸಿ]"),
    (["naada","nadageethe","state anthem"],
     "We shall now commence with the Naada Geethe — the State Anthem of Karnataka.\n"
     "I request all those present to please rise.\n[Naada Geethe plays]\nThank you. Please be seated.",
     "ನಾವು ಕರ್ನಾಟಕ ನಾಡಗೀತೆಯೊಂದಿಗೆ ಕಾರ್ಯಕ್ರಮವನ್ನು ಆರಂಭಿಸುತ್ತೇವೆ.\n"
     "ಎಲ್ಲರೂ ದಯವಿಟ್ಟು ಎದ್ದು ನಿಲ್ಲಬೇಕೆಂದು ವಿನಂತಿಸುತ್ತೇನೆ.\n"
     "[ನಾಡಗೀತೆ ಹಾಡಲಾಗುವುದು]\nಧನ್ಯವಾದಗಳು. ದಯವಿಟ್ಟು ಕುಳಿತುಕೊಳ್ಳಿ."),
    (["national anthem","jana gana","jai hind"],
     "We will now conclude with the National Anthem.\nI request everyone to please rise.\n"
     "[National Anthem plays]\nJai Hind! Jai Karnataka!\n"
     "Thank you all. The programme now stands concluded.",
     "ರಾಷ್ಟ್ರಗೀತೆಯೊಂದಿಗೆ ಕಾರ್ಯಕ್ರಮವನ್ನು ಮುಕ್ತಾಯಗೊಳಿಸುತ್ತೇವೆ.\n"
     "ಎಲ್ಲರೂ ದಯವಿಟ್ಟು ಎದ್ದು ನಿಲ್ಲಬೇಕೆಂದು ವಿನಂತಿಸುತ್ತೇನೆ.\n"
     "[ರಾಷ್ಟ್ರಗೀತೆ]\nಜೈ ಹಿಂದ್! ಜೈ ಕರ್ನಾಟಕ!"),
    (["lamp","lighting","inaugur","deepa"],
     "We will now proceed to the auspicious lighting of the lamp.\n"
     "I request [Chief Guest Name & Designation] and the distinguished guests\n"
     "to kindly come forward for the lamp lighting ceremony.",
     "ಈಗ ದೀಪ ಪ್ರಜ್ವಲನ ಕಾರ್ಯಕ್ರಮ ನಡೆಯಲಿದೆ.\n"
     "[ಮುಖ್ಯ ಅತಿಥಿ ಹೆಸರು] ಮತ್ತು ಗಣ್ಯರನ್ನು ದೀಪ ಬೆಳಗಿಸಲು ವಿನಂತಿಸುತ್ತೇವೆ."),
    (["welcome address","welcome speech"],
     "We will now have the Welcome Address.\n"
     "I request [Name], [Designation], to kindly address the gathering.\n"
     "[After speech] We thank [Name] for those inspiring words.",
     "ಈಗ ಸ್ವಾಗತ ಭಾಷಣ ನಡೆಯಲಿದೆ.\n"
     "[ಹೆಸರು], [ಹುದ್ದೆ] ಅವರನ್ನು ಸಭೆಯನ್ನು ಉದ್ದೇಶಿಸಿ ಮಾತನಾಡಲು ವಿನಂತಿಸುತ್ತೇವೆ."),
    (["keynote","inaugural address","chief guest"],
     "We are privileged to have the Keynote / Inaugural Address.\n"
     "I request the Chief Guest, [Full Name & Designation],\nto kindly deliver the Address.\n"
     "[After speech] Please join me in a round of applause.",
     "ಈಗ ಮುಖ್ಯ ಭಾಷಣ ನಡೆಯಲಿದೆ.\n"
     "[ಪೂರ್ಣ ಹೆಸರು ಮತ್ತು ಹುದ್ದೆ] ಅವರನ್ನು ಭಾಷಣ ಮಾಡಲು ವಿನಂತಿಸುತ್ತೇವೆ.\n"
     "[ಭಾಷಣದ ನಂತರ] ಚಪ್ಪಾಳೆಯೊಂದಿಗೆ ಅಭಿನಂದಿಸೋಣ."),
    (["perspective","industry","context setting","biotech","startup"],
     "We will now hear the perspective from [Name], [Designation].\n"
     "[After speech] Thank you, [Name], for those valuable insights.",
     "ಈಗ [ಹೆಸರು], [ಹುದ್ದೆ] ಅವರಿಂದ ದೃಷ್ಟಿಕೋನ ಕೇಳಲಿದ್ದೇವೆ.\n"
     "[ಭಾಷಣದ ನಂತರ] ಧನ್ಯವಾದಗಳು, [ಹೆಸರು] ಅವರಿಗೆ."),
    (["introduction","recorded message","ambassador","h.e."],
     "We are privileged to have an introduction by [Name], [Designation],\n"
     "followed by a recorded message from [Speaker Name].",
     "ಈಗ [ಹೆಸರು] ಅವರಿಂದ ಪರಿಚಯ ಮತ್ತು [ಸಂದೇಶ ನೀಡಿದ ಹೆಸರು] ಅವರ ಸಂದೇಶ ಪ್ರಸ್ತುತಿಯಾಗಲಿದೆ."),
    (["release","policy","souvenir","publication","launch"],
     "We will now proceed to the release of [Name of Publication / Policy].\n"
     "I request [Chief Guest / Name] and the distinguished guests to come forward.",
     "ಈಗ [ಪ್ರಕಾಶನ / ನೀತಿ ಹೆಸರು] ಬಿಡುಗಡೆ ಮಾಡಲಾಗುವುದು.\n"
     "[ಮುಖ್ಯ ಅತಿಥಿ] ಮತ್ತು ಗಣ್ಯರನ್ನು ಮುಂದೆ ಬರಲು ಕೋರುತ್ತೇವೆ."),
    (["felicitat","honour","award","memento"],
     "We will now proceed to the felicitation of our distinguished guests.\n"
     "I request [Presenter], [Designation], to kindly felicitate [Guest Name].",
     "ಈಗ ಗಣ್ಯ ಅತಿಥಿಗಳ ಸನ್ಮಾನ ಕಾರ್ಯಕ್ರಮ ನಡೆಯಲಿದೆ.\n"
     "[ಹೆಸರು] ಅವರನ್ನು [ಸನ್ಮಾನಿಸಲ್ಪಡುವ ಹೆಸರು] ಅವರನ್ನು ಸನ್ಮಾನಿಸಲು ವಿನಂತಿಸುತ್ತೇವೆ."),
    (["cultural","dance","music","performance","song"],
     "We will now be treated to a cultural performance by [Name / Group].\nPlease enjoy.",
     "ಈಗ [ಕಲಾವಿದ / ತಂಡ] ಅವರಿಂದ ಸಾಂಸ್ಕೃತಿಕ ಕಾರ್ಯಕ್ರಮ ಪ್ರಸ್ತುತಿಯಾಗಲಿದೆ.\nದಯವಿಟ್ಟು ಆನಂದಿಸಿ."),
    (["vote of thanks","vote of thank"],
     "We will now have the Vote of Thanks.\n"
     "I request [Name], [Designation], to kindly propose the Vote of Thanks.",
     "ಈಗ ವಂದನಾರ್ಪಣೆ ನಡೆಯಲಿದೆ.\n"
     "[ಹೆಸರು], [ಹುದ್ದೆ] ಅವರನ್ನು ವಂದನಾರ್ಪಣೆ ಸಲ್ಲಿಸಲು ವಿನಂತಿಸುತ್ತೇವೆ."),
    (["tea","coffee","lunch","break","networking","refreshment"],
     "We will now take a short break. The next session commences at [Time].\n"
     "Refreshments are available at [Location].",
     "ನಾವು ಈಗ ವಿರಾಮ ತೆಗೆದುಕೊಳ್ಳಲಿದ್ದೇವೆ.\nಮುಂದಿನ ಅಧಿವೇಶನ [ಸಮಯ]ಕ್ಕೆ ಪ್ರಾರಂಭವಾಗಲಿದೆ."),
    (["panel","discussion","roundtable","session"],
     "We will now move to the Panel Discussion.\n"
     "Moderator: [Name, Designation]. I request the panellists to take their seats.",
     "ಈಗ ಸಮಿತಿ ಚರ್ಚೆ ನಡೆಯಲಿದೆ.\nನಿರ್ವಾಹಕರು: [ಹೆಸರು, ಹುದ್ದೆ]."),
    (["address by","address from","address"],
     "We will now have an address by [Name], [Designation].\n"
     "I request [Name] to kindly come forward.\n[After speech] Thank you, [Name].",
     "ಈಗ [ಹೆಸರು], [ಹುದ್ದೆ] ಅವರಿಂದ ಭಾಷಣ ನಡೆಯಲಿದೆ.\n"
     "[ಹೆಸರು] ಅವರನ್ನು ಮುಂದೆ ಬರಲು ವಿನಂತಿಸುತ್ತೇವೆ."),
]

def get_script(item):
    t = item.lower()
    for keywords, eng, kan in SCRIPTS:
        if any(k in t for k in keywords):
            return eng, kan
    eng = f"We will now have — {item}.\nI request [Name / Designation] to kindly come forward.\n[MC Note: Add specific announcement text]"
    kan = f"ಈಗ — {item}.\n[ಹೆಸರು] ಅವರನ್ನು ಮುಂದೆ ಬರಲು ವಿನಂತಿಸುತ್ತೇವೆ."
    return eng, kan

def is_address(item):
    return any(k in item.lower() for k in ["address","keynote","remarks","speech","perspective","introduction"])

def get_category(item):
    t = item.lower()
    if any(k in t for k in ["address","keynote","remarks","speech","perspective","context","introduction"]): return "Address / Speech"
    if any(k in t for k in ["felicitat","release","souvenir","policy","publication"]): return "Felicitation / Release"
    if any(k in t for k in ["cultural","dance","music","performance"]): return "Cultural Programme"
    if any(k in t for k in ["naada","anthem","lamp","lighting","national","inaugur"]): return "Ceremonial"
    if any(k in t for k in ["welcome","dignitar","vote","thanks","break","tea","interval"]): return "Welcome / Housekeeping"
    return "General"

CAT_COLORS = {
    "Address / Speech":       "#D6E4F0",
    "Felicitation / Release": "#FDEBD0",
    "Cultural Programme":     "#F4D03F",
    "Ceremonial":             "#E8DAEF",
    "Welcome / Housekeeping": "#D5F5E3",
    "General":                "#FFFFFF",
}

# ── Time helpers ──────────────────────────────────────────────────────────────
def parse_time(t):
    for fmt in ["%I:%M %p","%I:%M%p","%H:%M","%I %p"]:
        try: return datetime.strptime(t.strip().upper(), fmt)
        except: pass
    return None

def fmt_time(dt): return dt.strftime("%I:%M %p").lstrip("0")
def fmt_slot(s,e): return f"{fmt_time(s)} – {fmt_time(e)}"

# ── Safe filename ─────────────────────────────────────────────────────────────
def safe_filename(name):
    import re
    return re.sub(r'[\\/*?:"<>|]','', name).strip() or "Programme"

# ── Word doc export ───────────────────────────────────────────────────────────
def build_word(event_name, event_date, venue, rows, logo_bytes=None):
    """
    Generates a clean Word doc matching the BTS 2023 MC Copy format:
    - Logo centred at top (if provided)
    - Title in maroon, centred
    - Date/venue line in grey, centred
    - Clean 3-col table: Timings | : | Programme Details
    - No colour fills — pure black & white
    - MC Script on page 2
    Uses Node.js docx library for reliable output.
    """
    import subprocess, json, base64, tempfile, os

    # Build script data
    script_rows = []
    for row in rows:
        eng, kan = get_script(row['item'])
        script_rows.append({
            'slot':     row['slot'],
            'item':     row['item'],
            'eng_lines': [l for l in eng.split('\n') if l.strip() or True],
            'kan_lines': [l for l in kan.split('\n') if l.strip() or True],
        })

    payload = {
        'event_name': event_name or 'INAUGURAL',
        'event_date': event_date or '',
        'venue':      venue or '',
        'start_time': rows[0]['start_str'] if rows else '',
        'total_mins': sum(r['duration'] for r in rows),
        'end_time':   rows[-1]['end_str'] if rows else '',
        'rows':       script_rows,
        'logo_b64':   base64.b64encode(logo_bytes).decode() if logo_bytes else '',
    }

    js_script = r"""
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
        PageBreak, ImageRun } = require('docx');
const fs = require('fs');

const data = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));

const thinBorder = { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" };
const noBorder   = { style: BorderStyle.NIL, size: 0, color: "FFFFFF" };
const allBorders = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder, insideH: thinBorder, insideV: thinBorder };
const noBorders  = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
const hdrBorder  = { style: BorderStyle.SINGLE, size: 6, color: "999999" };
const hdrBorders = { top: hdrBorder, bottom: hdrBorder, left: hdrBorder, right: hdrBorder, insideH: hdrBorder, insideV: hdrBorder };

const children = [];

// ── Logo (centred, above title) ──
if (data.logo_b64) {
    const logoBuffer = Buffer.from(data.logo_b64, 'base64');
    children.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 120 },
        children: [new ImageRun({
            data: logoBuffer,
            transformation: { width: 180, height: 70 },
            type: 'png',
        })]
    }));
}

// ── Title ──
const titleLine1 = `MINUTE-TO-MINUTE PROGRAMME FOR INAUGURAL CEREMONY OF`;
const titleLine2 = data.event_name.toUpperCase();
children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 40 },
    children: [new TextRun({ text: titleLine1, bold: true, font: "Arial", size: 22, color: "7B1B1B" })]
}));
children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 40 },
    children: [
        new TextRun({ text: titleLine2, bold: true, font: "Arial", size: 22, color: "7B1B1B" }),
        new TextRun({ text: "  (MC Copy)", bold: true, font: "Arial", size: 22, color: "C9A84C" }),
    ]
}));

// ── Date / Venue line ──
const detailParts = [];
if (data.event_date) detailParts.push(data.event_date);
if (data.start_time) detailParts.push(data.start_time);
if (data.venue)      detailParts.push(data.venue);
children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 240 },
    children: [new TextRun({ text: detailParts.join(' | '), font: "Arial", size: 18, color: "555555" })]
}));

// ── Programme table ──
const headerRow = new TableRow({
    tableHeader: true,
    children: [
        new TableCell({
            width: { size: 2400, type: WidthType.DXA },
            borders: hdrBorders,
            shading: { fill: "F5F5F5", type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 120, right: 80 },
            children: [new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: "Timings", bold: true, font: "Arial", size: 20 })]
            })]
        }),
        new TableCell({
            width: { size: 360, type: WidthType.DXA },
            borders: hdrBorders,
            shading: { fill: "F5F5F5", type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 60, right: 60 },
            children: [new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: " ", font: "Arial", size: 20 })]
            })]
        }),
        new TableCell({
            width: { size: 6400, type: WidthType.DXA },
            borders: hdrBorders,
            shading: { fill: "F5F5F5", type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 120, right: 80 },
            children: [new Paragraph({
                children: [new TextRun({ text: "Programme Details", bold: true, font: "Arial", size: 20 })]
            })]
        }),
    ]
});

const dataRows = data.rows.map((row, i) => new TableRow({
    children: [
        new TableCell({
            width: { size: 2400, type: WidthType.DXA },
            borders: allBorders,
            verticalAlign: VerticalAlign.TOP,
            margins: { top: 80, bottom: 80, left: 120, right: 80 },
            children: [new Paragraph({
                alignment: AlignmentType.LEFT,
                spacing: { before: 0, after: 0 },
                children: [new TextRun({ text: row.slot, font: "Arial", size: 18 })]
            })]
        }),
        new TableCell({
            width: { size: 360, type: WidthType.DXA },
            borders: allBorders,
            verticalAlign: VerticalAlign.TOP,
            margins: { top: 80, bottom: 80, left: 60, right: 60 },
            children: [new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 0 },
                children: [new TextRun({ text: ":", font: "Arial", size: 18 })]
            })]
        }),
        new TableCell({
            width: { size: 6400, type: WidthType.DXA },
            borders: allBorders,
            verticalAlign: VerticalAlign.TOP,
            margins: { top: 80, bottom: 80, left: 120, right: 80 },
            children: [new Paragraph({
                spacing: { before: 0, after: 0 },
                children: [new TextRun({ text: row.item, font: "Arial", size: 18 })]
            })]
        }),
    ]
}));

children.push(new Table({
    width: { size: 9160, type: WidthType.DXA },
    columnWidths: [2400, 360, 6400],
    rows: [headerRow, ...dataRows],
}));

// ── Total line ──
children.push(new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { before: 100, after: 0 },
    children: [new TextRun({
        text: `Total: ${data.total_mins} mins (${Math.floor(data.total_mins/60)}h ${data.total_mins%60}m)  |  Programme ends at ${data.end_time}`,
        font: "Arial", size: 16, italics: true, color: "555555"
    })]
}));

// ── Page break → MC Script ──
children.push(new Paragraph({ children: [new PageBreak()] }));

// ── MC Script heading ──
children.push(new Paragraph({
    spacing: { before: 0, after: 160 },
    children: [new TextRun({ text: "MC SCRIPT — Bilingual (English + ಕನ್ನಡ)", bold: true, font: "Arial", size: 24, color: "7B1B1B" })]
}));
children.push(new Paragraph({
    spacing: { before: 0, after: 300 },
    children: [new TextRun({
        text: "Replace all text in [brackets] with actual names/designations before the event.",
        font: "Arial", size: 18, italics: true, color: "888888"
    })]
}));

// ── MC Script entries ──
data.rows.forEach((row, i) => {
    children.push(new Paragraph({
        spacing: { before: 240, after: 60 },
        children: [
            new TextRun({ text: `${i+1}.  ${row.item}`, bold: true, font: "Arial", size: 20, color: "7B1B1B" }),
            new TextRun({ text: `  —  ${row.slot}`, font: "Arial", size: 18, color: "555555" }),
        ]
    }));

    const scriptTable = new Table({
        width: { size: 9160, type: WidthType.DXA },
        columnWidths: [4580, 4580],
        rows: [new TableRow({
            children: [
                new TableCell({
                    width: { size: 4580, type: WidthType.DXA },
                    borders: allBorders,
                    shading: { fill: "F5F9FF", type: ShadingType.CLEAR },
                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                    children: [
                        new Paragraph({ spacing: { before:0, after:60 },
                            children:[new TextRun({ text:"📢 English", bold:true, font:"Arial", size:18, color:"1A5276" })] }),
                        ...row.eng_lines.map(line => new Paragraph({
                            spacing: { before:0, after:40 },
                            children:[new TextRun({ text: line, font:"Arial", size:17 })]
                        }))
                    ]
                }),
                new TableCell({
                    width: { size: 4580, type: WidthType.DXA },
                    borders: allBorders,
                    shading: { fill: "FFFCF5", type: ShadingType.CLEAR },
                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                    children: [
                        new Paragraph({ spacing: { before:0, after:60 },
                            children:[new TextRun({ text:"📢 ಕನ್ನಡ", bold:true, font:"Nirmala UI", size:18, color:"7B5200" })] }),
                        ...row.kan_lines.map(line => new Paragraph({
                            spacing: { before:0, after:40 },
                            children:[new TextRun({ text: line, font:"Nirmala UI", size:17 })]
                        }))
                    ]
                }),
            ]
        })]
    });
    children.push(scriptTable);
});

const doc = new Document({
    sections: [{
        properties: {
            page: {
                size: { width: 11906, height: 16838 },
                margin: { top: 720, right: 720, bottom: 720, left: 720 }
            }
        },
        children
    }]
});

Packer.toBuffer(doc).then(buf => {
    fs.writeFileSync(process.argv[3], buf);
});
"""

    try:
        # Write payload to temp file
        tmp_dir  = tempfile.mkdtemp()
        data_path = os.path.join(tmp_dir, 'data.json')
        out_path  = os.path.join(tmp_dir, 'output.docx')
        js_path   = os.path.join(tmp_dir, 'build.js')

        with open(data_path, 'w', encoding='utf-8') as f:
            json.dump(payload, f, ensure_ascii=False)
        with open(js_path, 'w', encoding='utf-8') as f:
            f.write(js_script)

        result = subprocess.run(
            ['node', js_path, data_path, out_path],
            capture_output=True, text=True, timeout=30
        )
        if result.returncode != 0:
            return None, f"JS error: {result.stderr[:300]}"

        with open(out_path, 'rb') as f:
            buf = io.BytesIO(f.read())
        buf.seek(0)
        return buf, None

    except Exception as e:
        return None, str(e)


# ── Excel export ──────────────────────────────────────────────────────────────
def build_excel(event_name, event_date, venue, rows, logo_bytes=None):
    MAROON="7B1B1B"; GOLD="C9A84C"; WHITE="FFFFFF"; DARK="2C2C2C"
    CREAM="FFF8EE"; LT_GOLD="F5E6C8"; LT_MAROON="F2DADA"; GREY="F7F7F7"

    def sd(s="thin",c=GOLD): return Side(style=s,color=c)
    def bdr(s="thin"): b=sd(s); return Border(left=b,right=b,top=b,bottom=b)
    def fill(h): return PatternFill("solid",fgColor=h)
    def fnt(bold=False,color=DARK,size=10,italic=False):
        return Font(name="Arial",bold=bold,color=color,size=size,italic=italic)
    def aln(h="center",v="center",wrap=False):
        return Alignment(horizontal=h,vertical=v,wrap_text=wrap)
    def mbdr():
        m=Side(style="medium",color=MAROON)
        return Border(left=m,right=m,top=m,bottom=m)

    wb = Workbook()

    # ── Sheet 1: Programme Planner ──
    ws1=wb.active; ws1.title="Programme Planner"
    for col,w in {"A":22,"B":52,"C":18}.items():
        ws1.column_dimensions[col].width=w

    # ── Header: Left = event info, Right = logo ──
    # Row 1: Event Name (A1) | Logo (C1 spanning down)
    ws1.row_dimensions[1].height = 28
    ws1.row_dimensions[2].height = 22
    ws1.row_dimensions[3].height = 22
    ws1.row_dimensions[4].height = 22

    # Event name — left side
    c=ws1["A1"]
    c.value = f"M2M PROGRAMME FOR {(event_name or 'INAUGURAL').upper()}"
    c.font=Font(name="Arial",bold=True,color=WHITE,size=14)
    c.fill=fill(MAROON); c.alignment=aln(h="left")
    ws1.merge_cells("A1:B1")

    # Event details rows
    details_rows = [
        ("Event :", event_name or "—"),
        ("Date  :", event_date or "—"),
        ("Venue :", venue or "—"),
        ("Start :", rows[0]["start_str"] if rows else "—"),
    ]
    for di, (lbl, val) in enumerate(details_rows):
        r = 2 + di
        ws1.row_dimensions[r].height = 18
        cl = ws1.cell(row=r, column=1)
        cl.value = lbl
        cl.font = Font(name="Arial",bold=True,color=MAROON,size=9)
        cl.fill = fill(LT_GOLD); cl.alignment = aln(h="right")
        cv = ws1.cell(row=r, column=2)
        cv.value = val
        cv.font = Font(name="Arial",color=DARK,size=9)
        cv.fill = fill(WHITE); cv.border = bdr()
        cv.alignment = aln(h="left")

    # Logo — right side (col C, rows 1-4)
    if logo_bytes:
        from openpyxl.drawing.image import Image as XLImage
        try:
            img = XLImage(io.BytesIO(logo_bytes))
            img.height = 88; img.width = 160; img.anchor = "C1"
            ws1.add_image(img)
            ws1.column_dimensions["C"].width = 22
        except Exception:
            pass

    # (event details now in header rows above)

    ws1.row_dimensions[6].height=24
    for ci,hdr in enumerate(["Time Slot","Programme Item / Activity"],1):
        c=ws1.cell(row=6,column=ci)
        c.value=hdr; c.font=Font(name="Arial",bold=True,color=WHITE,size=10)
        c.fill=fill(MAROON); c.alignment=aln(wrap=True); c.border=bdr()

    from openpyxl.styles import PatternFill as PF
    for i,row in enumerate(rows):
        r=7+i; ws1.row_dimensions[r].height=22
        item=row.get('item',''); cat=get_category(item)
        cat_hex=CAT_COLORS.get(cat,"#FFFFFF").lstrip("#")
        alt=CREAM if i%2==0 else "FFFFFF"; bg=cat_hex if cat!="General" else alt
        for ci,val,bg2,fs in [
            (1,row.get('slot',''),LT_GOLD,Font(name="Arial",bold=True,color=MAROON,size=10)),
            (2,item,bg,fnt(color=DARK,bold=is_address(item)))]:
            c=ws1.cell(row=r,column=ci)
            c.value=val; c.fill=PF("solid",fgColor=bg2); c.font=fs
            c.alignment=aln(h="center" if ci==1 else "left",wrap=(ci==2))
            c.border=bdr()

    tr=7+len(rows); ws1.row_dimensions[tr].height=22
    total_mins=sum(r.get('duration',0) for r in rows if isinstance(r.get('duration'),int))
    c=ws1[f"A{tr}"]
    c.value=f"Total: {total_mins} mins ({total_mins//60}h {total_mins%60}m)  |  Ends: {rows[-1]['end_str'] if rows else ''}"
    c.font=Font(name="Arial",bold=True,color=GOLD,size=10)
    c.fill=fill(DARK); c.alignment=aln(h="left"); c.border=mbdr()
    ws1.merge_cells(f"A{tr}:B{tr}")
    ws1.freeze_panes="A7"

    # ── Sheet 2: Print View ──
    ws2=wb.create_sheet("Print View")
    ws2.column_dimensions["A"].width=22
    ws2.column_dimensions["B"].width=52
    ws2.column_dimensions["C"].width=22

    # Header: event info left, logo right
    ws2.row_dimensions[1].height=28
    c=ws2["A1"]
    c.value=f"M2M PROGRAMME FOR {(event_name or 'INAUGURAL').upper()}"
    c.font=Font(name="Arial",bold=True,color=WHITE,size=14)
    c.fill=fill(MAROON); c.alignment=aln(h="left")
    ws2.merge_cells("A1:B1")

    for di,(lbl,val) in enumerate([
        ("Event :", event_name or "—"),("Date  :", event_date or "—"),
        ("Venue :", venue or "—"),("Start :", rows[0]["start_str"] if rows else "—")]):
        r=2+di; ws2.row_dimensions[r].height=18
        cl=ws2.cell(row=r,column=1); cl.value=lbl
        cl.font=Font(name="Arial",bold=True,color=MAROON,size=9)
        cl.fill=fill(LT_GOLD); cl.alignment=aln(h="right")
        cv=ws2.cell(row=r,column=2); cv.value=val
        cv.font=Font(name="Arial",color=DARK,size=9)
        cv.fill=fill(WHITE); cv.border=bdr(); cv.alignment=aln(h="left")

    if logo_bytes:
        from openpyxl.drawing.image import Image as XLImage
        try:
            img2=XLImage(io.BytesIO(logo_bytes))
            img2.height=88; img2.width=160; img2.anchor="C1"
            ws2.add_image(img2)
        except Exception: pass

    ws2.row_dimensions[6].height=22
    for ci,hdr in enumerate(["Time Slot","Programme Item / Activity"],1):
        c=ws2.cell(row=6,column=ci)
        c.value=hdr; c.font=Font(name="Arial",bold=True,color=WHITE,size=11)
        c.fill=fill(MAROON); c.alignment=aln(); c.border=bdr()
    for i,row in enumerate(rows):
        r=7+i; ws2.row_dimensions[r].height=22
        item=row.get('item',''); cat=get_category(item)
        cat_hex=CAT_COLORS.get(cat,"#FFFFFF").lstrip("#")
        bg=cat_hex if cat!="General" else (CREAM if i%2==0 else "FFFFFF")
        for ci,val in[(1,row.get('slot','')), (2,item)]:
            c=ws2.cell(row=r,column=ci)
            c.value=val; c.fill=PF("solid",fgColor=bg)
            c.font=fnt(bold=(ci==2 and is_address(item)),
                       color="1A5276" if is_address(item) else DARK)
            c.alignment=aln(h="center" if ci==1 else "left",wrap=(ci==2))
            c.border=bdr()
    ws2.freeze_panes="A7"

    # ── Sheet 3: MC Script ──
    ws3=wb.create_sheet("MC Script")
    for col,w in {"A":22,"B":30,"C":50,"D":46,"E":20}.items():
        ws3.column_dimensions[col].width=w
    # MC header: info left, logo right
    ws3.row_dimensions[1].height=28
    c=ws3["A1"]
    c.value="🎤  MC SCRIPT — Bilingual (English + ಕನ್ನಡ)"
    c.font=Font(name="Arial",bold=True,color=WHITE,size=13)
    c.fill=fill(MAROON); c.alignment=aln(h="left")
    ws3.merge_cells("A1:D1")

    for di,(lbl,val) in enumerate([
        ("Event :", event_name or "—"),("Date  :", event_date or "—"),
        ("Venue :", venue or "—"),("Note  :", "Replace [brackets] with actual names")]):
        r=2+di; ws3.row_dimensions[r].height=16
        cl=ws3.cell(row=r,column=1); cl.value=lbl
        cl.font=Font(name="Arial",bold=True,color=MAROON,size=9)
        cl.fill=fill(LT_GOLD); cl.alignment=aln(h="right")
        cv=ws3.cell(row=r,column=2); cv.value=val
        cv.font=Font(name="Arial",color=DARK,size=9)
        cv.fill=fill(WHITE); cv.border=bdr(); cv.alignment=aln(h="left")
        ws3.merge_cells(f"B{r}:D{r}")

    if logo_bytes:
        from openpyxl.drawing.image import Image as XLImage
        try:
            img3=XLImage(io.BytesIO(logo_bytes))
            img3.height=88; img3.width=160; img3.anchor="E1"
            ws3.add_image(img3)
        except Exception: pass

    ws3.row_dimensions[6].height=26
    for ci,hdr in enumerate(["Time Slot","Programme Item","📢 English","📢 ಕನ್ನಡ"],1):
        c=ws3.cell(row=6,column=ci)
        c.value=hdr; c.font=Font(name="Arial",bold=True,color=WHITE,size=10)
        c.fill=fill(MAROON); c.alignment=aln(wrap=True); c.border=bdr()
    for i,row in enumerate(rows):
        r=7+i; ws3.row_dimensions[r].height=90
        item=row.get('item',''); cat=get_category(item)
        cat_hex=CAT_COLORS.get(cat,"#FFFFFF").lstrip("#")
        eng,kan=get_script(item)
        for ci,val,bg2,fs in [
            (1,row.get('slot',''),LT_GOLD,fnt(bold=True,color=MAROON,size=9)),
            (2,item,cat_hex,fnt(bold=is_address(item),color=DARK)),
            (3,eng,"F0F7FF" if i%2==0 else "FFFFFF",Font(name="Arial",color="1A1A2E",size=9)),
            (4,kan,"FFF8F0" if i%2==0 else "FFFFFF",Font(name="Nirmala UI",color="1A1A2E",size=9))]:
            c=ws3.cell(row=r,column=ci)
            c.value=val; c.fill=PF("solid",fgColor=bg2); c.font=fs
            c.alignment=Alignment(horizontal="center" if ci==1 else "left",
                                  vertical="top" if ci in[3,4] else "center",wrap_text=True)
            c.border=bdr()
    ws3.freeze_panes="A7"
    ws3.page_setup.orientation="landscape"

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ═══════════════════════════════════════════════════════════════════════════
# MAIN APP UI
# ═══════════════════════════════════════════════════════════════════════════

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Event Details")
    event_name = st.text_input("Event Name", placeholder="e.g. BTS 2025 Inauguration")
    event_date = st.text_input("Event Date", placeholder="e.g. 15-Aug-2025")
    venue      = st.text_input("Venue", placeholder="e.g. Taj Vivanta, Bengaluru")
    start_time = st.text_input("Start Time", value="10:00 AM")

    st.markdown("---")
    st.markdown("### 🖼️ Event Logo")
    logo_file = st.file_uploader(
        "Upload logo (PNG or JPG)",
        type=["png","jpg","jpeg"],
        help="Logo will appear in downloaded Excel and Word files"
    )
    logo_bytes = logo_file.read() if logo_file else None

    st.markdown("---")
    st.markdown("### ℹ️ How to use")
    st.markdown("""
1. Fill event details above
2. Upload logo (optional)
3. Add programme items + durations
4. Download Excel or Word
""")
    st.markdown("---")
    st.markdown("**Category colours:**")
    for cat, color in CAT_COLORS.items():
        if cat != "General":
            st.markdown(
                f"<span style='background:{color};padding:2px 8px;"
                f"border-radius:4px;font-size:0.78rem;color:#2C2C2C'>{cat}</span><br>",
                unsafe_allow_html=True)

# ── Title banner (plain — logo goes into downloaded files only) ──────────────
st.markdown("""
<div class="title-banner">
    <div class="title-banner-text">
        <h1>📋 M2M Programme Planner</h1>
        <p>Minute-to-Minute Schedule · Bilingual MC Script · Excel & Word Download</p>
    </div>
</div>
""", unsafe_allow_html=True)

if logo_bytes:
    st.info("✅ Logo uploaded — it will appear in your downloaded Excel and Word files.")

# ── Event info panel ──────────────────────────────────────────────────────────
if event_name or event_date or venue:
    info_parts = []
    if event_name: info_parts.append(f"<b>{event_name}</b>")
    if event_date: info_parts.append(f"📅 {event_date}")
    if venue:      info_parts.append(f"📍 {venue}")
    if start_time: info_parts.append(f"🕐 {start_time}")
    st.markdown(
        f'<div class="event-info-panel">'
        f'{"  &nbsp;|&nbsp;  ".join(info_parts)}</div>',
        unsafe_allow_html=True)

# ── Parse start time ──────────────────────────────────────────────────────────
start_dt = parse_time(start_time) if start_time else parse_time("10:00 AM")
if not start_dt:
    st.error(f"⚠️ Can't parse '{start_time}'. Use format like '10:00 AM' or '09:30 AM'")
    st.stop()

# ── Programme table ───────────────────────────────────────────────────────────
st.markdown('<div class="section-header">📝 Programme Items</div>',
            unsafe_allow_html=True)

DEFAULT_ITEMS = [
    {"item":"Welcome of Dignitaries to the Dais","duration":5,"remarks":"MC announces names"},
    {"item":"Naada Geethe (State Anthem)","duration":3,"remarks":"All stand"},
    {"item":"Lighting of the Lamp","duration":5,"remarks":"Chief Guest & dignitaries"},
    {"item":"Welcome Address by Host","duration":7,"remarks":""},
    {"item":"Release of Souvenir / Publication","duration":3,"remarks":""},
    {"item":"Keynote Address","duration":15,"remarks":"Chief Guest"},
    {"item":"Felicitation of Dignitaries","duration":10,"remarks":""},
    {"item":"Cultural Programme","duration":10,"remarks":""},
    {"item":"Vote of Thanks","duration":5,"remarks":""},
    {"item":"National Anthem","duration":2,"remarks":"All stand — Jai Hind"},
    {"item":"","duration":None,"remarks":""},
    {"item":"","duration":None,"remarks":""},
    {"item":"","duration":None,"remarks":""},
    {"item":"","duration":None,"remarks":""},
    {"item":"","duration":None,"remarks":""},
]

edited_df = st.data_editor(
    pd.DataFrame(DEFAULT_ITEMS),
    column_config={
        "item":     st.column_config.TextColumn("Programme Item / Activity", width="large"),
        "duration": st.column_config.NumberColumn("Duration (mins)", min_value=1, max_value=180, step=1),
        "remarks":  st.column_config.TextColumn("Remarks / Speaker", width="medium"),
    },
    num_rows="dynamic",
    use_container_width=True,
    height=430,
    key="programme_table"
)

# ── Calculate time slots ──────────────────────────────────────────────────────
rows = []
current = start_dt
for _, r in edited_df.iterrows():
    item     = str(r.get('item','')).strip() if r.get('item') else ''
    duration = r.get('duration')
    remarks  = str(r.get('remarks','')).strip() if r.get('remarks') else ''
    if not item or not duration: continue
    try: duration = int(duration)
    except: continue
    end = current + timedelta(minutes=duration)
    rows.append({
        'item': item, 'duration': duration,
        'start_str': fmt_time(current), 'end_str': fmt_time(end),
        'slot': fmt_slot(current, end), 'remarks': remarks,
    })
    current = end

# ── Output tabs + downloads ───────────────────────────────────────────────────
if rows:
    total_mins = sum(r['duration'] for r in rows)

    # Summary cards
    c1, c2, c3 = st.columns(3)
    for col, val, label in [
        (c1, str(len(rows)),        "Programme Items"),
        (c2, f"{total_mins//60}h {total_mins%60}m", "Total Duration"),
        (c3, rows[-1]['end_str'],   "Programme Ends"),
    ]:
        with col:
            st.markdown(
                f'<div class="summary-card">'
                f'<div style="font-size:1.6rem">{val}</div>'
                f'<div style="font-size:0.8rem;color:#F5E6C8">{label}</div>'
                f'</div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs(["📋 Schedule", "🎤 MC Script", "🖨️ Print View"])

    with tab1:
        st.markdown('<div class="section-header">📋 Minute-to-Minute Schedule</div>',
                    unsafe_allow_html=True)
        for i, row in enumerate(rows):
            cat   = get_category(row['item'])
            color = CAT_COLORS.get(cat,"#FFFFFF")
            addr  = is_address(row['item'])
            st.markdown(f"""
            <div style="display:flex;align-items:center;padding:8px 14px;
                        background:{color};border-radius:6px;margin:4px 0;
                        border-left:4px solid {'#1A5276' if addr else '#C9A84C'}">
                <span class="time-pill" style="min-width:150px;text-align:center">{row['slot']}</span>
                <span style="flex:1;margin-left:14px;font-weight:{'bold' if addr else 'normal'};
                             color:{'#1A5276' if addr else '#2C2C2C'}">{row['item']}</span>
            </div>""", unsafe_allow_html=True)

    with tab2:
        st.markdown('<div class="section-header">🎤 Bilingual MC Script</div>',
                    unsafe_allow_html=True)
        st.info("💡 Replace **[bracket]** text with actual names/designations before the event")
        for i, row in enumerate(rows):
            eng, kan = get_script(row['item'])
            with st.expander(f"**{i+1}. {row['item']}** — {row['slot']}", expanded=(i<2)):
                ce, ck = st.columns(2)
                with ce:
                    st.markdown("**📢 English**")
                    st.markdown(f'<div class="script-box">{eng.replace(chr(10),"<br>")}</div>',
                                unsafe_allow_html=True)
                with ck:
                    st.markdown("**📢 ಕನ್ನಡ**")
                    st.markdown(f'<div class="script-box-kn">{kan.replace(chr(10),"<br>")}</div>',
                                unsafe_allow_html=True)

    with tab3:
        st.markdown('<div class="section-header">🖨️ Print View</div>', unsafe_allow_html=True)
        if event_name:
            st.markdown(f"### M2M Programme for {event_name}")
        info = []
        if event_date: info.append(f"📅 {event_date}")
        if venue:      info.append(f"📍 {venue}")
        if info: st.markdown("  |  ".join(info))
        st.divider()
        st.dataframe(
            pd.DataFrame([{
                "Time Slot": r['slot'],
                "Programme Item": r['item'],
            } for r in rows]),
            use_container_width=True, hide_index=True, height=460)

    # ── Downloads ─────────────────────────────────────────────────────────────
    st.divider()
    st.markdown("### 📥 Download")

    fname_base = f"M2M Programme for {safe_filename(event_name)}" if event_name else "M2M Programme"

    dcol1, dcol2 = st.columns(2)

    with dcol1:
        excel_buf = build_excel(event_name, event_date, venue, rows, logo_bytes)
        st.download_button(
            label="⬇️ Download Excel (.xlsx)",
            data=excel_buf,
            file_name=f"{fname_base}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.caption("3 sheets: Programme Planner · Print View · MC Script")

    with dcol2:
        word_buf, word_err = build_word(event_name, event_date, venue, rows, logo_bytes)
        if word_buf:
            st.markdown('<div class="dl-word">', unsafe_allow_html=True)
            st.download_button(
                label="⬇️ Download Word (.docx)",
                data=word_buf,
                file_name=f"{fname_base}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
            st.markdown('</div>', unsafe_allow_html=True)
            st.caption("Includes Programme table + bilingual MC Script")
        else:
            st.warning(f"Word export unavailable: run `pip install python-docx` to enable")

else:
    st.info("👆 Add programme items and durations above to generate your schedule.")

# ── Footer ─────────────────────────────────────────────────────────────────────
st.divider()
st.markdown("""
<div style="text-align:center;color:#888;font-size:0.8rem">
    M2M Programme Planner · Built for Event Management Teams
</div>""", unsafe_allow_html=True)
