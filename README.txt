# Name Board Generator

Generates dignitary/speaker name boards (fold-over tent cards) from an Excel
list, as editable PowerPoint (.pptx) and optional PDF.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL Streamlit prints (usually http://localhost:8501).

PDF export requires LibreOffice (`soffice`) installed on the machine running
the app. If it isn't available, PPTX download still works — open it in
PowerPoint and use File > Export to get a PDF.

## Font — important, read before deploying

**No font file is included in this repo.** Alternate Gothic No.2 BT is a
licensed font; redistributing it in a public (or shared private)
repository would violate that license.

Instead, the app has a **font upload option in the sidebar** at runtime:
each user supplies their own already-licensed copy of the font, which is
used only for that session (held in temp storage, never committed or
persisted to the repo). This keeps the codebase itself license-clean.

Two things worth knowing:
1. **You (or whoever runs this) still need a valid license** to use
   Alternate Gothic No.2 BT — the upload option avoids the *app* being a
   distribution channel, but doesn't grant usage rights on its own.
2. **Font is not embedded inside the generated PPTX file.** The app sets the
   font *name* in the slide XML, but python-pptx does not embed the actual
   font binary into the .pptx. This means:
   - The in-app PDF preview/export will look correct (the uploaded font is
     installed server-side for rendering during that session).
   - If someone opens the downloaded .pptx on a different machine that
     doesn't have Alternate Gothic No.2 BT installed, PowerPoint will
     silently substitute a default font for the name/title text.
   - For guaranteed visual fidelity outside this app, install the actual
     font file on whichever machine will ultimately open/print the PPTX,
     or export to PDF (PDF rasterizes/embeds appearance, so it's safe to
     share as-is).

If no font is uploaded, the app falls back to a generic system font for
on-screen text-fitting calculations only; PowerPoint will still try to use
the configured font name and substitute if it isn't installed.

## Files

- `app.py` — Streamlit UI: upload Excel, preview, generate, download.
- `board_generator.py` — Core layout engine (font fitting, wrapping, slide
  building). No Streamlit dependency; can be reused or unit tested standalone.
- `fonts/` — Empty by design (see Fonts section above). Drop a locally
  licensed .ttf/.otf here for local runs, or use the in-app uploader.

## Excel format

The app focuses on just two things: each person's **Name**, and their
**Title/Company details** — and ignores everything else on the sheet.

- **Name column** — found by header word (`Name`, `Dignitary`, `Speaker`,
  `Guest`, etc.), or the first column if no header is recognized.
- **Title/Company** — every other populated column for that row is combined:
  - If there are genuinely separate **Title** and **Company** columns
    (recognized by header words like `Title`, `Designation`, `Position`),
    they're kept distinct and rendered as two stacked lines, exactly like
    the reference layout.
  - Otherwise (most real-world sheets — e.g. a single `Organization` column
    that already contains the full description), the column's text is kept
    intact as one combined line/block. **No guessing or comma-splitting is
    attempted** — splitting "Title, Company" apart is only reliable when the
    sheet already separates them into two real columns.
- **Ignored automatically**: serial number columns (`Sr No`, `#`, etc.),
  contact info (`Email`, `Phone`, etc.), fully blank rows, and section-divider
  rows sitting in the Name column with no other details (e.g. `DIGNITARIES`,
  `SPEAKERS`, `QUIZ`, `WORKSHOP CONDUCTORS`).
- A preview table is shown after upload (with an explanation of how it was
  interpreted) and is editable before generating, in case anything needs a
  manual tweak.

A starter template (clean `Name`/`Title`/`Company` columns) is available
from the sidebar, but it's not required — any reasonable layout works,
including a single sheet with mixed section headers and one combined
details column per person.

## Layout rules implemented

- Each dignitary gets one A4-landscape slide, split into two halves:
  top half rotated 180°, bottom half upright (fold-over tent card).
- Name: ALL CAPS, auto-shrinks to fit on one line.
- Title + Company: Title Case (minor words like "of", "and", "for" stay
  lowercase; apostrophe-s as in "Hon'ble" is not capitalized).
  - If both fit on their own line, they're stacked with a tight gap between
    them, and a larger gap above (between Name and Title).
  - If either is too long to fit on one line, Title and Company are merged
    into a single comma-separated block instead, which may wrap to 2 lines.

