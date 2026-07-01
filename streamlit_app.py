"""
Award Slide Generator - Streamlit web app
===========================================
A browser-based front end for generate_award_slides.py. Lets anyone on the
team upload the Excel results sheet + the PowerPoint certificate template
and download one combined deck (Nominee slide immediately followed by its
Winner slide, for every award), with no installation needed on their side.

The actual generation logic lives entirely in generate_award_slides.py and
is reused unchanged here -- this file is only the upload/download UI.

Run locally with:
    streamlit run streamlit_app.py

Deploy on Streamlit Community Cloud by pointing it at this file as the
app's entry point, with requirements.txt alongside it.
"""
import io
import os
import tempfile
import contextlib
import traceback

import streamlit as st

from generate_award_slides import parse_excel, build_combined_deck


st.set_page_config(
    page_title="Award Slide Generator",
    page_icon="🏆",
    layout="centered",
)

st.title("🏆 Award Slide Generator")
st.write(
    "Upload your results spreadsheet and your PowerPoint certificate "
    "template. You'll get back one ready-to-present deck: for every award, "
    "a Nominee slide immediately followed by its Winner slide."
)

with st.sidebar:
    st.header("📋 Template naming convention")
    st.caption(
        "Read this before uploading — it explains exactly what your "
        "PowerPoint template and Excel file need to contain so the "
        "generator can recognize them correctly."
    )

    st.subheader("PowerPoint template")
    st.markdown(
        """
A slide's role is decided by which **placeholder text box** it contains —
not by its position or slide number. Every token below is **optional**
except that a slide needs at least `<<NOMINEES>>` or `<<WINNER>>` to be
used at all. Each box keeps its own existing position, size, and
formatting — only its text changes.
        """
    )
    st.markdown(
        """
| Token | Meaning |
|---|---|
| `<<NOMINEES>>` | Marks this slide as the **Nominee** stencil. The box itself is filled with the full nominee list, one per line |
| `<<WINNER>>` | Marks this slide as the **Winner** stencil. The box is filled with the winning company's name |
| `<<AWARD CATEGORY>>` | *(Optional)* Where the award category title is inserted |
| `<<ZONE>>` | *(Optional)* Where the zone/region name is inserted |
| `<<nominees-word>>` | *(Optional)* Replaced with the literal word **NOMINEES** |
| `<<winner-word>>` | *(Optional)* Replaced with the literal word **WINNER** |
        """
    )
    st.error(
        "**Matching is literal and exact.** The box's text must be exactly "
        "one of the tokens above (case doesn't matter, but wording does). "
        "`<<winner placeholder>>` or `<<Winner Name>>` will **NOT** match "
        "`<<WINNER>>` — type the token exactly as shown.",
        icon="⚠️",
    )
    st.markdown(
        """
**Mix and match freely** depending on what each event needs:
- **Nominee slide + Winner slide** (two stencils) — the classic case: a
  slide with `<<NOMINEES>>`, immediately followed in the output by a slide
  with `<<WINNER>>`, for every award.
- **Winner-only** — for events with no nominee reveal, just include a
  single slide with `<<WINNER>>` (and skip `<<NOMINEES>>` entirely). Only
  Winner slides are generated.
- **Nominee-only** works the same way, the other direction.
- Add `<<AWARD CATEGORY>>` / `<<ZONE>>` / `<<nominees-word>>` /
  `<<winner-word>>` on either stencil only where you actually want that
  content to appear — leave them off entirely if you don't need them.
- Any other slide with none of these tokens (e.g. a title or thank-you
  slide) is left as-is and copied through once at the end of the deck.
        """
    )

    st.subheader("Excel spreadsheet")
    st.markdown(
        """
One sheet. Column **headers can say anything** — columns are recognized by
the *shape* of their data, not their names:
- **Category column** — award category text, filled only on the first row
  of each award's block.
- **Nominee column** — company/nominee names, filled on every row.
- **Result column** — a small repeating set of labels, e.g. `Winner`,
  `1st Runnerup`, `2nd Runnerup` (or similar wording).
- *(Optional)* **Zone column** — a small repeating set of region values,
  e.g. `North` / `South` / `East` / `West`.

A blank row separates one zone's group of nominees from the next zone's
group within the same category.
        """
    )

st.divider()

col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader(
        "Results spreadsheet", type=["xlsx", "xlsm"], key="excel_upload"
    )
with col2:
    template_file = st.file_uploader(
        "PowerPoint template", type=["pptx"], key="template_upload"
    )

generate_clicked = st.button(
    "Generate deck", type="primary",
    disabled=not (excel_file and template_file),
)

if not (excel_file and template_file):
    st.caption("Upload both files to enable the Generate button.")

if "deck_bytes" not in st.session_state:
    st.session_state.deck_bytes = None
    st.session_state.deck_log = ""
    st.session_state.deck_group_count = 0
    st.session_state.deck_error = None

if generate_clicked and excel_file and template_file:
    with tempfile.TemporaryDirectory() as tmp_dir:
        excel_path = os.path.join(tmp_dir, excel_file.name)
        template_path = os.path.join(tmp_dir, template_file.name)
        output_path = os.path.join(tmp_dir, "Award_Show_Deck.pptx")

        with open(excel_path, "wb") as f:
            f.write(excel_file.getbuffer())
        with open(template_path, "wb") as f:
            f.write(template_file.getbuffer())

        log_buffer = io.StringIO()
        # Reset any previous result before attempting a fresh generation.
        st.session_state.deck_bytes = None
        st.session_state.deck_error = None

        try:
            with st.spinner("Reading the spreadsheet and template, then building the deck..."):
                with contextlib.redirect_stdout(log_buffer):
                    groups = parse_excel(excel_path)
                    if not groups:
                        raise ValueError(
                            "No category/nominee rows were found. Check the Excel "
                            "file's layout."
                        )
                    build_combined_deck(template_path, groups, output_path)

            with open(output_path, "rb") as f:
                st.session_state.deck_bytes = f.read()
            st.session_state.deck_log = log_buffer.getvalue().strip()
            st.session_state.deck_group_count = len(groups)

        except ValueError as e:
            st.session_state.deck_log = log_buffer.getvalue().strip()
            st.session_state.deck_error = str(e)

        except Exception:
            # Anything unexpected (a bug, a malformed file python-pptx/
            # openpyxl can't parse, etc.) -- show the real traceback rather
            # than a vague "something went wrong" message, since whoever is
            # running this can usefully screenshot it for help even without
            # understanding it themselves.
            st.session_state.deck_error = "__unexpected__:" + traceback.format_exc()

# Render the most recent result (persists across reruns -- e.g. the rerun
# that happens when the download button itself is clicked -- instead of
# only showing up for the one render right after clicking "Generate deck").
if st.session_state.deck_error:
    if st.session_state.deck_log:
        with st.expander("Details", expanded=True):
            st.code(st.session_state.deck_log, language=None)
    if st.session_state.deck_error.startswith("__unexpected__:"):
        st.error("An unexpected error occurred.")
        st.code(st.session_state.deck_error.removeprefix("__unexpected__:"), language=None)
    else:
        st.error(f"Couldn't generate the deck: {st.session_state.deck_error}")

elif st.session_state.deck_bytes:
    if st.session_state.deck_log:
        with st.expander("Details", expanded=False):
            st.code(st.session_state.deck_log, language=None)

    st.success(f"Done — generated {st.session_state.deck_group_count} award categories/zones.")
    st.download_button(
        "⬇️ Download deck",
        data=st.session_state.deck_bytes,
        file_name="Award_Show_Deck.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        type="primary",
    )
