# Award Slide Generator — Streamlit app

A browser-based version of the award slide generator: upload the results
spreadsheet and the PowerPoint template, click a button, download the
finished deck. No installation needed for whoever uses it — only for
whoever deploys it once.

## Folder contents

- `streamlit_app.py` — the web app (upload UI, generate button, download button).
- `generate_award_slides.py` — the actual generation engine (column
  detection, template scanning, slide building). Identical to the
  command-line version; the app just calls into it.
- `requirements.txt` — Python packages needed, for Streamlit Community
  Cloud (or any other host) to install automatically.

## Run it locally first (recommended before deploying)

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

This opens it at `http://localhost:8501` in your browser.

## Deploy on Streamlit Community Cloud

1. Push this folder to a GitHub repository (public or private — Community
   Cloud can use either, a private repo just needs your GitHub account
   connected).
2. Go to [share.streamlit.io](https://share.streamlit.io), sign in with
   GitHub, click **"New app"**.
3. Pick the repository and branch, and set **Main file path** to
   `streamlit_app.py` (adjust the path if this folder isn't the repo root).
4. Click **Deploy**. It installs `requirements.txt` automatically and
   gives you a shareable `https://<something>.streamlit.app` link — send
   that to anyone on the team who needs to generate slides; they just need
   a browser.

### Updating it later

Any time `generate_award_slides.py` changes (a bug fix, a new detection
rule, etc.), just push the updated file to the same GitHub repo — Community
Cloud redeploys automatically within a minute or so. No manual redeploy
step needed.

## What it does NOT change

The detection logic — how columns are recognized in the Excel regardless
of header names, the winner-picking rule — is identical to the
command-line `.bat` version.

## Template naming convention

A slide's role is decided by which **placeholder text box** it contains.
Every token is optional except that a slide needs at least one of
`<<NOMINEES>>` / `<<WINNER>>` to be used as a stencil at all — mix and
match per event:

| Token | Meaning |
|---|---|
| `<<NOMINEES>>` | Marks this slide as the Nominee stencil; the box is filled with the full nominee list, one per line |
| `<<WINNER>>` | Marks this slide as the Winner stencil; the box is filled with the winning company's name |
| `<<AWARD CATEGORY>>` | *(Optional)* Award category title |
| `<<ZONE>>` | *(Optional)* Zone/region name, for categories split by region |
| `<<nominees-word>>` | *(Optional)* Replaced with the literal word **NOMINEES** |
| `<<winner-word>>` | *(Optional)* Replaced with the literal word **WINNER** |

Matching is **literal and exact** (case-insensitive only) — a box must
contain exactly one of these tokens, not a sentence containing the word
(e.g. `<<winner placeholder>>` will NOT match `<<WINNER>>`).

Each placeholder box keeps its own original position, size, and
formatting from the template — only its text content is replaced.

Events that only need a winner announcement (no nominee reveal) can use a
template with just one slide containing `<<WINNER>>` and skip
`<<NOMINEES>>` entirely; the deck will then contain only Winner slides.
