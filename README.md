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
of header names, how slides are recognized in the template by their
`<<NOMINEES>>` / `<<WINNER>>` / `<<AWARD_TEXT>>` / `<<ZONE>>` placeholders,
the winner-picking rule, the legacy-template fallback — is all identical
to the command-line `.bat` version. See the main project README for the
full explanation of those rules; this app is only a different way to run
the same engine.
