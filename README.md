# Hubspot Analysis Tool

A **portable, offline** version of the Hubspot email-list pipeline that used to run
in Google Colab. It runs as a small local website on your own computer — **no Google
Drive and no internet needed** to use it (internet is only required once, the very
first time, to install the Python packages).

It turns your raw Hubspot recipient exports into a finished, styled
**`masterlist.xlsx`** engagement report in 5 steps.

---

## Quick start (Windows)

1. **Download** this repository as a ZIP from GitHub → *Code ▸ Download ZIP* → unzip it
   anywhere (e.g. your Desktop).
2. Make sure **Python 3.9+** is installed. Get it from
   <https://www.python.org/downloads/> and **tick “Add Python to PATH”** during install.
3. **Double-click `start.bat`.**
   - The first run creates a private environment and installs Flask / pandas / openpyxl
     (one-time, needs internet, takes a minute or two).
   - Every run after that starts instantly.
4. Your browser opens automatically at **http://127.0.0.1:5000**.
   Keep the black window open while you work; close it to stop the tool.

---

## Where do my files go?

Everything lives inside **one dedicated folder**:

```
data/Hubspot/
├── 01 Hubspot Email List - BY Countries/   ← all sales.xlsx  +  external recipient files
├── 02 Hubspot Email List - BY Tiers/       ← SG Hubspot Email List - BY Tiers.xlsx
├── 03 Hubspot Email List - Main/           (output only)
├── 04 Hubspot Email List - By Publication/ ← one sub-folder per publication
├── 05 Hubspot Email list - Masterlist/     ← final masterlist.xlsx appears here
└── Output/                                  ← 01–04 generated files
```

The website shows you the **exact full path** of this folder and gives each step an
**“Open input folder”** button, so you never have to hunt for it.

---

## The 5 steps

| Step | What it does | Reads | Writes |
|------|--------------|-------|--------|
| **1 — BY Countries** | Merges internal + external recipient files, de-dupes emails, tags **Internal/External** (`@uobkayhian.com` ⇒ Internal). | `01 .../all sales.xlsx` + every other `.xlsx` in folder 01 | `Output/01/SG Hubspot Email List - BY Countries.xlsx` |
| **2 — BY Tiers** | Validates the tier file (needs columns **Email Address, Name, Region, Tier**) and counts people per region. | `02 .../SG Hubspot Email List - BY Tiers.xlsx` | `Output/02/...BY Tiers.xlsx` |
| **3 — Main** | Combines Countries + Tiers, overlays tier onto each contact, and re-labels **Tier 1–3 + China ⇒ Hong Kong**. | Output of Steps 1 & 2 | `Output/03/SG Hubspot Email List - Main.xlsx` |
| **4 — BY Publication** | Walks the publication folder tree (CN/EN aware), de-dupes by campaign ID, and builds the **Master Panel** mapping. | Folder `04 ...` | `Output/04/...BY Publication.xlsx` **+** `Master Panel ...xlsx` |
| **5 — Masterlist** | Builds the final multi-tab engagement report (open/click rates, loyalty tiers, per-publication breakdowns). | Output of Steps 3 & 4 | `05 .../masterlist.xlsx` |

Run them in order with each step’s **▶ Run** button, or hit **“Run all steps (1 → 5)”**.

### Folder 04 layout (publications)

```
04 Hubspot Email List - By Publication/
├── CIO Quarterly Review/
│       └── file.xlsx
└── China Bi-Weekly/
        ├── CN/  └── file.xlsx
        └── EN/  └── file.xlsx
```

- One sub-folder per publication. CN / EN variants go in `CN` / `EN` sub-folders.
- Do **not** drop Excel files loose in the root of folder 04.

---

## The mapping file (between Step 4 and Step 5)

Step 4 generates **`Output/04/Master Panel SG Hubspot Email - BY Publication.xlsx`**.
Its **Campaign Mapping** tab has two control columns you can edit before running Step 5:

- **MUTE OR NOT** — `1` excludes a campaign entirely; blank/`0` keeps it.
- **Into All Publications or not?** — `1`/blank includes it in the “0 All Publications”
  and loyalty summaries; `0` keeps it only inside its own publication table.

A **blank reference template** showing the exact columns is included at
[`samples/Master Panel - TEMPLATE.xlsx`](samples/Master%20Panel%20-%20TEMPLATE.xlsx).
Step 4 rebuilds the real file from your data each time it runs, so you normally just
tweak the two columns above and move on.

---

## Chrome extension — HubSpot Bulk Recipient Export (ZIP)

The site also hosts the **HubSpot Bulk Recipient Export** Chrome extension (the tool that
adds a *“Download selected (ZIP)”* button to HubSpot’s email manage page). On the website,
use **Download extension (.zip)**, then install it once:

1. Unzip it to a permanent folder (Chrome loads it from there — don’t delete it).
2. Open `chrome://extensions` → turn on **Developer mode** (top-right).
3. Click **Load unpacked** → pick the unzipped folder (the one with `manifest.json`).
4. In HubSpot → *Marketing ▸ Email*, tick the emails and click **Download selected (ZIP)**.

> Chrome only allows true one-click installs from the Web Store, so unlisted extensions
> are installed “unpacked” via the steps above. To share with a colleague, send them the
> ZIP and these 4 steps — they don’t need this analysis tool running to use the extension.
> The extension source also lives in [`extension/`](extension/).

## Notes

- This is a 1-to-1 port of the original Colab notebook logic — only the file paths
  changed (Google Drive → the local `data/Hubspot` folder).
- Outputs are plain `.xlsx` files you can open in Excel.
- To reset, delete the files inside `data/Hubspot/Output` and re-run.
- Your data files are **git-ignored** — they never get committed if you push this repo.

## macOS / Linux

There’s no `.bat`, but it still runs:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```
