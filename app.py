"""
Hubspot Analysis Tool - local web app.

Runs the 5-step Hubspot email pipeline (originally a Google Colab notebook)
entirely on your own machine. No Google Drive, no internet required.

Start it by double-clicking start.bat (Windows). The browser opens to
http://127.0.0.1:5000 automatically.
"""

import os
import io
import sys
import socket
import threading
import webbrowser
import contextlib
import traceback
from pathlib import Path

from flask import Flask, jsonify, request, send_file, render_template

from pipeline import (
    step1_countries,
    step2_tiers,
    step3_main,
    step4_publication,
    step5_masterlist,
)

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data" / "Hubspot"

app = Flask(__name__, template_folder="templates", static_folder="static")

# ---------------------------------------------------------------------------
# Step registry
# ---------------------------------------------------------------------------
STEPS = {
    "1": {
        "title": "Step 1 — BY Countries",
        "fn": step1_countries.run,
        "input_folder": "01 Hubspot Email List - BY Countries",
        "needs": "Put 'all sales.xlsx' (internal) plus every external recipient .xlsx here.",
        "output": "Output/01/SG Hubspot Email List - BY Countries.xlsx",
    },
    "2": {
        "title": "Step 2 — BY Tiers",
        "fn": step2_tiers.run,
        "input_folder": "02 Hubspot Email List - BY Tiers",
        "needs": "Put exactly 'SG Hubspot Email List - BY Tiers.xlsx' here (cols: Email Address, Name, Region, Tier).",
        "output": "Output/02/SG Hubspot Email List - BY Tiers.xlsx",
    },
    "3": {
        "title": "Step 3 — Main",
        "fn": step3_main.run,
        "input_folder": None,
        "needs": "Uses the outputs of Step 1 and Step 2. Run those first.",
        "output": "Output/03/SG Hubspot Email List - Main.xlsx",
    },
    "4": {
        "title": "Step 4 — BY Publication",
        "fn": step4_publication.run,
        "input_folder": "04 Hubspot Email List - By Publication",
        "needs": "One sub-folder per publication (CN/EN sub-folders allowed). Place each campaign .xlsx inside its publication folder.",
        "output": "Output/04/SG Hubspot Email List - BY Publication.xlsx  +  Master Panel SG Hubspot Email - BY Publication.xlsx",
    },
    "5": {
        "title": "Step 5 — Masterlist",
        "fn": step5_masterlist.run,
        "input_folder": None,
        "needs": "Uses Step 3 + Step 4 outputs. Review the Master Panel mapping (MUTE / Into All Publications) before running.",
        "output": "05 Hubspot Email list - Masterlist/masterlist.xlsx",
    },
}

OUTPUT_FILES = {
    "1": DATA_DIR / "Output" / "01" / "SG Hubspot Email List - BY Countries.xlsx",
    "2": DATA_DIR / "Output" / "02" / "SG Hubspot Email List - BY Tiers.xlsx",
    "3": DATA_DIR / "Output" / "03" / "SG Hubspot Email List - Main.xlsx",
    "4a": DATA_DIR / "Output" / "04" / "SG Hubspot Email List - BY Publication.xlsx",
    "4b": DATA_DIR / "Output" / "04" / "Master Panel SG Hubspot Email - BY Publication.xlsx",
    "5": DATA_DIR / "05 Hubspot Email list - Masterlist" / "masterlist.xlsx",
}


def ensure_data_dirs():
    """Create the dedicated data folders if they do not exist yet."""
    folders = [
        "01 Hubspot Email List - BY Countries",
        "02 Hubspot Email List - BY Tiers",
        "03 Hubspot Email List - Main",
        "04 Hubspot Email List - By Publication",
        "05 Hubspot Email list - Masterlist",
        "Output/01", "Output/02", "Output/03", "Output/04",
    ]
    for f in folders:
        (DATA_DIR / f).mkdir(parents=True, exist_ok=True)


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/info")
def api_info():
    steps_out = []
    for sid, meta in STEPS.items():
        folder = None
        if meta["input_folder"]:
            folder = str((DATA_DIR / meta["input_folder"]))
        steps_out.append({
            "id": sid,
            "title": meta["title"],
            "needs": meta["needs"],
            "output": meta["output"],
            "input_folder": folder,
        })
    return jsonify({
        "data_dir": str(DATA_DIR),
        "steps": steps_out,
    })


@app.route("/api/status")
def api_status():
    """Report which input folders have files and which outputs exist."""
    status = {}
    for sid, meta in STEPS.items():
        info = {"output_ready": False}
        if meta["input_folder"]:
            folder = DATA_DIR / meta["input_folder"]
            count = 0
            if folder.exists():
                for root, _dirs, files in os.walk(folder):
                    for fn in files:
                        if fn.lower().endswith((".xlsx", ".xlsm", ".xls")) and not fn.startswith("~$"):
                            count += 1
            info["input_count"] = count
        status[sid] = info

    outputs = {}
    for key, path in OUTPUT_FILES.items():
        outputs[key] = path.exists()
    return jsonify({"steps": status, "outputs": outputs})


@app.route("/api/run/<step_id>", methods=["POST"])
def api_run(step_id):
    meta = STEPS.get(step_id)
    if not meta:
        return jsonify({"ok": False, "log": f"Unknown step: {step_id}"}), 404

    ensure_data_dirs()
    buf = io.StringIO()
    ok = True
    try:
        with contextlib.redirect_stdout(buf):
            meta["fn"](DATA_DIR)
    except Exception as exc:  # surface the full reason to the UI
        ok = False
        buf.write("\n")
        buf.write("=" * 60 + "\n")
        buf.write(f"[FAILED] {type(exc).__name__}: {exc}\n")
        buf.write("-" * 60 + "\n")
        buf.write(traceback.format_exc())
    return jsonify({"ok": ok, "log": buf.getvalue()})


@app.route("/api/open-folder", methods=["POST"])
def api_open_folder():
    """Open a data folder in Windows Explorer (localhost convenience)."""
    rel = (request.json or {}).get("path", "")
    target = (DATA_DIR / rel).resolve() if rel else DATA_DIR
    try:
        target.mkdir(parents=True, exist_ok=True)
        if sys.platform.startswith("win"):
            os.startfile(str(target))  # noqa: S606 - local trusted path
        elif sys.platform == "darwin":
            os.system(f'open "{target}"')
        else:
            os.system(f'xdg-open "{target}"')
        return jsonify({"ok": True, "path": str(target)})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/download/<key>")
def download(key):
    path = OUTPUT_FILES.get(key)
    if not path or not path.exists():
        return "File not found. Run the step that produces it first.", 404
    return send_file(str(path), as_attachment=True, download_name=path.name)


# ---------------------------------------------------------------------------
# Startup
# ---------------------------------------------------------------------------
def _find_free_port(preferred=5000):
    for port in [preferred, 5001, 5002, 5050, 8000, 8080]:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", port)) != 0:
                return port
    return preferred


def main():
    ensure_data_dirs()
    port = _find_free_port(5000)
    url = f"http://127.0.0.1:{port}"

    # Only open the browser in the main (non-reloader) process.
    if not os.environ.get("WERKZEUG_RUN_MAIN"):
        threading.Timer(1.2, lambda: webbrowser.open(url)).start()

    print("=" * 64)
    print("  Hubspot Analysis Tool")
    print(f"  Open this in your browser:  {url}")
    print(f"  Data folder:                {DATA_DIR}")
    print("  (Close this window to stop the tool.)")
    print("=" * 64)

    app.run(host="127.0.0.1", port=port, debug=False)


if __name__ == "__main__":
    main()
