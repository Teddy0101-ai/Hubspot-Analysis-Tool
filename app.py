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
import zipfile
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
EXTENSION_DIR = APP_DIR / "extension"

app = Flask(__name__, template_folder="templates", static_folder="static")

# ---------------------------------------------------------------------------
# Step registry
# ---------------------------------------------------------------------------
STEPS = {
    "1": {
        "name": "Countries",
        "phase": "Contacts",
        "fn": step1_countries.run,
        "input_folder": "01 Hubspot Email List - BY Countries",
        "does": "Merge internal + external recipient lists, remove duplicates, tag each as Internal / External.",
        "provide": "Put all sales.xlsx (your internal list) plus every external recipient .xlsx into the folder. Email must be in column D.",
        "output": "SG Hubspot Email List - BY Countries.xlsx",
        "static": True,
    },
    "2": {
        "name": "Tiers",
        "phase": "Contacts",
        "fn": step2_tiers.run,
        "input_folder": "02 Hubspot Email List - BY Tiers",
        "does": "Check the tier file is valid and count people per region.",
        "provide": "One file named exactly 'SG Hubspot Email List - BY Tiers.xlsx' with columns: Email Address, Name, Region, Tier.",
        "output": "SG Hubspot Email List - BY Tiers.xlsx",
        "static": True,
    },
    "3": {
        "name": "Contact master",
        "phase": "Contacts",
        "fn": step3_main.run,
        "input_folder": None,
        "does": "Combine Countries + Tiers into one master list (Tier 1–3 in China are relabelled Hong Kong).",
        "provide": "Nothing — uses the results of steps 1 and 2.",
        "output": "SG Hubspot Email List - Main.xlsx",
        "static": False,
    },
    "4": {
        "name": "Publications",
        "phase": "Campaigns",
        "fn": step4_publication.run,
        "input_folder": "04 Hubspot Email List - By Publication",
        "does": "Scan the publication folders, de-duplicate campaigns, and build the Master Panel mapping.",
        "provide": "One sub-folder per publication (CN / EN sub-folders allowed). Put each campaign .xlsx inside its publication folder.",
        "output": "SG Hubspot Email List - BY Publication.xlsx + Master Panel ...xlsx",
        "static": False,
    },
    "5": {
        "name": "Masterlist report",
        "phase": "Report",
        "fn": step5_masterlist.run,
        "input_folder": None,
        "does": "Build the final multi-tab engagement report (open / click rates, loyalty tiers, per-publication breakdowns).",
        "provide": "Nothing — uses steps 3 and 4. Tip: review the Master Panel mapping first (see the note below the steps).",
        "output": "masterlist.xlsx",
        "static": False,
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
        rel_folder = meta["input_folder"] or ""
        steps_out.append({
            "id": sid,
            "name": meta["name"],
            "phase": meta["phase"],
            "does": meta["does"],
            "provide": meta["provide"],
            "output": meta["output"],
            "input_folder": rel_folder,
            "static": meta.get("static", False),
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
# Chrome extension (HubSpot Bulk Recipient Export)
# ---------------------------------------------------------------------------
def _extension_meta():
    """Read name/version/description from the extension manifest."""
    manifest = EXTENSION_DIR / "manifest.json"
    meta = {"name": "HubSpot Bulk Recipient Export (ZIP)", "version": "", "description": ""}
    try:
        import json
        data = json.loads(manifest.read_text(encoding="utf-8"))
        meta["name"] = data.get("name", meta["name"])
        meta["version"] = data.get("version", "")
        meta["description"] = data.get("description", "")
    except Exception:
        pass
    meta["available"] = (EXTENSION_DIR / "manifest.json").exists()
    return meta


@app.route("/api/extension")
def api_extension():
    return jsonify(_extension_meta())


@app.route("/download/extension")
def download_extension():
    """Zip the unpacked extension folder on the fly and serve it."""
    if not (EXTENSION_DIR / "manifest.json").exists():
        return "Extension folder not found.", 404

    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in sorted(EXTENSION_DIR.rglob("*")):
            if path.is_file():
                zf.write(path, path.relative_to(EXTENSION_DIR))
    mem.seek(0)
    return send_file(
        mem,
        mimetype="application/zip",
        as_attachment=True,
        download_name="HubSpot-Bulk-Recipient-Export.zip",
    )


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
