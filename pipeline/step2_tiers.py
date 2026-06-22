# ============================================
# 02 Hubspot Email List - BY Tiers
# Validate SG Hubspot Email List - BY Tiers.xlsx,
# count people by Region, and copy it to Output/02
# ============================================

import shutil
from pathlib import Path
import pandas as pd


def run(base_dir: Path):
    base_dir = Path(base_dir)

    # -----------------------------
    # Paths
    # -----------------------------
    INPUT_DIR = base_dir / "02 Hubspot Email List - BY Tiers"
    OUTPUT_DIR = base_dir / "Output" / "02"
    SOURCE_FILE = INPUT_DIR / "SG Hubspot Email List - BY Tiers.xlsx"
    DEST_FILE = OUTPUT_DIR / "SG Hubspot Email List - BY Tiers.xlsx"

    # -----------------------------
    # Required columns
    # -----------------------------
    REQUIRED_COLS = ["Email Address", "Name", "Region", "Tier"]

    # -----------------------------
    # Checks
    # -----------------------------
    if not INPUT_DIR.exists():
        raise FileNotFoundError(f"[ERROR] Input folder not found: {INPUT_DIR}")

    if not SOURCE_FILE.exists():
        raise FileNotFoundError(
            f"[ERROR] File not found. Expected exactly:\n  {SOURCE_FILE}"
        )

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # -----------------------------
    # Read file
    # -----------------------------
    df = pd.read_excel(SOURCE_FILE, sheet_name=0, dtype=object)

    # -----------------------------
    # Validate columns
    # -----------------------------
    actual_cols = [str(c).strip() for c in df.columns.tolist()]
    missing_cols = [c for c in REQUIRED_COLS if c not in actual_cols]

    if missing_cols:
        raise ValueError(
            "[ERROR] Column check failed.\n"
            f"Missing required column(s): {missing_cols}\n"
            f"Actual columns: {actual_cols}"
        )

    print("[OK] File exists.")
    print(f"[OK] Required columns found: {REQUIRED_COLS}")

    # -----------------------------
    # Count people by region
    # -----------------------------
    region_col = "Region"

    region_counts = (
        df[region_col]
        .astype(str)
        .str.strip()
        .replace("nan", pd.NA)
        .dropna()
        .value_counts(dropna=False)
        .sort_index()
    )

    print("\n[SUMMARY] Number of people by Region:")
    for region, count in region_counts.items():
        print(f" - {region}: {count}")

    # -----------------------------
    # Copy file to Output/02
    # -----------------------------
    shutil.copy2(SOURCE_FILE, DEST_FILE)

    print(f"\n[OK] File copied to: {DEST_FILE}")
