# ============================================
# 01 Hubspot Email List - BY Countries
# Local (offline) version of the original Colab cell.
#
# External files are also marked Internal if the email
# contains @uobkayhian.com or @uobkayhian.com.uk
# ============================================

from pathlib import Path
import pandas as pd


def run(base_dir: Path):
    base_dir = Path(base_dir)

    # -----------------------------
    # Paths
    # -----------------------------
    INPUT_DIR = base_dir / "01 Hubspot Email List - BY Countries"
    OUTPUT_DIR = base_dir / "Output" / "01"
    OUTPUT_FILE = OUTPUT_DIR / "SG Hubspot Email List - BY Countries.xlsx"
    INTERNAL_FILE = INPUT_DIR / "all sales.xlsx"

    # -----------------------------
    # Checks
    # -----------------------------
    if not INPUT_DIR.exists():
        raise FileNotFoundError(f"[ERROR] Input folder not found: {INPUT_DIR}")

    if not INTERNAL_FILE.exists():
        raise FileNotFoundError(
            f"[ERROR] Cannot find the internal base file 'all sales.xlsx' in:\n  {INPUT_DIR}"
        )

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # -----------------------------
    # Counters
    # -----------------------------
    internal_count = 0
    external_count = 0
    duplicate_skipped = 0
    internal_match_skipped = 0
    blank_skipped = 0
    files_processed = 0

    # -----------------------------
    # Helpers
    # -----------------------------
    def normalize_email(v):
        if pd.isna(v):
            return ""
        s = str(v).replace("\r", "").replace("\n", "").strip().lower()
        return s

    def is_internal_domain(email):
        email = normalize_email(email)
        return (
            "@uobkayhian.com" in email or
            "@uobkayhian.com.uk" in email
        )

    def read_first_sheet_excel(file_path):
        return pd.read_excel(file_path, sheet_name=0, dtype=object)

    def copy_row_a_to_e(row):
        # Keep exact logic: only copy first 5 columns
        return list(row.iloc[:5].values)

    # -----------------------------
    # Dictionaries / seen sets
    # -----------------------------
    dict_internal = set()
    dict_seen = set()

    # -----------------------------
    # Output storage
    # -----------------------------
    output_rows = []

    # -----------------------------
    # Process internal base file
    # -----------------------------
    print(f"[INFO] Processing internal base file: {INTERNAL_FILE.name}")

    df_internal = read_first_sheet_excel(INTERNAL_FILE)

    if df_internal.shape[1] < 5:
        raise ValueError(f"[ERROR] {INTERNAL_FILE.name} has fewer than 5 columns.")

    headers = list(df_internal.columns[:5]) + ["Contact Type"]

    # Column D in Excel = index 3 in pandas
    email_col_idx = 3
    if df_internal.shape[1] <= email_col_idx:
        raise ValueError(f"[ERROR] {INTERNAL_FILE.name} does not have column D (email column).")

    for i in range(1, len(df_internal)):
        email = normalize_email(df_internal.iloc[i, email_col_idx])

        if email == "":
            blank_skipped += 1
        elif email in dict_seen:
            duplicate_skipped += 1
        else:
            dict_seen.add(email)
            dict_internal.add(email)
            output_rows.append(copy_row_a_to_e(df_internal.iloc[i]) + ["Internal"])
            internal_count += 1

    # -----------------------------
    # Process other Excel files
    # -----------------------------
    for file_path in sorted(INPUT_DIR.iterdir()):
        if not file_path.is_file():
            continue

        lower_name = file_path.name.lower()

        # Skip temp Excel files
        if lower_name.startswith("~$"):
            continue

        if file_path.suffix.lower() not in [".xlsx", ".xlsm", ".xls"]:
            continue

        if lower_name in [
            "all sales.xlsx",
            "sg hubspot email list - by countries.xlsx",
        ]:
            continue

        files_processed += 1
        print(f"[INFO] Processing external file: {file_path.name}")

        try:
            df = read_first_sheet_excel(file_path)
        except Exception:
            print(f"[WARN] Skipped file (cannot open): {file_path.name}")
            continue

        if df.shape[1] <= email_col_idx:
            print(f"[WARN] Skipped file (no column D): {file_path.name}")
            continue

        for i in range(1, len(df)):
            email = normalize_email(df.iloc[i, email_col_idx])

            if email == "":
                blank_skipped += 1

            elif email in dict_internal:
                internal_match_skipped += 1

            elif email in dict_seen:
                duplicate_skipped += 1

            else:
                dict_seen.add(email)

                if is_internal_domain(email):
                    output_rows.append(copy_row_a_to_e(df.iloc[i]) + ["Internal"])
                    internal_count += 1
                else:
                    output_rows.append(copy_row_a_to_e(df.iloc[i]) + ["External"])
                    external_count += 1

    # -----------------------------
    # Build + save output
    # -----------------------------
    out_df = pd.DataFrame(output_rows, columns=headers)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        out_df.to_excel(writer, sheet_name="Contacts", index=False)

    print(f"[OK] Output created: {OUTPUT_FILE}")
    print(f"[SUMMARY] Internal kept: {internal_count}")
    print(f"[SUMMARY] External kept: {external_count}")
    print(f"[SUMMARY] External files processed: {files_processed}")
    print(f"[SUMMARY] Skipped blank emails: {blank_skipped}")
    print(f"[SUMMARY] Skipped because already Internal: {internal_match_skipped}")
    print(f"[SUMMARY] Skipped duplicate emails: {duplicate_skipped}")
