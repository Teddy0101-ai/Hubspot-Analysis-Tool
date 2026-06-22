# ============================================
# 03 Hubspot Email List - Main
# Merge:
#   Output/01/SG Hubspot Email List - BY Countries.xlsx
#   Output/02/SG Hubspot Email List - BY Tiers.xlsx
#
# For Tier 1 / Tier 2 / Tier 3, if Country = China,
# change Country to Hong Kong.
# ============================================

from pathlib import Path
import pandas as pd


def run(base_dir: Path):
    base_dir = Path(base_dir)

    # -----------------------------
    # Paths
    # -----------------------------
    COUNTRIES_PATH = base_dir / "Output" / "01" / "SG Hubspot Email List - BY Countries.xlsx"
    TIERS_PATH = base_dir / "Output" / "02" / "SG Hubspot Email List - BY Tiers.xlsx"
    OUTPUT_DIR = base_dir / "Output" / "03"
    OUTPUT_PATH = OUTPUT_DIR / "SG Hubspot Email List - Main.xlsx"

    # -----------------------------
    # Checks
    # -----------------------------
    if not COUNTRIES_PATH.exists():
        raise FileNotFoundError(
            f"[ERROR] Cannot find file (run Step 1 first): {COUNTRIES_PATH}"
        )

    if not TIERS_PATH.exists():
        raise FileNotFoundError(
            f"[ERROR] Cannot find file (run Step 2 first): {TIERS_PATH}"
        )

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # -----------------------------
    # Helper functions
    # -----------------------------
    def collapse_spaces(s):
        s = str(s)
        while "  " in s:
            s = s.replace("  ", " ")
        return s

    def clean_text(v):
        if pd.isna(v):
            return ""
        s = str(v)
        s = s.replace("\r", " ")
        s = s.replace("\n", " ")
        s = s.replace(chr(160), " ")
        s = collapse_spaces(s).strip()
        return s

    def normalize_email(v):
        return clean_text(v).lower()

    def get_first_token(s):
        s = collapse_spaces(str(s)).strip()
        if s == "":
            return ""
        return s.split(" ")[0]

    def get_last_token(s):
        s = collapse_spaces(str(s)).strip()
        if s == "":
            return ""
        return s.split(" ")[-1]

    def build_name(first_name_val, last_name_val):
        fn = clean_text(first_name_val)
        ln = clean_text(last_name_val)

        if fn == "" and ln == "":
            return ""
        if fn == "":
            return ln
        if ln == "":
            return fn

        fn_lower = collapse_spaces(fn).lower()
        ln_lower = collapse_spaces(ln).lower()

        first_token = get_first_token(fn_lower)
        last_token = get_last_token(fn_lower)

        if first_token == ln_lower:
            return fn
        elif last_token == ln_lower:
            return fn
        elif fn_lower == ln_lower:
            return fn
        else:
            return f"{fn} {ln}"

    def normalize_base_contact_type(v):
        s = clean_text(v).lower()
        if s == "internal":
            return "Internal"
        elif s == "external":
            return "External"
        else:
            return "External"

    def normalize_tier(v):
        s = clean_text(v).replace(" ", "").upper()
        if s in ["1", "TIER1"]:
            return "Tier 1"
        elif s in ["2", "TIER2"]:
            return "Tier 2"
        elif s in ["3", "TIER3"]:
            return "Tier 3"
        else:
            return ""

    # -----------------------------
    # Read BY Countries
    # col 2 = first name, col 3 = last name,
    # col 4 = email, col 5 = country, col 6 = contact type
    # -----------------------------
    print("[INFO] Reading BY Countries...")
    df_countries = pd.read_excel(COUNTRIES_PATH, sheet_name=0, dtype=object)

    if df_countries.shape[1] < 6:
        raise ValueError(
            f"[ERROR] BY Countries file must have at least 6 columns. Found: {df_countries.shape[1]}"
        )

    dict_rows = {}
    dict_tier = {}

    for i in range(1, len(df_countries)):
        email = normalize_email(df_countries.iloc[i, 3])

        if email != "":
            first_name = clean_text(df_countries.iloc[i, 1])
            last_name = clean_text(df_countries.iloc[i, 2])
            nm = build_name(first_name, last_name)
            country = clean_text(df_countries.iloc[i, 4])
            ctype = normalize_base_contact_type(df_countries.iloc[i, 5])

            if email not in dict_rows:
                dict_rows[email] = [nm, country, ctype]

    # -----------------------------
    # Read BY Tiers
    # col 1 = email, col 4 = tier
    # -----------------------------
    print("[INFO] Reading BY Tiers...")
    df_tiers = pd.read_excel(TIERS_PATH, sheet_name=0, dtype=object)

    if df_tiers.shape[1] < 4:
        raise ValueError(
            f"[ERROR] BY Tiers file must have at least 4 columns. Found: {df_tiers.shape[1]}"
        )

    for i in range(1, len(df_tiers)):
        email = normalize_email(df_tiers.iloc[i, 0])
        tier_text = normalize_tier(df_tiers.iloc[i, 3])

        if email != "" and tier_text != "":
            dict_tier[email] = tier_text

    # -----------------------------
    # Overlay tier onto dict_rows
    # -----------------------------
    for email, tier_text in dict_tier.items():
        if email in dict_rows:
            item = dict_rows[email]
            item[2] = tier_text
            dict_rows[email] = item

    # -----------------------------
    # Tier 1/2/3 + country China -> Hong Kong
    # -----------------------------
    for email, item in dict_rows.items():
        nm, country, ctype = item
        if ctype in ["Tier 1", "Tier 2", "Tier 3"] and clean_text(country).lower() == "china":
            country = "Hong Kong"
        dict_rows[email] = [nm, country, ctype]

    # -----------------------------
    # Write output in order:
    # Tier 1, Tier 2, Tier 3, Internal, External, Others
    # -----------------------------
    output_rows = []

    def write_rows_for_type(wanted_type):
        for email, row_item in dict_rows.items():
            ct = str(row_item[2]).strip().lower()
            if ct == wanted_type.lower():
                output_rows.append([email, row_item[0], row_item[1], row_item[2]])

    def write_rows_for_other():
        valid_types = {"tier 1", "tier 2", "tier 3", "internal", "external"}
        for email, row_item in dict_rows.items():
            ct = str(row_item[2]).strip().lower()
            if ct not in valid_types:
                output_rows.append([email, row_item[0], row_item[1], row_item[2]])

    write_rows_for_type("Tier 1")
    write_rows_for_type("Tier 2")
    write_rows_for_type("Tier 3")
    write_rows_for_type("Internal")
    write_rows_for_type("External")
    write_rows_for_other()

    out_df = pd.DataFrame(output_rows, columns=["Email", "Name", "Country", "Contact Type"])

    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        out_df.to_excel(writer, sheet_name="Main", index=False)

    print(f"[OK] Created: {OUTPUT_PATH}")
    print(f"[SUMMARY] Rows written: {len(output_rows)}")
