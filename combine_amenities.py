"""
Combine amenities from two spreadsheets into Spreadsheet A.

Rows are matched by Costar ID first, then by address (case-insensitive).
Amenities from both rows are merged and deduplicated (order preserved).
The result is written back to Spreadsheet A.

Usage:
    python combine_amenities.py

Edit the CONFIG section below to point at your files and column names.
"""

import os
import pandas as pd

# ── CONFIG ────────────────────────────────────────────────────────────────────
SPREADSHEET_A     = "PI Whyte Ave.xlsx"           # Input/output file for Dataset A
SPREADSHEET_B     = "Costar Office Dataset.xlsx"  # Source of additional amenities
OUTPUT_PATH       = "PI Whyte Ave updated.xlsx"   # Where to save the result

# Column names in Spreadsheet A
A_ID_COL         = "Costar ID"       # Primary match key (set to None to skip)
A_ADDRESS_COL    = "Primary address" # Fallback match key
A_AMENITIES_COL  = "Amenities"       # Column to update

# Column names in Spreadsheet B
B_ID_COL         = "Costar ID"       # Primary match key (set to None to skip)
B_ADDRESS_COL    = "Property Address"
B_AMENITIES_COL  = "Amenities"

DELIMITER         = ","              # Separator used inside amenity cells
OUTPUT_DELIMITER  = ", "            # Separator used when writing merged amenities
# ─────────────────────────────────────────────────────────────────────────────


def load_file(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[-1].lower()
    if ext == ".csv":
        return pd.read_csv(path, dtype=str)
    elif ext in (".xlsx", ".xls"):
        return pd.read_excel(path, dtype=str)
    raise ValueError(f"Unsupported file type: {ext}")


def split_amenities(raw) -> list[str]:
    """Split a delimited amenities cell into a list of stripped strings."""
    if pd.isna(raw) or str(raw).strip() == "":
        return []
    return [a.strip() for a in str(raw).split(DELIMITER) if a.strip()]


def merge_amenities(list_a: list[str], list_b: list[str]) -> list[str]:
    """Combine two amenity lists, keeping unique values (case-insensitive, A first)."""
    seen = set()
    merged = []
    for item in list_a + list_b:
        key = item.lower()
        if key not in seen:
            seen.add(key)
            merged.append(item)
    return merged


def normalise_address(addr) -> str:
    if pd.isna(addr):
        return ""
    return " ".join(str(addr).strip().lower().split())


def build_b_lookup(df_b: pd.DataFrame) -> tuple[dict, dict]:
    """Build two lookups from Dataset B: by ID and by normalised address."""
    id_lookup   = {}
    addr_lookup = {}

    has_id = B_ID_COL and B_ID_COL in df_b.columns

    for _, row in df_b.iterrows():
        amenities = split_amenities(row[B_AMENITIES_COL])

        if has_id and not pd.isna(row[B_ID_COL]):
            bid = str(row[B_ID_COL]).strip()
            if bid:
                existing = id_lookup.get(bid, [])
                id_lookup[bid] = merge_amenities(existing, amenities)

        addr = normalise_address(row[B_ADDRESS_COL])
        if addr:
            existing = addr_lookup.get(addr, [])
            addr_lookup[addr] = merge_amenities(existing, amenities)

    return id_lookup, addr_lookup


def main():
    df_a = load_file(SPREADSHEET_A)
    df_b = load_file(SPREADSHEET_B)

    print(f"Dataset A: {len(df_a)} rows")
    print(f"Dataset B: {len(df_b)} rows")

    id_lookup, addr_lookup = build_b_lookup(df_b)

    has_id = A_ID_COL and A_ID_COL in df_a.columns
    updated = 0

    for idx, row in df_a.iterrows():
        b_amenities = None

        # 1. Match by ID
        if has_id and not pd.isna(row.get(A_ID_COL)):
            aid = str(row[A_ID_COL]).strip()
            if aid in id_lookup:
                b_amenities = id_lookup[aid]

        # 2. Match by address
        if b_amenities is None:
            addr = normalise_address(row[A_ADDRESS_COL])
            if addr in addr_lookup:
                b_amenities = addr_lookup[addr]

        if b_amenities is None:
            continue  # No match in Dataset B

        a_amenities = split_amenities(row[A_AMENITIES_COL])
        merged = merge_amenities(a_amenities, b_amenities)

        if merged != a_amenities:
            df_a.at[idx, A_AMENITIES_COL] = OUTPUT_DELIMITER.join(merged)
            updated += 1

    print(f"Rows updated: {updated}")

    ext = os.path.splitext(OUTPUT_PATH)[-1].lower()
    if ext == ".csv":
        df_a.to_csv(OUTPUT_PATH, index=False)
    else:
        df_a.to_excel(OUTPUT_PATH, index=False)

    print(f"Saved to: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
