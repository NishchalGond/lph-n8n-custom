#!/usr/bin/env python3
"""
master_consolidator.py

Consolidates ONE BATCH of source Excel/CSV files (heterogeneous headers,
one property-listing sheet per file) into a single output workbook that
contains ONLY these 22 canonical columns, in this order:

    Name, Community, Sub-Community, Building/Cluster, Unit Number, Size,
    Plot Reg. No, Plot Number, DMNO, DMsubno, Bedroom, Type (Buyer/Seller),
    Mobile 1, Mobile 2, Mobile 3, Email Address, PI number, Nationality,
    Property Type, Date, Procedure Value, Developer, Project

Design notes (why it's built this way):
- Source files do NOT share one schema. Headers vary in name, order, and
  which row they sit on. So every file gets its own header detection +
  column-mapping pass rather than a fixed column-position copy.
- Per-file failures are caught and skipped; they do not stop the batch.
  Only an error outside the per-file loop (bad incoming dir, cannot write
  output, etc.) is treated as fatal (exit code 1) -- matching the n8n
  workflow's "Run Was Fatal?" check on exitCode == 1.
- Phone-like columns are handled specially: a file may have multiple
  DIFFERENT phone numbers under different headers (Mobile, Phone 2,
  Contact No, Tel...). These are assigned in the order they appear in the
  sheet to Mobile 1 / Mobile 2 / Mobile 3. Anything beyond 3 is dropped
  and reported in unmapped_columns so nothing is silently lost without a
  trace.
- Columns the alias table can't confidently match are SKIPPED (not
  guessed), and listed per-file in unmapped_columns in the JSON summary,
  so you can see exactly what didn't make it into the batch sheet.

Usage:
    python3 master_consolidator.py <incoming_dir> <batch_output_path> <run_number>

Prints exactly one JSON object to stdout (last line) for n8n to parse:
{
  "status": "success" | "partial" | "failed" | "fatal_error",
  "run_number": "001",
  "batch_output_file": "/abs/path/Consolidated_Batch_001.xlsx",
  "files_total": 60,
  "files_succeeded": 58,
  "files_failed": 2,
  "row_count": 12345,
  "failed_files": [{"file": "...", "error": "..."}],
  "unmapped_columns": {"some_file.xlsx": ["Weird Header 1", "..."]}
}
Exit code 0 for success/partial/failed (per-file issues never halt the run).
Exit code 1 only for a fatal/system-level error.
"""

import sys
import os
import re
import json
import glob
import difflib
import traceback
from datetime import datetime

import pandas as pd

CANONICAL_FIELDS = [
    "Name", "Community", "Sub-Community", "Building/Cluster", "Unit Number",
    "Size", "Plot Reg. No", "Plot Number", "DMNO", "DMsubno", "Bedroom",
    "Type (Buyer/Seller)", "Mobile 1", "Mobile 2", "Mobile 3",
    "Email Address", "PI number", "Nationality", "Property Type", "Date",
    "Procedure Value", "Developer", "Project",
]

# Canonical field -> list of normalized alias strings/patterns that should
# map to it. Matching is done on a normalized header (lowercased, stripped,
# punctuation collapsed to single spaces).
ALIASES = {
    "Name":                 ["name", "full name", "owner name", "contact name", "client name", "customer name"],
    "Community":            ["community", "community name"],
    "Sub-Community":        ["sub community", "subcommunity", "sub comm"],
    "Building/Cluster":     ["building", "cluster", "building cluster", "tower", "tower name", "building name"],
    "Unit Number":          ["unit no", "unit number", "unit num", "unit", "apt no", "apartment no", "flat no"],
    "Size":                 ["size", "area", "sq ft", "sqft", "area sqft", "size sqft"],
    "Plot Reg. No":         ["plot reg no", "plot registration no", "plot reg number", "plot registration number"],
    "Plot Number":          ["plot no", "plot number", "plot num"],
    "DMNO":                 ["dmno", "dm no", "dm number"],
    "DMsubno":              ["dmsubno", "dm subno", "dm sub no", "dm sub number"],
    "Bedroom":              ["bedroom", "bed room", "bedrooms", "beds", "br", "no of bedrooms"],
    "Type (Buyer/Seller)":  ["type buyer seller", "buyer seller", "type", "role", "buyer/seller"],
    "Email Address":        ["email address", "email", "e mail", "email id"],
    "PI number":            ["pi number", "pi no", "pi num"],
    "Nationality":          ["nationality", "nation"],
    "Property Type":        ["property type", "prop type", "unit type"],
    "Date":                 ["date", "transaction date", "reg date", "registration date"],
    "Procedure Value":      ["procedure value", "value", "amount", "price", "transaction value"],
    "Developer":            ["developer", "developer name"],
    "Project":              ["project", "project name"],
}

# Header patterns that indicate a phone-number-like column. These are
# collected in sheet order and distributed across Mobile 1/2/3.
PHONE_PATTERNS = ["mobile", "phone", "contact no", "contact number", "tel", "telephone", "cell"]

FUZZY_CUTOFF = 0.78
HEADER_SCAN_ROWS = 10  # how many top rows to scan when guessing the header row


def normalize(s):
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


# Build a flat lookup: normalized alias -> canonical field, plus a list of
# all normalized aliases for fuzzy matching.
ALIAS_LOOKUP = {}
ALL_ALIASES = []
for canon, aliases in ALIASES.items():
    for a in aliases:
        na = normalize(a)
        ALIAS_LOOKUP[na] = canon
        ALL_ALIASES.append(na)


def is_phone_header(norm_header):
    return any(p in norm_header for p in PHONE_PATTERNS)


def map_header(raw_header):
    """Return ('phone', None) | (canonical_field, None) | (None, None)."""
    norm = normalize(raw_header)
    if not norm:
        return None
    if is_phone_header(norm):
        return "PHONE"
    if norm in ALIAS_LOOKUP:
        return ALIAS_LOOKUP[norm]
    # substring match (e.g. "unit number (sqft)" contains "unit number")
    for alias, canon in ALIAS_LOOKUP.items():
        if alias in norm or norm in alias:
            return canon
    # fuzzy fallback
    close = difflib.get_close_matches(norm, ALL_ALIASES, n=1, cutoff=FUZZY_CUTOFF)
    if close:
        return ALIAS_LOOKUP[close[0]]
    return None


def find_header_row(raw_df):
    """Scan the first HEADER_SCAN_ROWS rows and pick the one whose cells
    match the most known aliases (canonical or phone). Falls back to row 0.
    """
    best_row = 0
    best_score = -1
    n = min(HEADER_SCAN_ROWS, len(raw_df))
    for i in range(n):
        row = raw_df.iloc[i]
        score = 0
        for cell in row:
            m = map_header(cell)
            if m:
                score += 1
        if score > best_score:
            best_score = score
            best_row = i
    return best_row, best_score


def read_source_file(path):
    """Load a source file's data with the header row auto-detected.
    Returns a DataFrame with the file's ORIGINAL headers as columns.
    """
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        raw = pd.read_csv(path, header=None, dtype=str, keep_default_na=False)
    else:
        engine = "xlrd" if ext == ".xls" else "openpyxl"
        raw = pd.read_excel(path, header=None, dtype=str, engine=engine, keep_default_na=False)

    if raw.empty:
        raise ValueError("file is empty")

    header_row, score = find_header_row(raw)
    if score <= 0:
        raise ValueError("could not detect a recognizable header row")

    headers = raw.iloc[header_row].tolist()
    data = raw.iloc[header_row + 1:].reset_index(drop=True)
    data.columns = [str(h) if h is not None else "" for h in headers]
    # drop fully blank rows
    data = data[~(data.apply(lambda r: all(str(v).strip() == "" for v in r), axis=1))]
    return data


def map_file_to_canonical(df, filename):
    """Map a source DataFrame (original headers) onto the 22 canonical
    columns. Returns (canonical_df, unmapped_headers list).
    """
    phone_cols_in_order = []
    canon_col_for = {}   # canonical field -> original column name (first match wins)
    unmapped = []

    for col in df.columns:
        result = map_header(col)
        if result is None:
            if str(col).strip() != "":
                unmapped.append(str(col))
            continue
        if result == "PHONE":
            phone_cols_in_order.append(col)
        else:
            if result not in canon_col_for:
                canon_col_for[result] = col
            else:
                # duplicate mapping to same canonical field -- keep first,
                # flag the rest as unmapped so nothing is silently dropped
                unmapped.append(str(col))

    # distribute phone columns across Mobile 1/2/3 in sheet order
    mobile_slots = ["Mobile 1", "Mobile 2", "Mobile 3"]
    for i, col in enumerate(phone_cols_in_order):
        if i < 3:
            canon_col_for[mobile_slots[i]] = col
        else:
            unmapped.append(str(col))

    out = pd.DataFrame()
    for field in CANONICAL_FIELDS:
        if field in canon_col_for:
            out[field] = df[canon_col_for[field]].astype(str).str.strip()
        else:
            out[field] = ""

    # drop rows that are entirely empty across all canonical fields
    out = out[~(out.apply(lambda r: all(str(v).strip() == "" for v in r), axis=1))]
    return out, unmapped


def main():
    if len(sys.argv) < 4:
        print(json.dumps({
            "status": "fatal_error",
            "error": "usage: master_consolidator.py <incoming_dir> <batch_output_path> <run_number>",
        }))
        sys.exit(1)

    incoming_dir = sys.argv[1]
    batch_output_path = sys.argv[2]
    run_number = sys.argv[3]

    try:
        if not os.path.isdir(incoming_dir):
            raise FileNotFoundError(f"incoming_dir not found: {incoming_dir}")

        patterns = ["*.xlsx", "*.xls", "*.xlsm", "*.csv"]
        files = []
        for p in patterns:
            files.extend(glob.glob(os.path.join(incoming_dir, p)))
        files = sorted(set(files))

        if not files:
            raise FileNotFoundError(f"no source files found in {incoming_dir}")

        frames = []
        failed_files = []
        unmapped_columns = {}
        files_succeeded = 0

        for path in files:
            fname = os.path.basename(path)
            try:
                raw_df = read_source_file(path)
                canon_df, unmapped = map_file_to_canonical(raw_df, fname)
                if unmapped:
                    unmapped_columns[fname] = unmapped
                frames.append(canon_df)
                files_succeeded += 1
            except Exception as e:
                failed_files.append({"file": fname, "error": str(e)})
                continue

        if frames:
            combined = pd.concat(frames, ignore_index=True)
        else:
            combined = pd.DataFrame(columns=CANONICAL_FIELDS)

        os.makedirs(os.path.dirname(os.path.abspath(batch_output_path)), exist_ok=True)
        combined.to_excel(batch_output_path, index=False, columns=CANONICAL_FIELDS)

        files_total = len(files)
        files_failed = len(failed_files)

        if files_succeeded == 0:
            status = "failed"
        elif files_failed == 0:
            status = "success"
        else:
            status = "partial"

        abs_output = os.path.abspath(batch_output_path)
        stat = os.stat(abs_output)

        result = {
            "status": status,
            "run_number": run_number,
            "batch_output_file": abs_output,
            "batch_file_name": os.path.basename(abs_output),
            "batch_directory": os.path.dirname(abs_output),
            "batch_extension": os.path.splitext(abs_output)[1],
            "batch_size_kb": round(stat.st_size / 1024, 2),
            "batch_last_write_time": datetime.fromtimestamp(stat.st_mtime).isoformat(),
            "files_total": files_total,
            "files_succeeded": files_succeeded,
            "files_failed": files_failed,
            "row_count": int(len(combined)),
            "failed_files": failed_files,
            "unmapped_columns": unmapped_columns,
            "generated_at": datetime.utcnow().isoformat() + "Z",
        }

        # Per-batch audit trail on disk (separate from the 5-column master
        # manifest) so failed/unmapped details aren't lost even though the
        # manifest itself intentionally stays to just those 5 columns.
        try:
            audit_dir = os.path.join(os.path.dirname(abs_output), "batch_logs")
            os.makedirs(audit_dir, exist_ok=True)
            audit_path = os.path.join(audit_dir, f"batch_{run_number}_detail.json")
            with open(audit_path, "w") as f:
                json.dump(result, f, indent=2)
        except Exception:
            pass  # audit trail is best-effort, never blocks the batch

        print(json.dumps(result))
        sys.exit(0)

    except Exception as e:
        print(json.dumps({
            "status": "fatal_error",
            "run_number": run_number,
            "error": str(e),
            "traceback": traceback.format_exc()[-1500:],
        }))
        sys.exit(1)


if __name__ == "__main__":
    main()
