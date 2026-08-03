#!/usr/bin/env python3
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

PHONE_PATTERNS = ["mobile", "phone", "contact no", "contact number", "tel", "telephone", "cell"]

FUZZY_CUTOFF = 0.78
HEADER_SCAN_ROWS = 30
LOOKAHEAD_ROWS = 4
MIN_LOOKAHEAD_FILL_RATIO = 0.5


def normalize(s):
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


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
    norm = normalize(raw_header)
    if not norm:
        return None
    if is_phone_header(norm):
        return "PHONE"
    if norm in ALIAS_LOOKUP:
        return ALIAS_LOOKUP[norm]
    for alias, canon in ALIAS_LOOKUP.items():
        if alias in norm or norm in alias:
            return canon
    close = difflib.get_close_matches(norm, ALL_ALIASES, n=1, cutoff=FUZZY_CUTOFF)
    if close:
        return ALIAS_LOOKUP[close[0]]
    return None


def _row_score(row):
    return sum(1 for cell in row if map_header(cell))


def _looks_like_data(row, header_score):
    """A row below the candidate header should be mostly filled in, and it
    should NOT itself look like another label/summary row (i.e. it
    shouldn't match aliases almost as often as the header did)."""
    non_empty = sum(1 for v in row if str(v).strip() != "")
    if non_empty == 0:
        return False
    fill_ratio = non_empty / max(len(row), 1)
    if header_score > 0 and _row_score(row) >= header_score * 0.6:
        return False
    return fill_ratio >= MIN_LOOKAHEAD_FILL_RATIO


def find_header_row(raw_df):
    """Scan the first HEADER_SCAN_ROWS rows for the best header candidate.
    A candidate only "qualifies" if the rows immediately beneath it look
    like real tabular data rather than more summary/label text -- this
    stops title-block / summary-section rows (e.g. "Unique Developers:")
    from being mistaken for the real header when they score similarly on
    alias matches. Qualifying candidates are preferred outright; among
    qualifying (or, failing that, all) candidates we prefer more good
    data rows beneath them, then higher score, then the LATER row on ties
    -- summary blocks tend to sit above the real table, not below it.
    """
    n = min(HEADER_SCAN_ROWS, len(raw_df))
    candidates = []
    for i in range(n):
        score = _row_score(raw_df.iloc[i])
        if score > 0:
            candidates.append((i, score))

    if not candidates:
        return 0, 0

    best_key = None
    best_row = candidates[0][0]
    best_score = candidates[0][1]
    for i, score in candidates:
        lookahead = raw_df.iloc[i + 1:i + 1 + LOOKAHEAD_ROWS]
        good_rows = sum(1 for _, r in lookahead.iterrows() if _looks_like_data(r, score))
        qualifies = len(lookahead) > 0 and good_rows >= min(2, len(lookahead))
        key = (qualifies, good_rows, score, i)
        if best_key is None or key > best_key:
            best_key = key
            best_row = i
            best_score = score

    return best_row, best_score


def read_source_file(path):
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
    data = data[~(data.apply(lambda r: all(str(v).strip() == "" for v in r), axis=1))]
    return data


def map_file_to_canonical(df, filename):
    phone_cols_in_order = []
    canon_col_for = {}
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
                unmapped.append(str(col))

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

    out = out[~(out.apply(lambda r: all(str(v).strip() == "" for v in r), axis=1))]
    return out, unmapped


def diagnose_folder(incoming_dir):
    """Utility: report which header row each source file resolved to and
    how many data rows it yielded, without writing an output file. Handy
    for spot-checking a batch after changing header detection."""
    import glob as _glob
    for p in sorted(_glob.glob(f"{incoming_dir}/*.xlsx") + _glob.glob(f"{incoming_dir}/*.xls") + _glob.glob(f"{incoming_dir}/*.csv")):
        try:
            df = read_source_file(p)
            out, unmapped = map_file_to_canonical(df, p)
            print(f"{os.path.basename(p):40s} headers={list(df.columns)[:5]}... rows={len(out)} unmapped={len(unmapped)}")
        except Exception as e:
            print(f"{os.path.basename(p):40s} FAILED: {e}")


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

        try:
            audit_dir = os.path.join(os.path.dirname(abs_output), "batch_logs")
            os.makedirs(audit_dir, exist_ok=True)
            audit_path = os.path.join(audit_dir, f"batch_{run_number}_detail.json")
            with open(audit_path, "w") as f:
                json.dump(result, f, indent=2)
        except Exception:
            pass

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
