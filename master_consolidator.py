#!/usr/bin/env python3
"""
master_consolidator.py

Consolidates a folder of messy source spreadsheets (xlsx/xls/xlsm/csv) with
inconsistent column headers into a single canonical-schema workbook.

Improvements over the original version:
  * Token-based header matching instead of naive substring matching, which
    previously caused false positives (e.g. a header normalizing to "no"
    could match almost any alias containing "no" as a substring).
  * Scored conflict resolution: when multiple source columns map to the same
    canonical field, the best match wins (based on match tier + distance)
    instead of "whichever column appeared first in the sheet".
  * Structured logging (via the `logging` module) instead of bare `print`
    calls, so diagnostic output never contaminates the JSON result printed
    on stdout.
  * Type hints and docstrings throughout for maintainability.
  * More detailed audit trail: unmapped columns now include *why* they were
    unmapped (no match vs. lost a conflict vs. overflow phone column).

Usage:
    python master_consolidator.py <incoming_dir> <batch_output_path> <run_number> [--verbose]
"""
from __future__ import annotations

import sys
import os
import re
import json
import glob
import logging
import difflib
import argparse
import traceback
from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional

import pandas as pd

# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #

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

# Match-quality tiers, used to resolve conflicts when two source columns
# want to claim the same canonical field. Lower number = better match.
TIER_EXACT = 0      # normalized header == alias, exactly
TIER_TOKEN = 1       # token-set match (all words of the shorter phrase present)
TIER_FUZZY = 2       # difflib close-match fallback

logger = logging.getLogger("master_consolidator")


# --------------------------------------------------------------------------- #
# Text normalization & alias index
# --------------------------------------------------------------------------- #

def normalize(s: Optional[str]) -> str:
    """Lowercase, strip punctuation, collapse whitespace."""
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


ALIAS_LOOKUP: dict[str, str] = {}
ALL_ALIASES: list[str] = []
for canon, aliases in ALIASES.items():
    for a in aliases:
        na = normalize(a)
        ALIAS_LOOKUP[na] = canon
        ALL_ALIASES.append(na)


def is_phone_header(norm_header: str) -> bool:
    tokens = set(norm_header.split())
    return any(
        p in norm_header if " " in p else p in tokens
        for p in PHONE_PATTERNS
    )


@dataclass
class HeaderMatch:
    canonical: Optional[str]   # canonical field name, or "PHONE", or None
    tier: int = 99             # lower is better
    distance: float = 1.0      # 0 = perfect, used to break ties within a tier


def _token_match(norm: str, alias: str) -> bool:
    """
    True if `alias`'s words are a subset of `norm`'s words (or vice versa),
    which is a much safer notion of "contains" than raw substring matching.
    This is what previously caused false positives like a header
    normalizing to 'no' matching almost any alias containing 'no'.
    """
    norm_tokens = set(norm.split())
    alias_tokens = set(alias.split())
    if not norm_tokens or not alias_tokens:
        return False
    return alias_tokens.issubset(norm_tokens) or norm_tokens.issubset(alias_tokens)


def map_header(raw_header: str) -> HeaderMatch:
    """
    Map a raw source header to a canonical field (or the sentinel "PHONE"
    for any phone-like column, resolved to Mobile 1/2/3 slots later).

    Matching proceeds in tiers, best first:
      1. Exact normalized match against a known alias.
      2. Token-set match (safer than substring matching).
      3. Fuzzy match via difflib, as a last resort.
    """
    norm = normalize(raw_header)
    if not norm:
        return HeaderMatch(None)

    if is_phone_header(norm):
        return HeaderMatch("PHONE", tier=TIER_EXACT, distance=0.0)

    if norm in ALIAS_LOOKUP:
        return HeaderMatch(ALIAS_LOOKUP[norm], tier=TIER_EXACT, distance=0.0)

    best: Optional[HeaderMatch] = None
    for alias, canon in ALIAS_LOOKUP.items():
        if _token_match(norm, alias):
            # Prefer the alias closest in length to the header (fewer
            # "extra" words = a tighter, more confident match).
            distance = abs(len(norm) - len(alias)) / max(len(norm), len(alias), 1)
            candidate = HeaderMatch(canon, tier=TIER_TOKEN, distance=distance)
            if best is None or candidate.distance < best.distance:
                best = candidate
    if best is not None:
        return best

    close = difflib.get_close_matches(norm, ALL_ALIASES, n=1, cutoff=FUZZY_CUTOFF)
    if close:
        ratio = difflib.SequenceMatcher(None, norm, close[0]).ratio()
        return HeaderMatch(ALIAS_LOOKUP[close[0]], tier=TIER_FUZZY, distance=1.0 - ratio)

    return HeaderMatch(None)


# --------------------------------------------------------------------------- #
# Header-row detection
# --------------------------------------------------------------------------- #

def _row_score(row) -> int:
    return sum(1 for cell in row if map_header(cell).canonical)


def _looks_like_data(row, header_score: int) -> bool:
    """
    A row below the candidate header should be mostly filled in, and it
    should NOT itself look like another label/summary row (i.e. it
    shouldn't match aliases almost as often as the header did).
    """
    non_empty = sum(1 for v in row if str(v).strip() != "")
    if non_empty == 0:
        return False
    fill_ratio = non_empty / max(len(row), 1)
    if header_score > 0 and _row_score(row) >= header_score * 0.6:
        return False
    return fill_ratio >= MIN_LOOKAHEAD_FILL_RATIO


def find_header_row(raw_df: pd.DataFrame) -> tuple[int, int]:
    """
    Scan the first HEADER_SCAN_ROWS rows for the best header candidate.

    A candidate only "qualifies" if the rows immediately beneath it look
    like real tabular data rather than more summary/label text -- this
    stops title-block / summary-section rows (e.g. "Unique Developers:")
    from being mistaken for the real header when they score similarly on
    alias matches. Qualifying candidates are preferred outright; among
    qualifying (or, failing that, all) candidates we prefer more good
    data rows beneath them, then higher score, then the LATER row on ties
    -- summary blocks tend to sit above the real table, not below it.

    Returns (header_row_index, score). If no candidate scores above zero,
    returns (0, 0) so the caller can raise a clear error.
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

    logger.debug(
        "header row candidates=%s -> chosen row=%d score=%d",
        candidates, best_row, best_score,
    )
    return best_row, best_score


# --------------------------------------------------------------------------- #
# File reading
# --------------------------------------------------------------------------- #

def read_source_file(path: str) -> pd.DataFrame:
    """Load a source file and slice it down to header row + data rows."""
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


# --------------------------------------------------------------------------- #
# Canonical mapping
# --------------------------------------------------------------------------- #

@dataclass
class MappingResult:
    frame: pd.DataFrame
    unmapped: list[dict] = field(default_factory=list)  # [{"column": ..., "reason": ...}]


def map_file_to_canonical(df: pd.DataFrame, filename: str) -> MappingResult:
    """
    Map a source dataframe's columns onto CANONICAL_FIELDS.

    Conflict handling: if multiple columns match the same canonical field,
    the best-scoring match (by tier, then distance) wins; the rest are
    recorded as unmapped with an explicit "lost to a better match" reason
    rather than being silently dropped.
    """
    phone_cols_in_order: list[str] = []
    # canon -> (column_name, HeaderMatch) for the current best claimant
    best_for: dict[str, tuple[str, HeaderMatch]] = {}
    unmapped: list[dict] = []

    for col in df.columns:
        col_str = str(col).strip()
        match = map_header(col)

        if match.canonical is None:
            if col_str != "":
                unmapped.append({"column": col_str, "reason": "no alias match found"})
            continue

        if match.canonical == "PHONE":
            phone_cols_in_order.append(col)
            continue

        canon = match.canonical
        if canon not in best_for:
            best_for[canon] = (col, match)
        else:
            existing_col, existing_match = best_for[canon]
            if (match.tier, match.distance) < (existing_match.tier, existing_match.distance):
                # New column is a stronger match; the old one loses.
                unmapped.append({
                    "column": str(existing_col).strip(),
                    "reason": f"lost to a stronger match for '{canon}' ({col_str})",
                })
                best_for[canon] = (col, match)
            else:
                unmapped.append({
                    "column": col_str,
                    "reason": f"'{canon}' already claimed by a stronger match ({existing_col})",
                })

    canon_col_for = {canon: col for canon, (col, _match) in best_for.items()}

    mobile_slots = ["Mobile 1", "Mobile 2", "Mobile 3"]
    for i, col in enumerate(phone_cols_in_order):
        if i < 3:
            canon_col_for[mobile_slots[i]] = col
        else:
            unmapped.append({
                "column": str(col).strip(),
                "reason": "more than 3 phone-like columns found; only first 3 kept",
            })

    out = pd.DataFrame()
    for field_name in CANONICAL_FIELDS:
        if field_name in canon_col_for:
            out[field_name] = df[canon_col_for[field_name]].astype(str).str.strip()
        else:
            out[field_name] = ""

    out = out[~(out.apply(lambda r: all(str(v).strip() == "" for v in r), axis=1))]
    logger.debug("%s: mapped columns=%s unmapped=%s", filename, canon_col_for, unmapped)
    return MappingResult(frame=out, unmapped=unmapped)


# --------------------------------------------------------------------------- #
# Diagnostics helper (unchanged behavior, now uses logging)
# --------------------------------------------------------------------------- #

def diagnose_folder(incoming_dir: str) -> None:
    """
    Utility: report which header row each source file resolved to and
    how many data rows it yielded, without writing an output file. Handy
    for spot-checking a batch after changing header detection.
    """
    for p in sorted(
        glob.glob(f"{incoming_dir}/*.xlsx")
        + glob.glob(f"{incoming_dir}/*.xls")
        + glob.glob(f"{incoming_dir}/*.csv")
    ):
        try:
            df = read_source_file(p)
            result = map_file_to_canonical(df, p)
            logger.info(
                "%-40s headers=%s... rows=%d unmapped=%d",
                os.path.basename(p), list(df.columns)[:5], len(result.frame), len(result.unmapped),
            )
        except Exception as e:
            logger.warning("%-40s FAILED: %s", os.path.basename(p), e)


# --------------------------------------------------------------------------- #
# Main entry point
# --------------------------------------------------------------------------- #

def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Consolidate messy source spreadsheets into a single canonical-schema workbook."
    )
    parser.add_argument("incoming_dir", help="folder containing source .xlsx/.xls/.xlsm/.csv files")
    parser.add_argument("batch_output_path", help="path to write the consolidated .xlsx output")
    parser.add_argument("run_number", help="identifier for this run, used in the audit log filename")
    parser.add_argument("-v", "--verbose", action="store_true", help="enable debug logging to stderr")
    return parser


def main() -> None:
    # Preserve the original positional-args contract for backward compatibility,
    # while adding an optional --verbose flag.
    if len(sys.argv) < 4:
        print(json.dumps({
            "status": "fatal_error",
            "error": "usage: master_consolidator.py <incoming_dir> <batch_output_path> <run_number> [--verbose]",
        }))
        sys.exit(1)

    args = build_arg_parser().parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.WARNING,
        format="%(asctime)s %(levelname)s %(message)s",
        stream=sys.stderr,
    )

    incoming_dir = args.incoming_dir
    batch_output_path = args.batch_output_path
    run_number = args.run_number

    try:
        if not os.path.isdir(incoming_dir):
            raise FileNotFoundError(f"incoming_dir not found: {incoming_dir}")

        patterns = ["*.xlsx", "*.xls", "*.xlsm", "*.csv"]
        files: list[str] = []
        for p in patterns:
            files.extend(glob.glob(os.path.join(incoming_dir, p)))
        files = sorted(set(files))

        if not files:
            raise FileNotFoundError(f"no source files found in {incoming_dir}")

        frames: list[pd.DataFrame] = []
        failed_files: list[dict] = []
        unmapped_columns: dict[str, list[dict]] = {}
        files_succeeded = 0

        for path in files:
            fname = os.path.basename(path)
            try:
                raw_df = read_source_file(path)
                result = map_file_to_canonical(raw_df, fname)
                if result.unmapped:
                    unmapped_columns[fname] = result.unmapped
                frames.append(result.frame)
                files_succeeded += 1
                logger.info("%s: OK, %d rows", fname, len(result.frame))
            except Exception as e:
                failed_files.append({"file": fname, "error": str(e)})
                logger.warning("%s: FAILED (%s)", fname, e)
                continue

        combined = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=CANONICAL_FIELDS)

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

        result_payload = {
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
                json.dump(result_payload, f, indent=2)
        except Exception:
            logger.warning("could not write audit log", exc_info=True)

        print(json.dumps(result_payload))
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
