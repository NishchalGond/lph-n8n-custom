#!/usr/bin/env python3
"""
master_consolidator.py
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
from datetime import datetime, UTC
from typing import Optional

import pandas as pd

CANONICAL_FIELDS = [
    "Name", "Community", "Sub-Community", "Building/Cluster", "Unit Number",
    "Size", "Plot Reg. No", "Plot Number", "DMNO", "DMsubno", "Bedroom",
    "Type (Buyer/Seller)", "Mobile 1", "Mobile 2", "Mobile 3",
    "Email Address", "PI number", "Nationality", "Property Type", "Date",
    "Procedure Value", "Developer", "Project",
]

ALIASES = {
    "Name": [
        "name", "full name", "owner name", "contact name", "client name", "customer name",
        "owner", "client", "name of owner", "client full name", "primary applicant name",
        "first name", "joint acct name", "account name", "nameen",
    ],
    "Community": [
        "community", "community name",
        "master location",
    ],
    "Sub-Community": [
        "sub community", "subcommunity", "sub comm",
        "sub-community",
    ],
    "Building/Cluster": [
        "building", "cluster", "building cluster", "tower", "tower name", "building name",
        "bldg", "bldg no", "building 1", "buildingname 2", "buildingnameen", "building/cluster",
    ],
    "Unit Number": [
        "unit no", "unit number", "unit num", "unit", "apt no", "apartment no", "flat no",
        "villa number", "villa no", "property number", "property no", "no of unit",
        "no of units", "unit id", "unitnumber", "flat number", "flat", "unit name",
    ],
    "Size": [
        "size", "area", "sq ft", "sqft", "area sqft", "size sqft",
        "actual size", "unit size", "actual area",
    ],
    "Plot Reg. No": [
        "plot reg no", "plot registration no", "plot reg number", "plot registration number",
        "reg no", "registration number", "regis",
    ],
    "Plot Number": [
        "plot no", "plot number", "plot num",
        "land number", "land no", "landnumber", "plotno",
    ],
    "DMNO": [
        "dmno", "dm no", "dm number",
        "municipality number", "municipality no",
    ],
    "DMsubno": [
        "dmsubno", "dm subno", "dm sub no", "dm sub number",
        "municipality sub no", "municipality subno",
    ],
    "Bedroom": [
        "bedroom", "bed room", "bedrooms", "beds", "br", "no of bedrooms",
        "bhk", "no bhk", "bed", "rooms", "rooms description",
    ],
    "Type (Buyer/Seller)": [
        "type buyer seller", "buyer seller", "type", "role", "buyer/seller",
        "transaction type", "party type",
    ],
    "Email Address": [
        "email address", "email", "e mail", "email id",
        "e-mail", "email add",
    ],
    "PI number": [
        "pi number", "pi no", "pi num",
        "pino",
    ],
    "Nationality": [
        "nationality", "nation",
    ],
    "Property Type": [
        "property type", "prop type", "unit type",
        "sub type", "flat typology",
    ],
    "Date": [
        "date", "transaction date", "reg date", "registration date",
        "procedure date", "date of transaction",
    ],
    "Procedure Value": [
        "procedure value", "value", "amount", "price", "transaction value",
        "procedurevalue", "transaction amount",
    ],
    "Developer": [
        "developer", "developer name",
        "project developer",
    ],
    "Project": [
        "project", "project name",
        "master project", "emaar project", "master project land", "project lnd", "sub project",
    ],
    "Name_ids": [],
}
del ALIASES["Name_ids"]

PHONE_PATTERNS = [
    "mobile", "phone", "contact no", "contact number", "tel", "telephone", "cell",
    "mobile no", "mobile number", "phone no", "primary phone", "phone mobile",
    "primary mobile number", "secondary mobile", "secondary phone", "alternate number",
    "alt number", "alternative number", "second contact", "other number",
    "telephone number", "telephone residence", "telephone office", "general",
]

FUZZY_CUTOFF = 0.72
HEADER_SCAN_ROWS = 30
LOOKAHEAD_ROWS = 4
MIN_LOOKAHEAD_FILL_RATIO = 0.5

TIER_EXACT = 0
TIER_TOKEN = 1
TIER_FUZZY = 2

logger = logging.getLogger("master_consolidator")


def normalize(s: Optional[str]) -> str:
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
    canonical: Optional[str]
    tier: int = 99
    distance: float = 1.0


def _token_match(norm: str, alias: str) -> bool:
    norm_tokens = set(norm.split())
    alias_tokens = set(alias.split())
    if not norm_tokens or not alias_tokens:
        return False
    return alias_tokens.issubset(norm_tokens) or norm_tokens.issubset(alias_tokens)


def map_header(raw_header: str) -> HeaderMatch:
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


def _row_score(row) -> int:
    return sum(1 for cell in row if map_header(cell).canonical)


def _looks_like_data(row, header_score: int) -> bool:
    non_empty = sum(1 for v in row if str(v).strip() != "")
    if non_empty == 0:
        return False
    fill_ratio = non_empty / max(len(row), 1)
    if header_score > 0 and _row_score(row) >= header_score * 0.6:
        return False
    return fill_ratio >= MIN_LOOKAHEAD_FILL_RATIO


def find_header_row(raw_df: pd.DataFrame) -> tuple[int, int]:
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


def _dedupe_headers(headers: list[str]) -> list[str]:
    """
    Make column labels unique the same way pandas would if it had parsed
    the header row itself (via header=N). We build `data.columns` by hand
    below instead of letting pandas do it, so duplicate raw headers (e.g.
    two columns both literally named "PROPERTY TYPE" in a messy source
    file) come through as true duplicate labels. That's a landmine: a
    duplicate label makes `df[that_label]` return a DataFrame instead of
    a Series, which crashes `.astype(str).str.strip()` later and takes the
    ENTIRE file down with it (caught by the per-file try/except, so it's
    silent unless you're reading failed_files in the run's JSON output).
    Deduping here (PROPERTY TYPE, PROPERTY TYPE.1, ...) keeps both columns
    independently selectable, exactly like a normal pandas header read.
    """
    seen: dict[str, int] = {}
    result = []
    for h in headers:
        label = str(h) if h is not None else ""
        if label in seen:
            seen[label] += 1
            result.append(f"{label}.{seen[label]}")
        else:
            seen[label] = 0
            result.append(label)
    return result


def _dedupe_headers(headers: list[str]) -> list[str]:
    """
    Make column labels unique the same way pandas would if it had parsed
    the header row itself (via header=N). We build `data.columns` by hand
    below instead of letting pandas do it, so duplicate raw headers (e.g.
    two columns both literally named "PROPERTY TYPE" in a messy source
    file) come through as true duplicate labels. That's a landmine: a
    duplicate label makes `df[that_label]` return a DataFrame instead of
    a Series, which crashes `.astype(str).str.strip()` later and takes the
    ENTIRE file down with it (caught by the per-file try/except, so it's
    silent unless you're reading failed_files in the run's JSON output).
    Deduping here (PROPERTY TYPE, PROPERTY TYPE.1, ...) keeps both columns
    independently selectable, exactly like a normal pandas header read.
    """
    seen: dict[str, int] = {}
    result = []
    for h in headers:
        label = str(h) if h is not None else ""
        if label in seen:
            seen[label] += 1
            result.append(f"{label}.{seen[label]}")
        else:
            seen[label] = 0
            result.append(label)
    return result


class EmptySheetError(ValueError):
    """Raised for a genuinely empty sheet/file -- not a real failure, just nothing to read."""


def list_source_units(path: str) -> list[tuple[str, pd.DataFrame]]:
    """
    Return a list of (unit_label, raw_dataframe) pairs to process for this
    source file.

    CSV files have no concept of sheets, so they're a single unit.

    Excel workbooks (.xlsx/.xls/.xlsm) are expanded to ONE UNIT PER SHEET.
    Previously this script called pd.read_excel() with no sheet_name arg,
    which silently defaults to sheet 0 only -- any additional sheets in a
    workbook were never read at all, no error, no log entry, just gone.
    For a workbook like an Arabian Ranches file that has both a "Prop"
    sheet (property registry) and a separate "Arabian Ranches Owner" sheet
    (owner/contact registry, different schema entirely), that meant the
    entire second sheet's data was invisible to every run. Each sheet here
    gets its own independent header-row detection and its own column
    mapping, since different sheets can use completely different header
    layouts.
    """
    ext = os.path.splitext(path)[1].lower()
    fname = os.path.basename(path)
    if ext == ".csv":
        raw = pd.read_csv(path, header=None, dtype=str, keep_default_na=False)
        return [(fname, raw)]

    engine = "xlrd" if ext == ".xls" else "openpyxl"
    xls = pd.ExcelFile(path, engine=engine)
    units = []
    multi = len(xls.sheet_names) > 1
    for sheet_name in xls.sheet_names:
        raw = pd.read_excel(xls, sheet_name=sheet_name, header=None, dtype=str, keep_default_na=False)
        label = f"{fname} [{sheet_name}]" if multi else fname
        units.append((label, raw))
    return units


def parse_raw_frame(raw: pd.DataFrame) -> pd.DataFrame:
    """Slice a raw (header=None) dataframe down to header row + data rows."""
    if raw.empty:
        raise EmptySheetError("sheet is empty")

    header_row, score = find_header_row(raw)
    if score <= 0:
        raise ValueError("could not detect a recognizable header row")

    headers = raw.iloc[header_row].tolist()
    data = raw.iloc[header_row + 1:].reset_index(drop=True)
    data.columns = _dedupe_headers(headers)
    data = data[~(data.apply(lambda r: all(str(v).strip() == "" for v in r), axis=1))]
    return data


def read_source_file(path: str) -> pd.DataFrame:
    """
    Back-compat single-frame reader (used by diagnose_folder). Reads only
    the first unit (sheet) of a file -- prefer list_source_units() +
    parse_raw_frame() directly for anything that needs full multi-sheet
    coverage, which is now the main pipeline's default behavior.
    """
    units = list_source_units(path)
    _, raw = units[0]
    return parse_raw_frame(raw)


@dataclass
class MappingResult:
    frame: pd.DataFrame
    unmapped: list[dict] = field(default_factory=list)


def map_file_to_canonical(df: pd.DataFrame, filename: str) -> MappingResult:
    phone_cols_in_order: list[str] = []
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
            col_data = df[canon_col_for[field_name]]
            if isinstance(col_data, pd.DataFrame):
                # Defensive fallback: should no longer happen now that
                # _dedupe_headers() prevents true duplicate labels, but if
                # it ever does, take the first occurrence and log it
                # instead of crashing (a crash here previously took the
                # ENTIRE source file down, not just this one column).
                logger.warning(
                    "%s: column '%s' resolved to %d duplicate columns; using the first",
                    filename, canon_col_for[field_name], col_data.shape[1],
                )
                col_data = col_data.iloc[:, 0]
            out[field_name] = col_data.astype(str).str.strip()
        else:
            out[field_name] = ""

    out = out[~(out.apply(lambda r: all(str(v).strip() == "" for v in r), axis=1))]
    logger.debug("%s: mapped columns=%s unmapped=%s", filename, canon_col_for, unmapped)
    return MappingResult(frame=out, unmapped=unmapped)


def diagnose_folder(incoming_dir: str) -> None:
    for p in sorted(
        glob.glob(f"{incoming_dir}/*.xlsx")
        + glob.glob(f"{incoming_dir}/*.xls")
        + glob.glob(f"{incoming_dir}/*.csv")
    ):
        try:
            units = list_source_units(p)
        except Exception as e:
            logger.warning("%-40s FAILED to open: %s", os.path.basename(p), e)
            continue
        for label, raw in units:
            try:
                df = parse_raw_frame(raw)
                result = map_file_to_canonical(df, label)
                logger.info(
                    "%-50s headers=%s... rows=%d unmapped=%d",
                    label, list(df.columns)[:5], len(result.frame), len(result.unmapped),
                )
            except EmptySheetError:
                logger.info("%-50s empty, skipped", label)
            except Exception as e:
                logger.warning("%-50s FAILED: %s", label, e)


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
                units = list_source_units(path)
            except Exception as e:
                # Workbook wouldn't even open (corrupt file, wrong engine, etc.)
                failed_files.append({"file": fname, "error": f"could not open: {e}"})
                logger.warning("%s: FAILED to open (%s)", fname, e)
                continue

            file_had_success = False
            for label, raw in units:
                try:
                    data = parse_raw_frame(raw)
                except EmptySheetError:
                    # A genuinely empty extra tab (e.g. a stray "Sheet2"
                    # with nothing in it) isn't a failure -- just skip it
                    # quietly instead of cluttering failed_files.
                    logger.info("%s: empty, skipped", label)
                    continue
                except Exception as e:
                    failed_files.append({"file": label, "error": str(e)})
                    logger.warning("%s: FAILED (%s)", label, e)
                    continue

                result = map_file_to_canonical(data, label)
                if result.unmapped:
                    unmapped_columns[label] = result.unmapped
                frames.append(result.frame)
                file_had_success = True
                logger.info("%s: OK, %d rows", label, len(result.frame))

            if file_had_success:
                files_succeeded += 1

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
            "generated_at": datetime.now(UTC).isoformat(),
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
