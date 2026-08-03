"""
LPH Master Database Consolidation Engine
=========================================
Reads every worksheet of every xlsx/xlsm/xls/csv file in a source folder,
normalizes inconsistent headers onto one shared schema tailored to LPH's
real-estate owner-contact data, auto-extends that schema when genuinely new
fields appear, and appends everything into one growing Master Database --
with zero data loss, full logging, and non-destructive data enrichment.

Design principles:
  - Every worksheet is read. Only truly empty sheets are skipped.
  - Header normalization is synonym-based and CONFIGURABLE (HEADER_SYNONYMS
    below). Unknown headers are never dropped -- they become new columns.
  - Headerless sheets (first row is already data) are detected and given
    inferred column names instead of losing that row as a fake header.
  - One bad sheet/file never stops the run: errors are caught per-sheet,
    logged, and processing continues.
  - Master Database is append-only across runs: existing rows/columns are
    never deleted, reordered, or overwritten; new columns are appended.
  - Original values are NEVER modified. Enrichment (phone normalization,
    email validation, cross-file duplicate detection) is added as new,
    separate columns -- the raw value stays exactly as supplied.
  - A processed-file registry (content hash + sheet name) prevents
    re-ingesting the same sheet twice across repeated runs.
  - A post-write size/row-count guard flags abnormal master-file growth in
    the run summary so a future runaway-growth bug gets caught in a
    Telegram notification instead of surfacing later as an unopenable
    100MB+ file in Excel.
  - Runs headless from n8n's Execute Command node: takes a JSON config,
    prints exactly one JSON summary line to stdout, and uses exit codes
    (0 = clean, 1 = completed with warnings/errors logged, 2 = fatal) so
    the workflow can branch on success/failure without parsing prose.

Extending later (AI header mapping, fuzzy matching, address normalization,
incremental cloud sync, etc.) is a matter of adding functions in the
PLUGIN POINTS section without touching the core loop.
"""

import argparse
import csv
import hashlib
import json
import re
import sys
import time
import unicodedata
from datetime import datetime, timezone
from pathlib import Path

import openpyxl

# ============================================================================
# CONFIG -- header synonym map, tailored to LPH real-estate owner-contact
# data. Add new synonyms here; no code changes needed elsewhere. Keys are
# canonical column names as they appear in the Master Database.
# ============================================================================
HEADER_SYNONYMS = {
    "Name": ["name", "owner name", "owner", "client name", "owners name",
             "customer name", "name of owner", "client", "full name",
             "first name", "client full name", "primary applicant name",
             "nameen", "joint acct name", "account name"],
    "Community": ["community", "community name", "master location"],
    "Sub-Community": ["sub community", "subcommunity", "sub-community"],
    "Building/Cluster": ["building", "tower", "building name", "building 1",
                          "buildingname 2", "buildingnameen", "tower name",
                          "bldg.", "bldg. no.", "cluster", "building/cluster",
                          "building cluster"],
    "Unit Number": ["unit", "unit no", "unit number", "villa number", "villa no",
                     "property number", "property no", "no of unit", "no of units",
                     "unit id", "unitnumber", "flat number", "flat", "unit name"],
    "Size": ["size", "actual size", "unit size", "area", "actual area"],
    "Plot Reg. No": ["plot reg no", "plot reg. no", "plot registration number",
                       "reg no", "registration number", "regis"],
    "Plot Number": ["plot number", "plot no", "land number", "land no", "landnumber", "plotno"],
    "DMNO": ["dm no", "dmno", "municipality number", "municipality no"],
    "DMsubno": ["dm sub no", "dmsubno", "municipality sub no", "municipality subno",
                 "dm sub number"],
    "Bedroom": ["bhk", "no bhk", "bedrooms", "bedroom", "no of bedrooms", "bed", "beds",
                 "rooms", "rooms description"],
    "Type (Buyer/Seller)": ["type buyer seller", "buyer seller", "buyer/seller",
                              "type (buyer/seller)", "transaction type", "party type"],
    "Email Address": ["email", "e-mail", "email address", "e mail", "email add"],
    "PI number": ["pi number", "pi no", "pino", "pi num"],
    "Nationality": ["nationality", "nation"],
    "Property Type": ["property type", "unit type", "sub type", "flat typology"],
    "Date": ["date", "transaction date", "procedure date", "date of transaction"],
    "Procedure Value": ["procedure value", "procedurevalue", "transaction amount",
                          "transaction value", "amount"],
    "Developer": ["developer", "project developer"],
    "Project": ["project", "project name", "master project", "emaar project",
                "master project land", "project lnd", "sub project"],
    # Kept for zero-data-loss, but not part of the required main list -- these
    # still consolidate their own synonyms into one column each, just ordered
    # after the main headers instead of before them.
    "Serial No": ["serial no", "serial number", "sno", "sr no"],
    "Emirates ID Number": ["idnumber", "uaeidnumber", "emirates id number"],
    "Passport Number": ["passport"],
    "Date of Birth": ["birthdate", "dob"],
    "Gender": ["gender"],
}

# The fixed, required column order for the consolidated output. These
# ALWAYS appear in this exact order, even if a given run's data happens
# not to populate one of them yet (unlike other auto-discovered columns,
# which only appear when at least one row actually has a value).
MAIN_HEADERS = [
    "Name", "Community", "Sub-Community", "Building/Cluster", "Unit Number",
    "Size", "Plot Reg. No", "Plot Number", "DMNO", "DMsubno", "Bedroom",
    "Type (Buyer/Seller)", "Mobile 1", "Mobile 2", "Mobile 3",
    "Email Address", "PI number", "Nationality", "Property Type", "Date",
    "Procedure Value", "Developer", "Project",
]

# Phone-like raw headers are handled separately from the normal
# canonical_for() mapping: instead of every synonym collapsing onto one
# fixed target column, each row's phone-like values are collected in the
# order their columns appear and distributed across Mobile 1 / Mobile 2 /
# Mobile 3 (a 4th+ distinct number gets appended onto Mobile 3 rather than
# lost). This list is every synonym that means "this column is *a* phone
# number", not any specific one of the three slots. Normalized into a set
# further down, once normalize_header() exists.
PHONE_SYNONYMS_RAW = [
    "phone", "mobile", "number", "contact number", "phone number",
    "contact no", "contact", "mobile no", "mobile number", "tel",
    "telephone", "phone no", "primary phone", "phone mobile", "mobile 1",
    "primary mobile number", "poa mobile no.", "secondary mobile",
    "secondary phone", "alternate number", "alt number",
    "alternative number", "second contact", "other number", "mobile 2",
    "poa phone no.", "mobile no.3", "mobile phone3", "mobile 3", "phone 2",
    "phone no.3", "telephone number", "telephone residence",
    "telephone office", "phone 1", "general",
]

# NOTE ON THIS LIST: this is a starting expansion based on one real production
# dataset, not an exhaustive mapping. Columns that still show up as sparse,
# near-duplicate fields after a run (e.g. two columns that clearly mean the
# same real-world thing) should be added here as new synonyms -- that is the
# intended, ongoing maintenance loop for this file. Deliberately NOT merged in
# this pass: fields that look similar but carry different units or meanings
# (e.g. "Built-up area sqm" vs "Built-up area sqft" vs "Plot area sqft" --
# merging these would silently mix square-meter and square-foot values in one
# column) or different real-world concepts (e.g. "Residence Country" is not
# the same fact as "Nationality"). Those are left as separate columns on
# purpose rather than guessed at.

SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xls", ".csv"}

# Meta/enrichment columns are always appended at the very end, after every
# real data column, so real data never gets pushed around by these.
# "Record ID" is the one exception -- write_master pins it as the FIRST
# column, since it's the stable key meant for merging/updating against in
# future runs, not incidental bookkeeping like the others.
META_COLUMNS = ["Record ID", "_source_file", "_source_sheet", "_source_row", "_ingested_at"]
ENRICHMENT_COLUMNS = ["Mobile Number (Normalized)", "Mobile Country", "Mobile Number Valid",
                       "Email Valid", "Possible Duplicate Of"]

# ----------------------------------------------------------------------------
# SAFETY GUARD -- after writing the master file, its size and row count are
# checked against these thresholds. Crossing either one does NOT stop the
# run or touch the data (nothing here ever blocks a write); it only flags
# "size_warning": true and a human-readable reason in the JSON summary, so
# whatever's consuming that summary (the n8n Telegram notify step) surfaces
# it immediately instead of it going unnoticed until someone can't open the
# file in Excel. Raise these if the real dataset legitimately grows past
# them over time -- they're a tripwire for ABNORMAL growth, not a hard cap.
# ============================================================================
MAX_SAFE_MASTER_SIZE_MB = 50
MAX_SAFE_ROW_COUNT = 200_000


# ============================================================================
# HEADER NORMALIZATION
# ============================================================================
def normalize_header(raw):
    """Collapse whitespace/punctuation/case differences so equivalent headers
    compare equal. E.g. "Owner`s Name", "Owner's Name", "  owner  name " all
    normalize to "owners name"."""
    if raw is None:
        return ""
    s = str(raw)
    s = unicodedata.normalize("NFKD", s)
    s = s.replace("\n", " ").replace("\t", " ").replace("\r", " ")
    s = s.lower()
    s = re.sub(r"[`'’‘]", "", s)      # apostrophe/backtick variants -> remove
    s = re.sub(r"[,.]", "", s)        # commas/periods -> remove
    s = re.sub(r"[\-_/]", " ", s)     # hyphen/underscore/slash -> word separator
    s = re.sub(r"\s+", " ", s).strip()
    return s


PHONE_SYNONYMS_NORMALIZED = {normalize_header(s) for s in PHONE_SYNONYMS_RAW}


class SchemaRegistry:
    """Tracks the growing canonical column list and the mapping from every
    raw header seen so far -> canonical column. New, never-seen headers
    become new columns automatically (schema auto-extension)."""

    def __init__(self, existing_columns=None):
        self.columns = list(existing_columns) if existing_columns else []
        self._norm_to_canonical = {}
        for canon, synonyms in HEADER_SYNONYMS.items():
            for syn in synonyms:
                self._norm_to_canonical[normalize_header(syn)] = canon
            self._norm_to_canonical[normalize_header(canon)] = canon
        for col in self.columns:
            self._norm_to_canonical.setdefault(normalize_header(col), col)
        for main_header in MAIN_HEADERS:
            self.ensure_column(main_header)

    def canonical_for(self, raw_header):
        nk = normalize_header(raw_header)
        if not nk:
            return None
        if nk in self._norm_to_canonical:
            canon = self._norm_to_canonical[nk]
        else:
            canon = re.sub(r"\s+", " ", str(raw_header).strip())
            self._norm_to_canonical[nk] = canon
        self.ensure_column(canon)  # ALWAYS register, synonym match or not
        return canon

    def ensure_column(self, canon):
        if canon not in self.columns:
            self.columns.append(canon)


# ============================================================================
# HEADER-VS-DATA DETECTION (headerless sheets)
# ============================================================================
_KNOWN_HEADER_NORMS = set()
for _canon, _syns in HEADER_SYNONYMS.items():
    _KNOWN_HEADER_NORMS.add(normalize_header(_canon))
    for _s in _syns:
        _KNOWN_HEADER_NORMS.add(normalize_header(_s))


_EMAIL_LOOSE_RE = re.compile(r"[^@\s]+@[^@\s]+\.[^@\s]+")

SKIP_SHEET_NAME_SUBSTRINGS = ["instruction", "readme", "notes", "summary", "dashboard"]

_PIVOT_MARKER_NORMS = {"row labels", "column labels", "grand total"}


def _sheet_name_excluded(sheet_name):
    n = str(sheet_name or "").strip().lower()
    return any(sub in n for sub in SKIP_SHEET_NAME_SUBSTRINGS)


def _looks_like_pivot_table(rows, max_scan=5):
    for row in rows[:max_scan]:
        for c in row:
            if c is None:
                continue
            norm = normalize_header(c)
            if norm in _PIVOT_MARKER_NORMS or norm.startswith("count of"):
                return True
    return False


def _looks_like_data_value(s):
    s = s.strip()
    if _EMAIL_LOOSE_RE.search(s):
        return True
    digits_only = re.sub(r"[^\d]", "", s)
    if digits_only and len(digits_only) >= 4 and len(digits_only) >= len(s) - 3:
        return True
    if re.match(r"^[A-Za-z0-9]+(-[A-Za-z0-9]+){2,}$", s):
        return True
    words = s.split()
    if len(words) >= 2 and all(w.isalpha() for w in words):
        return True
    return False


def looks_like_header(row):
    non_empty = [str(c).strip() for c in row if c is not None and str(c).strip() != ""]
    if len(non_empty) < 2:
        return False
    if any(_EMAIL_LOOSE_RE.search(c) for c in non_empty):
        return False
    if any(re.sub(r"[\s]", "", c).isdigit() for c in non_empty):
        return False
    if any(normalize_header(c) in _KNOWN_HEADER_NORMS for c in non_empty):
        return True
    label_like = sum(1 for c in non_empty if not _looks_like_data_value(c))
    return label_like >= max(1, len(non_empty) * 0.6)


def find_header_row(rows, max_scan=10):
    for i, row in enumerate(rows[:max_scan]):
        if looks_like_header(row):
            return i
    return None


_BEDROOM_LIKE_RE = re.compile(r"\b(bedroom|studio|\bbr\b|bhk)\b", re.IGNORECASE)

KNOWN_SUB_COMMUNITY_NAMES = {"richmond", "topanga"}

KNOWN_VIEW_FACING_VALUES = {"front", "back", "side", "corner",
                              "sea view", "garden view", "pool view",
                              "park view", "community view", "boulevard view"}


def infer_generic_headers(sample_rows, width):
    headers = [None] * width
    for col in range(width):
        vals = [r[col] for r in sample_rows[:30]
                 if col < len(r) and r[col] is not None and str(r[col]).strip() != ""]
        if not vals:
            headers[col] = f"Column {col + 1}"
            continue
        n = len(vals)
        digit_like = sum(1 for v in vals
                          if str(v).replace(" ", "").replace("+", "").isdigit()
                          and len(str(v).replace(" ", "")) >= 7)
        email_like = sum(1 for v in vals if _EMAIL_LOOSE_RE.search(str(v)))
        code_like = sum(1 for v in vals
                          if re.match(r"^[A-Za-z0-9/\-]{4,}$", str(v).strip().replace(" ", ""))
                          and any(ch.isdigit() for ch in str(v)))
        bedroom_like = sum(1 for v in vals if isinstance(v, str) and _BEDROOM_LIKE_RE.search(v))
        community_like = sum(1 for v in vals
                               if isinstance(v, str) and v.strip().lower() in KNOWN_SUB_COMMUNITY_NAMES)
        view_like = sum(1 for v in vals
                          if isinstance(v, str) and v.strip().lower() in KNOWN_VIEW_FACING_VALUES)
        alpha_multiword = sum(1 for v in vals
                               if isinstance(v, str) and len(v.split()) >= 2
                               and v.replace(" ", "").isalpha())
        if email_like / n > 0.6:
            headers[col] = "Email Address"
        elif digit_like / n > 0.6:
            headers[col] = "Mobile 1"
        elif bedroom_like / n > 0.6:
            headers[col] = "Bedrooms"
        elif community_like / n > 0.6:
            headers[col] = "Community"
        elif view_like / n > 0.6:
            headers[col] = "View / Facing"
        elif alpha_multiword / n > 0.6:
            headers[col] = "Name"
        elif code_like / n > 0.6:
            headers[col] = "Unit Number"
        else:
            headers[col] = f"Column {col + 1}"
    return headers


# ============================================================================
# NON-DESTRUCTIVE ENRICHMENT (PLUGIN POINTS)
# ============================================================================
import phonenumbers
from phonenumbers import NumberParseException

# Friendly names for the country codes phonenumbers detects, so the sheet
# reads "Saudi Arabia" instead of a bare ISO code "SA". Not exhaustive --
# any code not listed here just falls back to showing the ISO code itself,
# which is still meaningful, just less pretty.
COUNTRY_NAMES = {
    "AE": "UAE", "SA": "Saudi Arabia", "IQ": "Iraq", "IR": "Iran",
    "IN": "India", "PK": "Pakistan", "EG": "Egypt", "JO": "Jordan",
    "LB": "Lebanon", "SY": "Syria", "KW": "Kuwait", "QA": "Qatar",
    "BH": "Bahrain", "OM": "Oman", "YE": "Yemen", "GB": "UK",
    "US": "USA", "CA": "Canada", "AU": "Australia", "FR": "France",
    "DE": "Germany", "PH": "Philippines", "BD": "Bangladesh",
    "NG": "Nigeria", "CN": "China", "RU": "Russia", "TR": "Turkey",
    "ZA": "South Africa", "KE": "Kenya", "MA": "Morocco", "TN": "Tunisia",
    "DZ": "Algeria", "LK": "Sri Lanka", "NP": "Nepal", "AF": "Afghanistan",
    "SD": "Sudan", "ET": "Ethiopia", "MY": "Malaysia", "ID": "Indonesia",
    "IT": "Italy", "ES": "Spain", "NL": "Netherlands", "CH": "Switzerland",
}

_GARBAGE_PHONE_VALUES = {"null", "n/a", "na", "none", "-", "0", "00", "nil"}


def normalize_phone(raw, default_region="AE"):
    """Best-effort international phone parsing for a NEW column -- never
    touches the original value. default_region is only used as a fallback
    guess for numbers with NO country code at all (bare local format); a
    number that already carries a country code (leading + or 00) is parsed
    using that, not assumed to be UAE. Returns (e164_or_None,
    country_name_or_None, is_valid_bool_or_None)."""
    if raw is None:
        return None, None, None
    s = str(raw).strip()
    if not s or s.lower() in _GARBAGE_PHONE_VALUES:
        return None, None, None
    s = s.replace("|", "")  # source data sometimes has stray pipe separators, e.g. "971|55-2600133"
    if s.startswith("00"):
        s = "+" + s[2:]
    try:
        parsed = phonenumbers.parse(s, None if s.startswith("+") else default_region)
    except NumberParseException:
        return None, None, None
    valid = phonenumbers.is_valid_number(parsed)
    possible = phonenumbers.is_possible_number(parsed)
    if not (valid or possible):
        return None, None, None
    e164 = phonenumbers.format_number(parsed, phonenumbers.PhoneNumberFormat.E164)
    region = phonenumbers.region_code_for_number(parsed)
    country_name = COUNTRY_NAMES.get(region, region)
    return e164, country_name, valid


EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def is_valid_email(raw):
    if raw is None or str(raw).strip() == "":
        return None
    return bool(EMAIL_RE.match(str(raw).strip()))


def enrich_record(rec):
    for slot in ("Mobile 1", "Mobile 2", "Mobile 3"):
        raw = rec.get(slot)
        if raw is None:
            continue
        e164, country_name, valid = normalize_phone(raw)
        if e164:
            rec["Mobile Number (Normalized)"] = e164
            rec["Mobile Country"] = country_name
            rec["Mobile Number Valid"] = valid
            break  # first successfully-parsed slot wins for duplicate-detection purposes
    email = rec.get("Email Address")
    if email is not None:
        rec["Email Valid"] = is_valid_email(email)


def detect_duplicates(records):
    seen_phone = {}
    seen_email = {}
    for idx, rec in enumerate(records):
        matches = set()
        norm_phone = rec.get("Mobile Number (Normalized)")
        if norm_phone:
            if norm_phone in seen_phone:
                matches.add(seen_phone[norm_phone])
            else:
                seen_phone[norm_phone] = idx
        email = rec.get("Email Address")
        if email:
            key = str(email).strip().lower()
            if key in seen_email:
                matches.add(seen_email[key])
            else:
                seen_email[key] = idx
        if matches:
            rec["Possible Duplicate Of"] = ", ".join(str(m + 2) for m in sorted(matches))


# ============================================================================
# FILE READING (per-sheet, error-isolated)
# ============================================================================
def sha256_of_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def now_iso():
    return datetime.now(timezone.utc).isoformat()


def _split_header_and_data(non_empty_rows):
    if not non_empty_rows:
        return None
    if _looks_like_pivot_table(non_empty_rows):
        return "PIVOT"
    hdr_idx = find_header_row(non_empty_rows)
    if hdr_idx is None:
        width = max(len(r) for r in non_empty_rows)
        header = infer_generic_headers(non_empty_rows, width)
        data = non_empty_rows
    else:
        header = non_empty_rows[hdr_idx]
        data = non_empty_rows[hdr_idx + 1:]
    return header, data


def read_sheet(ws):
    rows = []
    for row in ws.iter_rows(values_only=True):
        rows.append(row)
        if len(rows) > 200000:
            break
    non_empty_rows = [r for r in rows if any(c is not None and str(c).strip() != "" for c in r)]
    return _split_header_and_data(non_empty_rows)


def read_csv(path):
    with open(path, newline="", encoding="utf-8-sig", errors="replace") as f:
        rows = list(csv.reader(f))
    non_empty_rows = [r for r in rows if any(c.strip() for c in r)]
    return _split_header_and_data(non_empty_rows)


def read_workbook(path, registry, log_entries, processed_registry):
    records = []
    ext = path.suffix.lower()
    try:
        file_hash = sha256_of_file(path)
    except Exception as e:
        log_entries.append({"file": path.name, "sheet": None, "status": "ERROR",
                             "error": f"Could not read file: {e}", "rows_imported": 0,
                             "timestamp": now_iso()})
        return records

    if ext == ".csv":
        sheets = [("Sheet1", None)]
    else:
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        except Exception as e:
            log_entries.append({"file": path.name, "sheet": None, "status": "ERROR",
                                 "error": f"Failed to open workbook: {e}", "rows_imported": 0,
                                 "timestamp": now_iso()})
            return records
        sheets = [(name, wb[name]) for name in wb.sheetnames]

    for sheet_name, ws in sheets:
        t0 = time.time()
        if _sheet_name_excluded(sheet_name):
            log_entries.append({"file": path.name, "sheet": sheet_name,
                                 "status": "SKIPPED_EXCLUDED_SHEET_NAME", "rows_imported": 0,
                                 "timestamp": now_iso(), "duration_sec": round(time.time() - t0, 3)})
            continue
        sheet_key = f"{file_hash}::{sheet_name}"
        if sheet_key in processed_registry:
            log_entries.append({"file": path.name, "sheet": sheet_name,
                                 "status": "SKIPPED_DUPLICATE", "rows_imported": 0,
                                 "timestamp": now_iso(), "duration_sec": round(time.time() - t0, 3)})
            continue
        try:
            result = read_csv(path) if ext == ".csv" else read_sheet(ws)
            if result == "PIVOT":
                log_entries.append({"file": path.name, "sheet": sheet_name,
                                     "status": "SKIPPED_PIVOT_TABLE", "rows_imported": 0,
                                     "timestamp": now_iso(), "duration_sec": round(time.time() - t0, 3)})
                continue
            if result is None:
                log_entries.append({"file": path.name, "sheet": sheet_name,
                                     "status": "EMPTY_SKIPPED", "rows_imported": 0,
                                     "timestamp": now_iso(), "duration_sec": round(time.time() - t0, 3)})
                continue
            header, data = result

            new_cols_before = set(registry.columns)
            col_map = []
            for h in header:
                if normalize_header(h) in PHONE_SYNONYMS_NORMALIZED:
                    col_map.append("__PHONE__")
                else:
                    col_map.append(registry.canonical_for(h))

            PHONE_LIKE_CANONICALS = {"Telephone Number", "Alternate Number(s)"}

            imported = 0
            for i, row in enumerate(data, start=1):
                rec = {}
                phone_vals = []  # collected in column order, deduped, assigned to Mobile 1/2/3 after the loop
                for h_idx, canon in enumerate(col_map):
                    val = row[h_idx] if h_idx < len(row) else None
                    if val is None or str(val).strip() == "":
                        continue
                    val_str = str(val).strip()
                    if canon == "__PHONE__":
                        if val_str not in phone_vals:
                            phone_vals.append(val_str)
                        continue
                    if canon is None:
                        continue
                    if canon in rec and str(rec[canon]).strip() != val_str:
                        if canon in PHONE_LIKE_CANONICALS:
                            existing_parts = [p.strip() for p in str(rec[canon]).split(" / ")]
                            if val_str not in existing_parts:
                                rec[canon] = str(rec[canon]) + " / " + val_str
                        else:
                            suffix = 2
                            alt = f"{canon} ({suffix})"
                            while alt in rec:
                                suffix += 1
                                alt = f"{canon} ({suffix})"
                            rec[alt] = val
                            registry.ensure_column(alt)
                    else:
                        rec[canon] = val
                if phone_vals:
                    if len(phone_vals) >= 1:
                        rec["Mobile 1"] = phone_vals[0]
                    if len(phone_vals) >= 2:
                        rec["Mobile 2"] = phone_vals[1]
                    if len(phone_vals) >= 3:
                        # a 4th+ distinct number is appended onto Mobile 3 rather than dropped
                        rec["Mobile 3"] = " / ".join(phone_vals[2:])
                if not rec:
                    continue
                enrich_record(rec)
                rec["_source_file"] = path.name
                rec["_source_sheet"] = sheet_name
                rec["_source_row"] = i
                rec["_ingested_at"] = now_iso()
                records.append(rec)
                imported += 1

            new_cols = set(registry.columns) - new_cols_before
            processed_registry.add(sheet_key)
            log_entries.append({"file": path.name, "sheet": sheet_name, "status": "OK",
                                 "rows_imported": imported,
                                 "columns_mapped": len([c for c in col_map if c]),
                                 "new_columns_created": sorted(new_cols),
                                 "timestamp": now_iso(), "duration_sec": round(time.time() - t0, 3)})
        except Exception as e:
            log_entries.append({"file": path.name, "sheet": sheet_name, "status": "ERROR",
                                 "error": str(e), "rows_imported": 0, "timestamp": now_iso(),
                                 "duration_sec": round(time.time() - t0, 3)})
            continue

    return records


# ============================================================================
# MASTER DATABASE READ/WRITE
# ============================================================================
def _is_corrupted_column_name(name):
    s = str(name).strip()
    if _EMAIL_LOOSE_RE.search(s):
        return True
    if re.sub(r"[\s]", "", s).isdigit():
        return True
    return False


def load_existing_master(path):
    if not path.exists():
        return [], []
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["Master"]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], []
    header = list(rows[0])
    records = []
    for r in rows[1:]:
        rr = list(r) + [None] * (len(header) - len(r))
        if not any(x is not None and str(x).strip() != "" for x in rr):
            continue
        records.append(dict(zip(header, rr)))
    reserved = set(META_COLUMNS + ENRICHMENT_COLUMNS)
    real_columns = [c for c in header
                     if c not in reserved and not _is_corrupted_column_name(c)]
    return real_columns, records


def _sanitize_sheet_name(name, used_names):
    """Excel sheet names: max 31 chars, no : \\ / ? * [ ] characters, and
    must be unique within the workbook. Truncates and de-dupes as needed."""
    s = re.sub(r'[:\\/?*\[\]]', '-', str(name)).strip()
    if not s:
        s = "Unassigned"
    s = s[:31]
    base = s
    n = 2
    while s.lower() in used_names:
        suffix = f" ({n})"
        s = base[: 31 - len(suffix)] + suffix
        n += 1
    used_names.add(s.lower())
    return s


def _group_key_for(rec):
    """Grouping priority for the per-community sheets: an explicit Community
    value, then Project, then the source file name (cleaned up) -- since
    most rows in this dataset only carry a reliable project identity via
    which file they came from, not a populated Community/Project column."""
    for field in ("Community", "Project"):
        val = rec.get(field)
        if val is not None and str(val).strip() != "":
            return str(val).strip()
    src = rec.get("_source_file")
    if src:
        return re.sub(r"\.(xlsx|xlsm|xls|csv)$", "", str(src), flags=re.IGNORECASE).replace("_", " ").strip()
    return "Unassigned"


def write_master(path, columns, records, group_sheets=True):
    extra_cols = [c for c in columns if c not in MAIN_HEADERS]
    rest_meta = [c for c in META_COLUMNS + ENRICHMENT_COLUMNS
                 if c not in MAIN_HEADERS and c not in extra_cols and c != "Record ID"]
    ordered_cols = ["Record ID"] + MAIN_HEADERS + extra_cols + rest_meta
    wb = openpyxl.Workbook(write_only=True)

    ws = wb.create_sheet("Master")
    ws.append(ordered_cols)
    for rec in records:
        ws.append([rec.get(c, None) for c in ordered_cols])

    if group_sheets:
        groups = {}
        for rec in records:
            key = _group_key_for(rec)
            groups.setdefault(key, []).append(rec)
        used_names = {"master"}
        for key in sorted(groups.keys(), key=lambda k: (-len(groups[k]), k.lower())):
            sheet_name = _sanitize_sheet_name(key, used_names)
            gws = wb.create_sheet(sheet_name)
            gws.append(ordered_cols)
            for rec in groups[key]:
                gws.append([rec.get(c, None) for c in ordered_cols])

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def write_log(path, log_entries):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w") as f:
        json.dump(log_entries, f, indent=2, default=str)


# ============================================================================
# MAIN ENTRY POINT
# ============================================================================
def consolidate(source_dir, master_path, registry_path, log_path):
    source_dir = Path(source_dir)
    master_path = Path(master_path)
    registry_path = Path(registry_path)
    log_path = Path(log_path)

    existing_cols, existing_records = load_existing_master(master_path)
    processed_registry = set()
    if registry_path.exists():
        try:
            processed_registry = set(json.loads(registry_path.read_text()))
        except Exception:
            processed_registry = set()

    registry = SchemaRegistry(existing_columns=existing_cols)
    log_entries = []
    all_new_records = []

    if not source_dir.exists():
        raise FileNotFoundError(f"Source directory does not exist: {source_dir}")

    next_id = 1
    for r in existing_records:
        try:
            rid = int(r.get("Record ID") or 0)
        except (TypeError, ValueError):
            rid = 0
        if rid >= next_id:
            next_id = rid + 1

    files = sorted(p for p in source_dir.rglob("*")
                    if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS)

    for path in files:
        recs = read_workbook(path, registry, log_entries, processed_registry)
        for rec in recs:
            rec["Record ID"] = next_id
            next_id += 1
        all_new_records.extend(recs)

    combined = existing_records + all_new_records
    detect_duplicates(combined)
    used_columns = [c for c in registry.columns
                     if any(r.get(c) not in (None, "") for r in combined)]
    write_master(master_path, used_columns, combined)
    write_log(log_path, log_entries)
    registry_path.parent.mkdir(parents=True, exist_ok=True)
    registry_path.write_text(json.dumps(sorted(processed_registry), indent=2))

    # --------------------------------------------------------------------
    # SAFETY GUARD: check what actually landed on disk, not what we think
    # we wrote. Runs every time, costs nothing, never blocks a write --
    # only makes abnormal growth visible instead of silent.
    # --------------------------------------------------------------------
    try:
        master_file_size_mb = round(master_path.stat().st_size / (1024 * 1024), 2)
    except OSError:
        master_file_size_mb = None
    size_warning = bool(
        (master_file_size_mb is not None and master_file_size_mb > MAX_SAFE_MASTER_SIZE_MB)
        or len(combined) > MAX_SAFE_ROW_COUNT
    )

    dup_count = sum(1 for r in combined if r.get("Possible Duplicate Of"))
    summary = {
        "status": "warning_large_master" if size_warning else "ok",
        "files_scanned": len(files),
        "new_rows_imported": len(all_new_records),
        "total_rows_in_master": len(combined),
        "total_columns": len(used_columns),
        "possible_duplicates_flagged": dup_count,
        "sheets_ok": sum(1 for l in log_entries if l["status"] == "OK"),
        "sheets_empty_skipped": sum(1 for l in log_entries if l["status"] == "EMPTY_SKIPPED"),
        "sheets_duplicate_skipped": sum(1 for l in log_entries if l["status"] == "SKIPPED_DUPLICATE"),
        "sheets_errored": sum(1 for l in log_entries if l["status"] == "ERROR"),
        "errors": [l for l in log_entries if l["status"] == "ERROR"],
        "master_path": str(master_path),
        "log_path": str(log_path),
        "master_file_size_mb": master_file_size_mb,
        "size_warning": size_warning,
    }
    return summary


def load_config(config_path):
    cfg = json.loads(Path(config_path).read_text())
    required = ["source_dir", "master_path", "registry_path", "log_path"]
    missing = [k for k in required if k not in cfg]
    if missing:
        raise ValueError(f"Config missing required keys: {missing}")
    return cfg


def main():
    parser = argparse.ArgumentParser(description="LPH Master Database Consolidation Engine")
    parser.add_argument("--config", help="Path to JSON config file (source_dir, master_path, "
                                          "registry_path, log_path)")
    parser.add_argument("source_dir", nargs="?", help="Folder of source spreadsheets")
    parser.add_argument("master_path", nargs="?", help="Path to Master_Database.xlsx")
    parser.add_argument("registry_path", nargs="?", help="Path to processed_registry.json")
    parser.add_argument("log_path", nargs="?", help="Path to import_log.json")
    args = parser.parse_args()

    try:
        if args.config:
            cfg = load_config(args.config)
            source_dir, master_path = cfg["source_dir"], cfg["master_path"]
            registry_path, log_path = cfg["registry_path"], cfg["log_path"]
        elif args.source_dir and args.master_path:
            source_dir, master_path = args.source_dir, args.master_path
            registry_path = args.registry_path or str(Path(master_path).with_suffix(".registry.json"))
            log_path = args.log_path or str(Path(master_path).with_suffix(".import_log.json"))
        else:
            print(json.dumps({"status": "fatal_error",
                               "error": "Provide --config <file> OR source_dir + master_path args"}))
            sys.exit(2)

        summary = consolidate(source_dir, master_path, registry_path, log_path)
        print(json.dumps(summary))
        sys.exit(1 if (summary["sheets_errored"] > 0 or summary["size_warning"]) else 0)

    except Exception as e:
        print(json.dumps({"status": "fatal_error", "error": str(e)}))
        sys.exit(2)


if __name__ == "__main__":
    main()
