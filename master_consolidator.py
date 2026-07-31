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
    "Unit Number": ["unit", "unit no", "unit number", "villa number", "villa no",
                     "property number", "property no", "land number", "land no",
                     "plot number", "plot no", "no of unit", "no of units", "unit id",
                     "unitnumber", "flat number", "flat", "unit name"],
    "Mobile Number": ["phone", "mobile", "number", "contact number", "phone number",
                        "contact no", "contact", "mobile no", "mobile number",
                        "tel", "telephone", "phone no", "primary phone",
                        "phone mobile", "mobile 1", "primary mobile number",
                        "poa mobile no."],
    "Secondary Mobile Number": ["secondary mobile", "secondary phone", "alternate number",
                                  "alt number", "alternative number", "second contact",
                                  "other number", "mobile 2", "mobile no.3",
                                  "mobile phone3", "mobile 3", "poa phone no."],
    "Telephone Number": ["telephone number", "telephone residence", "telephone office",
                          "phone 1", "phone 2", "phone no.3", "general"],
    "Email Address": ["email", "e-mail", "email address", "e mail", "email add"],
    "Community": ["community", "community name", "master location", "sub community"],
    "Developer": ["developer", "project developer"],
    "Building / Tower": ["building", "tower", "building name", "building 1",
                          "buildingname 2", "buildingnameen", "tower name",
                          "bldg.", "bldg. no."],
    "Project": ["project", "project name", "master project", "emaar project",
                "master project land", "project lnd", "sub project"],
    "Phase": ["phase", "project phase"],
    "Bedrooms": ["bhk", "no bhk", "bedrooms", "no of bedrooms", "bed", "beds",
                 "rooms", "rooms description", "flat typology"],
    "Serial No": ["serial no", "serial number", "sno", "sr no"],
    "Nationality": ["nationality", "nation"],
    "Emirates ID Number": ["idnumber", "uaeidnumber", "emirates id number"],
    "Passport Number": ["passport"],
    "Date of Birth": ["birthdate", "dob"],
    "Gender": ["gender"],
}

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
ENRICHMENT_COLUMNS = ["Mobile Number (Normalized)", "Email Valid", "Possible Duplicate Of"]


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


class SchemaRegistry:
    """Tracks the growing canonical column list and the mapping from every
    raw header seen so far -> canonical column. New, never-seen headers
    become new columns automatically (schema auto-extension)."""

    def __init__(self, existing_columns=None):
        self.columns = list(existing_columns) if existing_columns else []
        self._norm_to_canonical = {}
        for canon, synonyms in HEADER_SYNONYMS.items():
            for syn in synonyms:
                self._norm_to_canonical[syn] = canon
            self._norm_to_canonical[normalize_header(canon)] = canon
        for col in self.columns:
            self._norm_to_canonical.setdefault(normalize_header(col), col)

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
# Precompute the set of every known synonym (normalized) so header detection
# can lean on real domain vocabulary rather than a generic text/number guess.
# This is the strong signal: an actual header cell like "Owner`s Name" or
# "Contact No." will normalize to something in this set almost every time.
_KNOWN_HEADER_NORMS = set()
for _canon, _syns in HEADER_SYNONYMS.items():
    _KNOWN_HEADER_NORMS.add(normalize_header(_canon))
    for _s in _syns:
        _KNOWN_HEADER_NORMS.add(normalize_header(_s))


_EMAIL_LOOSE_RE = re.compile(r"[^@\s]+@[^@\s]+\.[^@\s]+")

# Sheets whose NAME matches one of these (case-insensitive substring) are
# never real owner-contact data -- internal notes, instructions, README
# tabs, dashboards, etc. Skipped entirely regardless of content.
SKIP_SHEET_NAME_SUBSTRINGS = ["instruction", "readme", "notes", "summary", "dashboard"]

# A row containing any of these (normalized) is a pivot-table artifact
# ("Row Labels", "Grand Total") -- if found in the first few rows of a
# sheet, the whole sheet is a pivot table export, not owner records, and
# gets skipped. This is what let country totals like "200"/"18701" and
# country names end up looking like real "Name"/count fields earlier.
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
    """True if a cell's shape screams 'this is a data value', e.g. a phone
    number, an email address, a short numeric code/ID, a unit/serial code
    (letters+digits+hyphens), or a full personal name (two-plus alphabetic
    words) -- as opposed to a short header label."""
    s = s.strip()
    if _EMAIL_LOOSE_RE.search(s):
        return True  # contains an email address -- never a header
    digits_only = re.sub(r"[^\d]", "", s)
    if digits_only and len(digits_only) >= 4 and len(digits_only) >= len(s) - 3:
        return True  # phone-number- or numeric-code-shaped (e.g. "53", "493996",
                      # "971 569346555") -- lowered from a 7-digit floor because
                      # short numeric IDs/codes were previously slipping through
                      # and getting miscounted as header-like labels
    if re.match(r"^[A-Za-z0-9]+(-[A-Za-z0-9]+){2,}$", s):
        return True  # code-shaped, e.g. SC-YN7-CON-CR48-101
    words = s.split()
    if len(words) >= 2 and all(w.isalpha() for w in words):
        return True  # full-name-shaped, e.g. "SHAIKHA ALMARZOOQI" / "Ahmed Ali"
    return False


def looks_like_header(row):
    """A real header row either contains a cell matching known field
    vocabulary (Name, Mobile Number, Email, etc. and their synonyms -- the
    strong signal), or, failing that, is mostly short labels that don't look
    like phone numbers, codes, emails, or personal names (the fallback
    signal). Guards against sheets whose first row is already data, which
    would otherwise be misread as a header and silently dropped -- and,
    just as importantly, guards against an actual DATA row (someone's real
    name/email/phone number) being misread as a header, which would turn
    private data into permanent column names across the whole database."""
    non_empty = [str(c).strip() for c in row if c is not None and str(c).strip() != ""]
    if len(non_empty) < 2:
        return False
    # Hard veto: a row containing an email address is never a real header,
    # regardless of how many other cells look label-like.
    if any(_EMAIL_LOOSE_RE.search(c) for c in non_empty):
        return False
    # Hard veto: a row containing a bare numeric value (any length -- "53",
    # "200", "493996") is never a real header either. Real header labels are
    # words/phrases; a naked number in a header row means this is actually a
    # data row (an ID, unit count, code, etc.) that slipped past the other
    # checks. This is what let single names like "Amjad" paired with a short
    # code like "200" get misread as a header on some sheets.
    if any(re.sub(r"[\s]", "", c).isdigit() for c in non_empty):
        return False
    if any(normalize_header(c) in _KNOWN_HEADER_NORMS for c in non_empty):
        return True
    label_like = sum(1 for c in non_empty if not _looks_like_data_value(c))
    return label_like >= max(1, len(non_empty) * 0.6)


def find_header_row(rows, max_scan=10):
    """Scan the first few rows for one that looks like a real header. Returns
    None if no row looks header-like -- caller must treat the sheet as
    headerless rather than risk dropping real data as a fake header."""
    for i, row in enumerate(rows[:max_scan]):
        if looks_like_header(row):
            return i
    return None


_BEDROOM_LIKE_RE = re.compile(r"\b(bedroom|studio|\bbr\b|bhk)\b", re.IGNORECASE)


def infer_generic_headers(sample_rows, width):
    """For headerless sheets: lightweight type inference per column so data
    still lands in a sensible bucket instead of being lost under 'Column N'."""
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
        # Codes like "PRK1/2E416F" -- allow "/" alongside "-" since real unit
        # codes in this data use both separators. Spaces are stripped first
        # because a handful of source rows have stray spaces mid-code
        # (e.g. "PRKI /SD119/2F558F").
        code_like = sum(1 for v in vals
                          if re.match(r"^[A-Za-z0-9/\-]{4,}$", str(v).strip().replace(" ", ""))
                          and any(ch.isdigit() for ch in str(v)))
        # "Four Bedroom", "Studio", "2 BHK" etc. are unit-type descriptions,
        # not personal names -- checked BEFORE the name check below, since a
        # phrase like "Four Bedroom" is two alphabetic words and would
        # otherwise be indistinguishable from a real name like "Ahmed Ali"
        # by word-count/shape alone. This was misclassifying an entire
        # headerless sheet's bedroom-count column as "Name".
        bedroom_like = sum(1 for v in vals if isinstance(v, str) and _BEDROOM_LIKE_RE.search(v))
        alpha_multiword = sum(1 for v in vals
                               if isinstance(v, str) and len(v.split()) >= 2
                               and v.replace(" ", "").isalpha())
        if email_like / n > 0.6:
            headers[col] = "Email Address"
        elif digit_like / n > 0.6:
            headers[col] = "Mobile Number"
        elif bedroom_like / n > 0.6:
            headers[col] = "Bedrooms"
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
def normalize_uae_phone(raw):
    """Best-effort UAE phone normalization for a NEW column -- never touches
    the original value. Returns None if it doesn't look like a phone at all."""
    if raw is None or str(raw).strip() == "":
        return None
    digits = re.sub(r"[^\d]", "", str(raw))
    if not digits:
        return None
    if digits.startswith("00971"):
        digits = digits[2:]
    if digits.startswith("971"):
        return "+" + digits
    if digits.startswith("0"):
        return "+971" + digits[1:]
    if len(digits) == 9:  # local number missing leading 0
        return "+971" + digits
    if len(digits) >= 8:
        return "+" + digits
    return None


EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def is_valid_email(raw):
    if raw is None or str(raw).strip() == "":
        return None
    return bool(EMAIL_RE.match(str(raw).strip()))


def enrich_record(rec):
    """Adds enrichment columns in place without ever altering original fields."""
    mobile = rec.get("Mobile Number")
    if mobile is not None:
        norm = normalize_uae_phone(mobile)
        if norm:
            rec["Mobile Number (Normalized)"] = norm
    email = rec.get("Email Address")
    if email is not None:
        rec["Email Valid"] = is_valid_email(email)


def detect_duplicates(records):
    """Cross-file duplicate detection by normalized phone or email. Flags
    (does not merge/delete) so nothing is ever lost -- just surfaced for
    manual review. Populates 'Possible Duplicate Of' with the row number(s)
    of earlier matching records."""
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
            rec["Possible Duplicate Of"] = ", ".join(str(m + 2) for m in sorted(matches))  # +2: header row + 1-index


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
    """Shared by read_sheet and read_csv: given a sheet's non-empty rows,
    decide whether row 0 is a real header (find_header_row) or whether the
    sheet is headerless, and return (header, data) either way. Kept as one
    function so the xlsx and csv code paths can never drift out of sync on
    this logic again. Returns None for a genuinely empty sheet, and the
    string "PIVOT" if the sheet is a pivot-table export rather than real
    owner records (caller is responsible for logging/skipping that case)."""
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
    """Returns (header_row, data_rows) for one worksheet, or None if the
    sheet is completely empty."""
    rows = []
    for row in ws.iter_rows(values_only=True):
        rows.append(row)
        if len(rows) > 200000:  # safety cap per sheet; raise for very large sheets
            break
    non_empty_rows = [r for r in rows if any(c is not None and str(c).strip() != "" for c in r)]
    return _split_header_and_data(non_empty_rows)


def read_csv(path):
    with open(path, newline="", encoding="utf-8-sig", errors="replace") as f:
        rows = list(csv.reader(f))
    non_empty_rows = [r for r in rows if any(c.strip() for c in r)]
    return _split_header_and_data(non_empty_rows)


def read_workbook(path, registry, log_entries, processed_registry):
    """Reads every worksheet (including hidden) in one workbook. Returns list
    of normalized record dicts. Never raises -- errors are logged per sheet
    so one bad sheet never stops the run."""
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
            col_map = [registry.canonical_for(h) for h in header]

            imported = 0
            for i, row in enumerate(data, start=1):
                rec = {}
                for h_idx, canon in enumerate(col_map):
                    if canon is None:
                        continue
                    val = row[h_idx] if h_idx < len(row) else None
                    if val is not None and str(val).strip() != "":
                        rec[canon] = val
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
            continue  # one bad sheet never stops the run

    return records


# ============================================================================
# MASTER DATABASE READ/WRITE
# ============================================================================
def _is_corrupted_column_name(name):
    """True if a column NAME (not a cell value) looks like it was actually
    someone's data that got mistaken for a header in an earlier, buggier run
    -- e.g. "Amjad", "200", "ghaith.albezreh@yahoo.com". Applies the exact
    same shape rules as looks_like_header's hard vetoes, one column name at a
    time. This lets load_existing_master self-heal a Master Database that
    still has old corrupted columns in it, instead of silently carrying that
    corruption forward on every future run (which is what happened here --
    the header-detection fix stopped NEW corruption but did nothing about
    corruption already saved in the file from before the fix existed)."""
    s = str(name).strip()
    if _EMAIL_LOOSE_RE.search(s):
        return True
    if re.sub(r"[\s]", "", s).isdigit():
        return True
    return False


def load_existing_master(path):
    """Returns (real_data_columns, records). Meta/enrichment columns are
    deliberately excluded from the returned column list -- they're always
    computed fresh and re-appended by write_master, so feeding them back
    into the SchemaRegistry would pollute the real schema and inflate the
    column count a little more on every single rerun. Columns whose NAME
    itself looks corrupted (see _is_corrupted_column_name) are quarantined
    the same way: dropped from the schema going forward, though the
    underlying row data is untouched -- nothing is deleted, the bad column
    just stops being propagated into every future rebuild."""
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


def write_master(path, columns, records):
    rest_meta = [c for c in META_COLUMNS + ENRICHMENT_COLUMNS if c not in columns and c != "Record ID"]
    ordered_cols = ["Record ID"] + columns + rest_meta
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet("Master")
    ws.append(ordered_cols)
    for rec in records:
        ws.append([rec.get(c, None) for c in ordered_cols])
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
            processed_registry = set()  # corrupt registry -> safest is to start fresh, not crash

    registry = SchemaRegistry(existing_columns=existing_cols)
    log_entries = []
    all_new_records = []

    if not source_dir.exists():
        raise FileNotFoundError(f"Source directory does not exist: {source_dir}")

    # Record ID is a stable, append-only synthetic primary key -- existing
    # rows KEEP whatever ID they were already assigned (never renumbered);
    # new rows continue the sequence from the current highest ID. This is
    # what makes it safe to use "Record ID" to key against this Master
    # Database from other systems/future imports across repeated runs.
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
    # Columns that are 100% empty across every current row (e.g. a leftover
    # "Column1" from a past run whose one populating row is gone) add no
    # value and just clutter the sheet -- drop them from this write. Nothing
    # is lost: if a future import ever uses that exact header text again,
    # canonical_for() simply re-creates the column at that point.
    used_columns = [c for c in registry.columns
                     if any(r.get(c) not in (None, "") for r in combined)]
    write_master(master_path, used_columns, combined)
    write_log(log_path, log_entries)
    registry_path.parent.mkdir(parents=True, exist_ok=True)
    registry_path.write_text(json.dumps(sorted(processed_registry), indent=2))

    dup_count = sum(1 for r in combined if r.get("Possible Duplicate Of"))
    summary = {
        "status": "ok",
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
        sys.exit(1 if summary["sheets_errored"] > 0 else 0)

    except Exception as e:
        print(json.dumps({"status": "fatal_error", "error": str(e)}))
        sys.exit(2)


if __name__ == "__main__":
    main()
