#!/usr/bin/env python3
"""
write_master_manifest.py

Renders the running master manifest (a JSON array of row objects,
maintained by the n8n workflow) into a single .xlsx, in the same 5-column
style as LPH's existing File_List.xlsx:

    Name | DirectoryName | Extension | Size (KB) | LastWriteTime

One row = one batch output file (Consolidated_Batch_NNN.xlsx). The n8n
workflow appends a new row to the JSON state after every successful batch
save, then calls this script to re-render the .xlsx before uploading it.

Usage:
    python3 write_master_manifest.py <state_json_path> <manifest_xlsx_path>

Prints a one-line JSON status to stdout. Exit code 1 only on a fatal
read/write error (state file is treated as authoritative; a missing state
file is NOT fatal -- it's just treated as an empty manifest).
"""

import sys
import os
import json
import pandas as pd

MANIFEST_COLUMNS = ["Name", "DirectoryName", "Extension", "Size (KB)", "LastWriteTime"]


def main():
    if len(sys.argv) < 3:
        print(json.dumps({"status": "fatal_error", "error": "usage: write_master_manifest.py <state_json_path> <manifest_xlsx_path>"}))
        sys.exit(1)

    state_path = sys.argv[1]
    manifest_path = sys.argv[2]

    try:
        if os.path.isfile(state_path):
            with open(state_path, "r") as f:
                rows = json.load(f)
            if not isinstance(rows, list):
                raise ValueError("state file does not contain a JSON array")
        else:
            rows = []

        df = pd.DataFrame(rows)
        for col in MANIFEST_COLUMNS:
            if col not in df.columns:
                df[col] = []
        df = df[MANIFEST_COLUMNS]

        os.makedirs(os.path.dirname(os.path.abspath(manifest_path)), exist_ok=True)
        df.to_excel(manifest_path, index=False, sheet_name="File_List")

        print(json.dumps({
            "status": "success",
            "manifest_file": os.path.abspath(manifest_path),
            "row_count": int(len(df)),
        }))
        sys.exit(0)

    except Exception as e:
        print(json.dumps({"status": "fatal_error", "error": str(e)}))
        sys.exit(1)


if __name__ == "__main__":
    main()
