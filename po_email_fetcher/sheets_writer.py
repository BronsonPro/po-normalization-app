"""
Writes PO rows to a Google Sheet using a service account.

One-time setup (2 min, no admin approval needed):
  1. Go to console.cloud.google.com -> create/select a project.
  2. Enable "Google Sheets API".
  3. Create a Service Account (IAM & Admin -> Service Accounts).
  4. Create a JSON key for it and download it.
  5. Open your target Google Sheet -> click Share -> paste the service
     account's email (looks like xxx@xxx.iam.gserviceaccount.com) -> give
     it Editor access.
  6. Set the environment variable GOOGLE_SERVICE_ACCOUNT_JSON to the full
     contents of that JSON key file (as a string), and GOOGLE_SHEET_ID to
     the sheet's ID (the long string in the sheet's URL).

Sheet columns (row 1 headers, in this exact order):
  Fetched At | Party | PO Number | PO Date | PO Quantity | Email Subject |
  Extractor Used | Status | Error | Message ID

IMPORTANT: the worksheet connection is cached at module level (_ws_cache)
so a run only opens/authenticates the sheet ONCE, no matter how many PO
rows it writes. Writing many rows one-call-per-row was hitting Google's
"60 read requests per minute" quota on large batches - rows are now
collected and sent with a single append_rows() batch call instead.
"""

import os
import json
import gspread
from google.oauth2.service_account import Credentials

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

HEADERS = [
    "Fetched At",
    "Party",
    "PO Number",
    "PO Date",
    "PO Quantity",
    "Email Subject",
    "Extractor Used",
    "Status",
    "Error",
    "Message ID",
]

_ws_cache = None  # cached worksheet handle, reused for the life of the process


def _get_worksheet():
    global _ws_cache
    if _ws_cache is not None:
        return _ws_cache

    creds_json = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"])
    creds = Credentials.from_service_account_info(creds_json, scopes=SCOPES)
    client = gspread.authorize(creds)

    sheet_id = os.environ["GOOGLE_SHEET_ID"]
    sh = client.open_by_key(sheet_id)
    ws = sh.sheet1

    # Ensure headers exist (only checked once per run now)
    first_row = ws.row_values(1)
    if first_row != HEADERS:
        ws.update("A1", [HEADERS])

    _ws_cache = ws
    return ws


def get_existing_message_ids() -> set:
    """Used for dedup - never process the same email twice."""
    ws = _get_worksheet()
    all_values = ws.get_all_values()
    if len(all_values) <= 1:
        return set()
    # Message ID is the last column
    return {row[-1] for row in all_values[1:] if row}


def append_po_rows_batch(rows: list):
    """
    Writes many rows in a single API call instead of one call per row.
    `rows` is a list of dicts, each with the same keys as append_po_row's
    arguments used to take individually (fetched_at, party, po_number, ...).
    """
    if not rows:
        return

    ws = _get_worksheet()
    values = [
        [
            r["fetched_at"],
            r["party"],
            r["po_number"],
            r["po_date"],
            r["po_quantity"],
            r["email_subject"],
            r["extractor_used"],
            r["status"],
            r["error"],
            r["message_id"],
        ]
        for r in rows
    ]

    # Chunk writes to stay well under Sheets API limits on very large batches
    CHUNK_SIZE = 200
    for i in range(0, len(values), CHUNK_SIZE):
        ws.append_rows(values[i : i + CHUNK_SIZE], value_input_option="RAW")
