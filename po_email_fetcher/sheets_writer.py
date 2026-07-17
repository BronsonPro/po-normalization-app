"""
Writes PO rows to a Google Sheet using a service account, split across two tabs:
  - "Sheet1" (main tab)  -> SUCCESS and NEEDS REVIEW rows (the ones worth acting on)
  - "Ignored" (second tab) -> IGNORED... and FAILED - no party match rows
Kept separate for a trial period so you can compare/verify before deleting
the ignored rows outright.

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

Sheet columns (row 1 headers, in this exact order, same on both tabs):
  Fetched At | Party | PO Number | PO Date | PO Quantity | Email Subject |
  Sender Address | Extractor Used | Status | Error | Message ID

IMPORTANT: worksheet connections are cached at module level so a run only
opens/authenticates each tab ONCE, no matter how many rows it writes -
writing one-call-per-row was hitting Google's rate limit on large batches.
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
    "Sender Address",
    "Extractor Used",
    "Status",
    "Error",
    "Message ID",
]

IGNORED_TAB_NAME = "Ignored"

_ws_cache = {}       # {"main": worksheet, "ignored": worksheet}
_spreadsheet_cache = None


def _get_spreadsheet():
    global _spreadsheet_cache
    if _spreadsheet_cache is not None:
        return _spreadsheet_cache

    creds_json = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"])
    creds = Credentials.from_service_account_info(creds_json, scopes=SCOPES)
    client = gspread.authorize(creds)

    sheet_id = os.environ["GOOGLE_SHEET_ID"]
    _spreadsheet_cache = client.open_by_key(sheet_id)
    return _spreadsheet_cache


def _ensure_headers(ws):
    first_row = ws.row_values(1)
    if first_row != HEADERS:
        ws.update("A1", [HEADERS])


def _get_main_worksheet():
    if "main" in _ws_cache:
        return _ws_cache["main"]
    sh = _get_spreadsheet()
    ws = sh.sheet1
    _ensure_headers(ws)
    _ws_cache["main"] = ws
    return ws


def _get_ignored_worksheet():
    if "ignored" in _ws_cache:
        return _ws_cache["ignored"]
    sh = _get_spreadsheet()
    try:
        ws = sh.worksheet(IGNORED_TAB_NAME)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=IGNORED_TAB_NAME, rows=1000, cols=len(HEADERS))
    _ensure_headers(ws)
    _ws_cache["ignored"] = ws
    return ws


def _is_ignored_status(status: str) -> bool:
    return status.startswith("IGNORED") or status.startswith("FAILED")


def _all_values_both_tabs():
    main_values = _get_main_worksheet().get_all_values()
    ignored_values = _get_ignored_worksheet().get_all_values()
    main_rows = main_values[1:] if len(main_values) > 1 else []
    ignored_rows = ignored_values[1:] if len(ignored_values) > 1 else []
    return main_rows + ignored_rows


def get_existing_message_ids() -> set:
    """Used for dedup - never process the same email twice. Checks BOTH tabs."""
    all_rows = _all_values_both_tabs()
    return {row[-1] for row in all_rows if row}


def get_existing_po_quantities() -> set:
    """
    Returns a set of (party, po_number) pairs that already have a non-empty
    PO Quantity logged (checked across both tabs). Used so a duplicate/
    reminder email for a PO already captured elsewhere (e.g. Blink often
    sends the same PO_BCPL subject multiple times, only one with the
    attachment) can be recognized as a duplicate instead of flagged
    "needs review" every time.
    """
    all_rows = _all_values_both_tabs()

    party_idx = HEADERS.index("Party")
    po_number_idx = HEADERS.index("PO Number")
    po_qty_idx = HEADERS.index("PO Quantity")

    result = set()
    for row in all_rows:
        if len(row) <= max(party_idx, po_number_idx, po_qty_idx):
            continue
        party = row[party_idx].strip()
        po_number = row[po_number_idx].strip()
        po_qty = row[po_qty_idx].strip()
        if party and po_number and po_qty:
            result.add((party, po_number))
    return result


def _row_to_values(r: dict) -> list:
    return [
        r["fetched_at"],
        r["party"],
        r["po_number"],
        r["po_date"],
        r["po_quantity"],
        r["email_subject"],
        r["sender_address"],
        r["extractor_used"],
        r["status"],
        r["error"],
        r["message_id"],
    ]


def append_po_rows_batch(rows: list):
    """
    Splits rows between the main tab (SUCCESS / NEEDS REVIEW) and the
    Ignored tab (IGNORED... / FAILED - no party match), writing each group
    in a single batch call per tab instead of one call per row.
    """
    if not rows:
        return

    main_values = []
    ignored_values = []
    for r in rows:
        values = _row_to_values(r)
        if _is_ignored_status(r["status"]):
            ignored_values.append(values)
        else:
            main_values.append(values)

    CHUNK_SIZE = 200

    if main_values:
        ws = _get_main_worksheet()
        for i in range(0, len(main_values), CHUNK_SIZE):
            ws.append_rows(main_values[i : i + CHUNK_SIZE], value_input_option="RAW")

    if ignored_values:
        ws = _get_ignored_worksheet()
        for i in range(0, len(ignored_values), CHUNK_SIZE):
            ws.append_rows(ignored_values[i : i + CHUNK_SIZE], value_input_option="RAW")
