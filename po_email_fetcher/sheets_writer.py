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


def _get_worksheet():
    creds_json = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"])
    creds = Credentials.from_service_account_info(creds_json, scopes=SCOPES)
    client = gspread.authorize(creds)

    sheet_id = os.environ["GOOGLE_SHEET_ID"]
    sh = client.open_by_key(sheet_id)
    ws = sh.sheet1

    # Ensure headers exist
    first_row = ws.row_values(1)
    if first_row != HEADERS:
        ws.update("A1", [HEADERS])

    return ws


def get_existing_message_ids() -> set:
    """Used for dedup - never process the same email twice."""
    ws = _get_worksheet()
    all_values = ws.get_all_values()
    if len(all_values) <= 1:
        return set()
    # Message ID is the last column
    return {row[-1] for row in all_values[1:] if row}


def append_po_row(
    fetched_at: str,
    party: str,
    po_number: str,
    po_date: str,
    po_quantity: str,
    email_subject: str,
    extractor_used: str,
    status: str,
    error: str,
    message_id: str,
):
    ws = _get_worksheet()
    ws.append_row(
        [
            fetched_at,
            party,
            po_number,
            po_date,
            po_quantity,
            email_subject,
            extractor_used,
            status,
            error,
            message_id,
        ]
    )
