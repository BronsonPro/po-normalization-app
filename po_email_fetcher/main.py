"""
Run this once a day (manually, or via a button in your Streamlit app).

Flow:
  1. Fetch emails from the last LOOKBACK_DAYS with PDF attachments
     (a lookback window + dedup-by-message-id is used instead of a saved
     "last run" timestamp, so a missed day never causes a gap).
  2. Identify the party from sender + subject.
  3. Extract PO Number / PO Date / PO Quantity from the PDF.
  4. Write ALL rows to the Google Sheet in ONE batch call at the end -
     success or failure both get logged, so nothing is silently missed,
     and a single batch write avoids hitting Google's API rate limits
     on days with a lot of emails.
  5. Print a summary at the end.
"""

from datetime import datetime, timedelta, timezone

from party_config import identify_party
from graph_email_fetcher import fetch_new_emails
from party_extractors import extract_fields
from sheets_writer import append_po_rows_batch, get_existing_message_ids

LOOKBACK_DAYS = 2


def run():
    since = (datetime.now(timezone.utc) - timedelta(days=LOOKBACK_DAYS)).strftime(
        "%Y-%m-%dT%H:%M:%SZ"
    )

    print(f"Fetching emails since {since}...")
    emails = fetch_new_emails(since)
    print(f"Found {len(emails)} emails with PDF attachments.")

    existing_ids = get_existing_message_ids()

    summary = {
        "total_emails_seen": len(emails),
        "fetched": 0,
        "success": 0,
        "failed": 0,
        "skipped_duplicate": 0,
        "unmatched_party": 0,
        "no_pdf": 0,
        "ignored": 0,
    }

    rows_to_write = []  # collected here, written in one batch call at the end

    for email in emails:
        if email["message_id"] in existing_ids:
            summary["skipped_duplicate"] += 1
            continue

        party = identify_party(email["sender_address"], email["subject"])
        fetched_at = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")

        if party == "IGNORE":
            summary["ignored"] += 1
            rows_to_write.append(
                {
                    "fetched_at": fetched_at,
                    "party": "IGNORED",
                    "po_number": "",
                    "po_date": "",
                    "po_quantity": "",
                    "email_subject": email["subject"],
                    "extractor_used": "none",
                    "status": "IGNORED - known non-PO sender",
                    "error": f"sender={email['sender_address']}",
                    "message_id": email["message_id"],
                }
            )
            continue

        if not party:
            summary["unmatched_party"] += 1
            rows_to_write.append(
                {
                    "fetched_at": fetched_at,
                    "party": "UNKNOWN",
                    "po_number": "",
                    "po_date": "",
                    "po_quantity": "",
                    "email_subject": email["subject"],
                    "extractor_used": "none",
                    "status": "FAILED - no party match",
                    "error": f"sender={email['sender_address']}",
                    "message_id": email["message_id"],
                }
            )
            continue

        if not email["pdf_attachments"]:
            summary["no_pdf"] += 1
            rows_to_write.append(
                {
                    "fetched_at": fetched_at,
                    "party": party,
                    "po_number": "",
                    "po_date": "",
                    "po_quantity": "",
                    "email_subject": email["subject"],
                    "extractor_used": "none",
                    "status": "NEEDS REVIEW - no PDF attachment",
                    "error": "",
                    "message_id": email["message_id"],
                }
            )
            continue

        for attachment in email["pdf_attachments"]:
            summary["fetched"] += 1
            fields = extract_fields(party, attachment["content_bytes"])

            got_all_fields = fields["po_number"] and fields["po_date"] and fields["po_quantity"]
            status = "SUCCESS" if got_all_fields and not fields["error"] else "NEEDS REVIEW"

            if status == "SUCCESS":
                summary["success"] += 1
            else:
                summary["failed"] += 1

            rows_to_write.append(
                {
                    "fetched_at": fetched_at,
                    "party": party,
                    "po_number": fields["po_number"],
                    "po_date": fields["po_date"],
                    "po_quantity": fields["po_quantity"],
                    "email_subject": email["subject"],
                    "extractor_used": fields["extractor_used"],
                    "status": status,
                    "error": fields["error"],
                    "message_id": email["message_id"],
                }
            )

    print(f"Writing {len(rows_to_write)} rows to Google Sheet in a single batch...")
    append_po_rows_batch(rows_to_write)

    print("\n--- Run Summary ---")
    print(f"Total emails in window   : {summary['total_emails_seen']}")
    print(f"New PDFs processed       : {summary['fetched']}")
    print(f"Successful               : {summary['success']}")
    print(f"Needs review             : {summary['failed']}")
    print(f"No PDF attachment        : {summary['no_pdf']}")
    print(f"Unmatched sender         : {summary['unmatched_party']}")
    print(f"Ignored (known non-PO)   : {summary['ignored']}")
    print(f"Skipped (already logged) : {summary['skipped_duplicate']}")
    print(
        "\nReconciliation check: every email above is now a row in the sheet "
        "(fetched, no-PDF, unmatched, or already logged) - the counts should "
        "add up to the total emails in window."
    )

    return summary


if __name__ == "__main__":
    run()
