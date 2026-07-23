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
import re

from party_config import identify_party, is_po_subject
from graph_email_fetcher import fetch_new_emails
from party_extractors import extract_fields, extract_subject_fields
from existing_parsers_bridge import pick_best_attachment
from sheets_writer import append_po_rows_batch, get_existing_message_ids, get_existing_po_quantities

LOOKBACK_DAYS = 3


def run():
    since = (datetime.now(timezone.utc) - timedelta(days=LOOKBACK_DAYS)).strftime(
        "%Y-%m-%dT%H:%M:%SZ"
    )

    print(f"Fetching emails since {since}...")
    emails = fetch_new_emails(since)
    print(f"Found {len(emails)} emails with PDF attachments.")

    existing_ids = get_existing_message_ids()
    existing_po_with_qty = get_existing_po_quantities()  # (party, po_number) pairs already complete

    summary = {
        "total_emails_seen": len(emails),
        "fetched": 0,
        "success": 0,
        "failed": 0,
        "skipped_duplicate": 0,
        "unmatched_party": 0,
        "no_pdf": 0,
        "ignored": 0,
        "not_a_po": 0,
        "duplicate_po": 0,
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
                    "sender_address": email["sender_address"],
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
                    "sender_address": email["sender_address"],
                    "extractor_used": "none",
                    "status": "FAILED - no party match",
                    "error": f"sender={email['sender_address']}",
                    "message_id": email["message_id"],
                }
            )
            continue

        if not is_po_subject(party, email["subject"]):
            summary["not_a_po"] += 1
            rows_to_write.append(
                {
                    "fetched_at": fetched_at,
                    "party": party,
                    "po_number": "",
                    "po_date": "",
                    "po_quantity": "",
                    "email_subject": email["subject"],
                    "sender_address": email["sender_address"],
                    "extractor_used": "none",
                    "status": "IGNORED - not a PO notification",
                    "error": "subject didn't match known PO pattern for this party",
                    "message_id": email["message_id"],
                }
            )
            continue

        # Zomato only ever CCs a copy of Blink's PO_BCPL thread - so ANY
        # reply from Zomato is redundant with data already captured via
        # Blink or the original Zomato email, regardless of whether this
        # particular reply happens to carry an attachment.
        is_reply_subject = bool(re.search(r"\bre:", email["subject"], re.IGNORECASE))
        if party == "Zomato" and is_reply_subject:
            summary["ignored"] += 1
            rows_to_write.append(
                {
                    "fetched_at": fetched_at,
                    "party": party,
                    "po_number": "",
                    "po_date": "",
                    "po_quantity": "",
                    "email_subject": email["subject"],
                    "sender_address": email["sender_address"],
                    "extractor_used": "none",
                    "status": "IGNORED - reply thread (Zomato CC)",
                    "error": "duplicate of a PO already captured via Blink/original email",
                    "message_id": email["message_id"],
                }
            )
            continue

        if not email["extractable_attachments"]:
            is_reply = bool(re.search(r"\bre:", email["subject"], re.IGNORECASE))

            if is_reply:
                summary["ignored"] += 1
                rows_to_write.append(
                    {
                        "fetched_at": fetched_at,
                        "party": party,
                        "po_number": "",
                        "po_date": "",
                        "po_quantity": "",
                        "email_subject": email["subject"],
                        "sender_address": email["sender_address"],
                        "extractor_used": "none",
                        "status": "IGNORED - reply thread, no attachment",
                        "error": "likely a reply to a PO already captured separately",
                        "message_id": email["message_id"],
                    }
                )
                continue

            subject_fields = extract_subject_fields(party, email["subject"])

            if subject_fields["po_number"] and subject_fields["po_date"]:
                # Got PO Number + Date from the subject itself - real data,
                # just missing quantity since there's no attachment here.
                summary["fetched"] += 1
                summary["failed"] += 1
                rows_to_write.append(
                    {
                        "fetched_at": fetched_at,
                        "party": party,
                        "po_number": subject_fields["po_number"],
                        "po_date": subject_fields["po_date"],
                        "po_quantity": "",
                        "email_subject": email["subject"],
                        "sender_address": email["sender_address"],
                        "extractor_used": "subject-line",
                        "status": "NEEDS REVIEW - quantity not found (no attachment on this email)",
                        "error": "",
                        "message_id": email["message_id"],
                    }
                )
                continue

            summary["no_pdf"] += 1
            other_names = email.get("other_attachment_names", [])

            if other_names:
                detail = f"had non-extractable attachment(s): {', '.join(other_names)}"
            else:
                detail = "no attachments at all"
            rows_to_write.append(
                {
                    "fetched_at": fetched_at,
                    "party": party,
                    "po_number": "",
                    "po_date": "",
                    "po_quantity": "",
                    "email_subject": email["subject"],
                    "sender_address": email["sender_address"],
                    "extractor_used": "none",
                    "status": "NEEDS REVIEW - no PDF/Excel attachment",
                    "error": detail,
                    "message_id": email["message_id"],
                }
            )
            continue

        attachment = pick_best_attachment(party, email["extractable_attachments"])
        summary["fetched"] += 1
        fields = extract_fields(party, attachment["content_bytes"], attachment["file_type"])

        # Fill gaps from the subject line if the attachment extraction
        # missed PO Number or PO Date (Quantity has no subject fallback).
        if not fields["po_number"] or not fields["po_date"]:
            subject_fields = extract_subject_fields(party, email["subject"])
            fields["po_number"] = fields["po_number"] or subject_fields["po_number"]
            fields["po_date"] = fields["po_date"] or subject_fields["po_date"]

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
                "sender_address": email["sender_address"],
                "extractor_used": fields["extractor_used"],
                "status": status,
                "error": fields["error"],
                "message_id": email["message_id"],
            }
        )

    # Build the set of (party, po_number) pairs that DO have quantity data,
    # combining what's already in the sheet with what this batch just found -
    # so a duplicate/reminder email (no attachment) for the same PO gets
    # recognized as redundant instead of flagged "needs review" every run.
    complete_po_pairs = set(existing_po_with_qty)
    for r in rows_to_write:
        if r["party"] and r["po_number"] and r["po_quantity"]:
            complete_po_pairs.add((r["party"], r["po_number"]))

    for r in rows_to_write:
        if (
            r["status"] == "NEEDS REVIEW - quantity not found (no attachment on this email)"
            and (r["party"], r["po_number"]) in complete_po_pairs
        ):
            r["status"] = "IGNORED - duplicate, PO already captured with quantity elsewhere"
            r["error"] = ""
            summary["failed"] -= 1
            summary["duplicate_po"] += 1

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
    print(f"Ignored (not a PO email) : {summary['not_a_po']}")
    print(f"Ignored (duplicate PO)   : {summary['duplicate_po']}")
    print(f"Skipped (already logged) : {summary['skipped_duplicate']}")
    print(
        "\nReconciliation check: every email above is now a row in the sheet "
        "(fetched, no-PDF, unmatched, or already logged) - the counts should "
        "add up to the total emails in window."
    )

    return summary


if __name__ == "__main__":
    run()
