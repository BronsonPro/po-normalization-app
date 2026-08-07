"""
Fetches emails (with PDF attachments) from a Microsoft 365 mailbox using the
Microsoft Graph API, via app-only (client credentials) OAuth2 authentication.

Requires an Azure app registration with:
  - Application permission: Mail.Read (admin-consented)
  - A client secret

Set these via environment variables (do NOT hardcode / commit them):
  AZURE_TENANT_ID
  AZURE_CLIENT_ID
  AZURE_CLIENT_SECRET
  MAILBOX_ADDRESS      -> the inbox to read, e.g. po@orbinter.com
"""

import os
import base64
import requests
import msal
import zipfile
import io

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def get_access_token():
    tenant_id = os.environ["AZURE_TENANT_ID"]
    client_id = os.environ["AZURE_CLIENT_ID"]
    client_secret = os.environ["AZURE_CLIENT_SECRET"]

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.ConfidentialClientApplication(
        client_id, authority=authority, client_credential=client_secret
    )

    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )

    if "access_token" not in result:
        raise RuntimeError(
            f"Failed to get Graph API token: {result.get('error_description', result)}"
        )
    return result["access_token"]


def fetch_new_emails(since_iso_timestamp: str):
    """
    Returns a list of dicts, one per email received after since_iso_timestamp -
    EVERY email in the window is included, even if it has no attachments or
    no extractable one, so nothing can be silently skipped:
        {
            "message_id": str,
            "subject": str,
            "sender_address": str,
            "received_at": str,
            "extractable_attachments": [
                {"filename": str, "content_bytes": bytes, "file_type": "pdf"|"xlsx"|"xls"},
                ...
            ],
            "other_attachment_names": [str, ...]
            # extractable_attachments covers both PDF and Excel POs (some
            # parties like Big Basket send Excel instead of PDF).
            # other_attachment_names lists anything else found (images, docs,
            # csv, etc.) purely for visibility in the status log.
        }

    since_iso_timestamp example: "2026-07-07T00:00:00Z"
    """
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}
    mailbox = os.environ["MAILBOX_ADDRESS"]

    url = (
        f"{GRAPH_BASE}/users/{mailbox}/mailFolders/inbox/messages"
        f"?$filter=receivedDateTime ge {since_iso_timestamp}"
        f"&$select=id,subject,from,receivedDateTime,hasAttachments"
        f"&$top=50"
    )

    results = []

    while url:
        resp = requests.get(url, headers=headers)
        resp.raise_for_status()
        data = resp.json()

        for msg in data.get("value", []):
            sender_address = (
                msg.get("from", {})
                .get("emailAddress", {})
                .get("address", "")
            )
            subject = msg.get("subject", "")
            message_id = msg["id"]
            received_at = msg.get("receivedDateTime", "")

            extractable_attachments = []
            other_attachment_names = []
            if msg.get("hasAttachments"):
                extractable_attachments, other_attachment_names = _fetch_attachments(
                    headers, mailbox, message_id
                )

            results.append(
                {
                    "message_id": message_id,
                    "subject": subject,
                    "sender_address": sender_address,
                    "received_at": received_at,
                    "extractable_attachments": extractable_attachments,
                    "other_attachment_names": other_attachment_names,
                }
            )

        # Handle pagination
        url = data.get("@odata.nextLink")

    return results


def _classify_zip_contents(zip_bytes: bytes, name_l: str) -> str:
    """
    Determines whether a zip attachment holds PDFs or Excel/CSV files by
    actually looking inside it - not just guessing from the filename.
    (Health & Glow's zip, for example, is named
    "103776-BNO104-ORB_INTERNATIONAL__JEVA_.zip" - no hint of "pdf" in the
    name at all, even though it contains PDFs - so filename-only guessing
    was picking the wrong attachment.) Falls back to filename hints only if
    the zip can't be opened/inspected for some reason.
    """
    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            names = [n.lower() for n in zf.namelist()]
            if any(n.endswith(".pdf") for n in names):
                return "zip_pdf"
            if any(n.endswith(".xlsx") or n.endswith(".xls") for n in names):
                return "zip_excel"
            return "zip"
    except Exception:
        if "pdf" in name_l:
            return "zip_pdf"
        elif "excel" in name_l or "xls" in name_l:
            return "zip_excel"
        return "zip"


def _fetch_attachments(headers, mailbox, message_id):
    url = f"{GRAPH_BASE}/users/{mailbox}/messages/{message_id}/attachments"
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    data = resp.json()

    extractable = []
    other_names = []
    for att in data.get("value", []):
        name = att.get("name", "")
        name_l = name.lower()
        content_type = att.get("contentType", "").lower()

        if name_l.endswith(".pdf") or "pdf" in content_type:
            file_type = "pdf"
        elif name_l.endswith(".xlsx"):
            file_type = "xlsx"
        elif name_l.endswith(".xls"):
            file_type = "xls"
        elif name_l.endswith(".csv"):
            file_type = "csv"
        elif name_l.endswith(".zip"):
            file_type = "zip"  # refined below, once we have the actual bytes
        else:
            file_type = None

        if file_type:
            content_bytes_b64 = att.get("contentBytes")
            if content_bytes_b64:
                content_bytes = base64.b64decode(content_bytes_b64)

                if file_type == "zip":
                    file_type = _classify_zip_contents(content_bytes, name_l)

                extractable.append(
                    {
                        "filename": name,
                        "content_bytes": content_bytes,
                        "file_type": file_type,
                    }
                )
            else:
                other_names.append(name or content_type or "unnamed attachment")
        else:
            other_names.append(name or content_type or "unnamed attachment")

    return extractable, other_names
