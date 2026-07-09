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
    no PDF, so nothing can be silently skipped:
        {
            "message_id": str,
            "subject": str,
            "sender_address": str,
            "received_at": str,
            "pdf_attachments": [ {"filename": str, "content_bytes": bytes}, ... ]
            # pdf_attachments is [] if the email had no PDF - caller logs this
            # as "No PDF attachment" instead of dropping it.
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

            pdf_attachments = []
            if msg.get("hasAttachments"):
                pdf_attachments = _fetch_pdf_attachments(headers, mailbox, message_id)

            results.append(
                {
                    "message_id": message_id,
                    "subject": subject,
                    "sender_address": sender_address,
                    "received_at": received_at,
                    "pdf_attachments": pdf_attachments,
                }
            )

        # Handle pagination
        url = data.get("@odata.nextLink")

    return results


def _fetch_pdf_attachments(headers, mailbox, message_id):
    url = f"{GRAPH_BASE}/users/{mailbox}/messages/{message_id}/attachments"
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    data = resp.json()

    pdfs = []
    for att in data.get("value", []):
        name = att.get("name", "")
        content_type = att.get("contentType", "")
        if name.lower().endswith(".pdf") or "pdf" in content_type.lower():
            content_bytes_b64 = att.get("contentBytes")
            if content_bytes_b64:
                pdfs.append(
                    {
                        "filename": name,
                        "content_bytes": base64.b64decode(content_bytes_b64),
                    }
                )
    return pdfs
