"""Fetch BCECN PDF attachments from Outlook (Exchange/O365) via OAuth2."""

import os
import tempfile
from datetime import datetime, timedelta
from exchangelib import (
    Account,
    Configuration,
    DELEGATE,
    Identity,
    Q,
)
from exchangelib.protocol import BaseProtocol
from exchangelib.credentials import OAuth2AuthorizationCodeCredentials
import msal


# Microsoft's well-known client ID for "Microsoft Office" public client.
# Works for device-code flow against any O365 tenant without app registration.
# Microsoft Outlook Mobile - commonly pre-approved in enterprise tenants
MS_OFFICE_CLIENT_ID = "27922004-5251-4030-b22d-91ecd9a37ea4"
EWS_SCOPE = ["https://outlook.office365.com/EWS.AccessAsUser.All"]


def _acquire_token_device_code(email: str, client_id: str = MS_OFFICE_CLIENT_ID, status_callback=None):
    """Acquire an OAuth2 access token using the device code flow.

    Args:
        email: User's email address (used to derive tenant domain).
        client_id: Azure AD application (client) ID.
        status_callback: Optional callable(message: str) for status updates
                         (e.g. "Go to https://... and enter code ABC123").

    Returns:
        dict with 'access_token' on success.

    Raises:
        RuntimeError on authentication failure.
    """
    tenant = email.split("@")[1]
    authority = f"https://login.microsoftonline.com/{tenant}"
    app = msal.PublicClientApplication(client_id, authority=authority)

    flow = app.initiate_device_flow(scopes=EWS_SCOPE)
    if "user_code" not in flow:
        raise RuntimeError(f"Could not initiate device flow: {flow.get('error_description', 'unknown error')}")

    message = flow["message"]  # e.g. "To sign in, use a web browser to open ..."
    if status_callback:
        status_callback(message)
    else:
        print(message)

    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        error = result.get("error_description", result.get("error", "unknown error"))
        raise RuntimeError(f"OAuth2 authentication failed: {error}")

    return result


def connect_outlook(email: str, server: str = "outlook.office365.com", status_callback=None) -> Account:
    """Connect to Exchange/O365 mailbox using OAuth2 device code flow.

    Args:
        email: User's corporate email address.
        server: EWS server hostname.
        status_callback: Optional callable(message: str) for auth status updates.

    Returns:
        Connected exchangelib Account.
    """
    token_result = _acquire_token_device_code(email=email, status_callback=status_callback)
    access_token = token_result["access_token"]

    credentials = OAuth2AuthorizationCodeCredentials(
        access_token={"access_token": access_token, "token_type": "Bearer"},
    )

    config = Configuration(server=server, credentials=credentials)
    account = Account(
        primary_smtp_address=email,
        config=config,
        autodiscover=False,
        access_type=DELEGATE,
    )
    return account


def fetch_bcecn_pdfs(
    account: Account,
    sender_filter: str = None,
    days_back: int = 7,
    output_dir: str = None,
) -> list[str]:
    """Search inbox for BCECN / Bell Canada emails and download PDF attachments.

    Args:
        account: Connected Exchange account.
        sender_filter: Optional sender email to filter by.
        days_back: How many days back to search.
        output_dir: Directory to save PDFs. Defaults to temp dir.

    Returns:
        List of file paths to downloaded PDFs.
    """
    if output_dir is None:
        output_dir = tempfile.mkdtemp(prefix="bcecn_")
    os.makedirs(output_dir, exist_ok=True)

    since = datetime.now() - timedelta(days=days_back)

    # Build query: subject contains BCECN or Bell Canada
    subject_filter = Q(subject__contains="BCECN") | Q(subject__contains="Bell Canada")
    query = subject_filter & Q(datetime_received__gte=since)
    if sender_filter:
        query &= Q(sender__contains=sender_filter)

    downloaded = []
    for item in account.inbox.filter(query).order_by("-datetime_received"):
        for attachment in item.attachments:
            if hasattr(attachment, "name") and attachment.name and attachment.name.lower().endswith(".pdf"):
                name_upper = attachment.name.upper()
                if "BCECN" in name_upper or "BELL CANADA" in name_upper:
                    filepath = os.path.join(output_dir, attachment.name)
                    with open(filepath, "wb") as f:
                        f.write(attachment.content)
                    print(f"Downloaded: {attachment.name} (from {item.sender})")
                    downloaded.append(filepath)

    if not downloaded:
        print("No BCECN / Bell Canada PDFs found in inbox for the given criteria.")

    return downloaded
