"""Fetch BCECN PDF attachments from Outlook (Exchange/O365)."""

import os
import tempfile
from datetime import datetime, timedelta
from exchangelib import (
    Account,
    Credentials,
    Configuration,
    DELEGATE,
    Q,
)


def connect_outlook(email: str, password: str, server: str = "outlook.office365.com") -> Account:
    """Connect to Exchange/O365 mailbox."""
    credentials = Credentials(email, password)
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
    """Search inbox for BCECN emails and download PDF attachments.

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

    # Build query: subject contains BCECN, has attachments
    query = Q(subject__contains="BCECN") & Q(datetime_received__gte=since)
    if sender_filter:
        query &= Q(sender__contains=sender_filter)

    downloaded = []
    for item in account.inbox.filter(query).order_by("-datetime_received"):
        for attachment in item.attachments:
            if hasattr(attachment, "name") and attachment.name and attachment.name.lower().endswith(".pdf"):
                if "BCECN" in attachment.name.upper():
                    filepath = os.path.join(output_dir, attachment.name)
                    with open(filepath, "wb") as f:
                        f.write(attachment.content)
                    print(f"Downloaded: {attachment.name} (from {item.sender})")
                    downloaded.append(filepath)

    if not downloaded:
        print("No BCECN PDFs found in inbox for the given criteria.")

    return downloaded
