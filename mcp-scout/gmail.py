# © 2026 Iryna Subbotina. All Rights Reserved. Proprietary — not open source. See LICENSE/NOTICE. Unauthorized copying, redistribution, or claim of authorship is prohibited.
"""Outlook Mail integration (Microsoft Graph) — read and summarize recent emails."""
from __future__ import annotations

import logging
import re
from datetime import datetime, timedelta, timezone

import gauth

log = logging.getLogger(__name__)

_SKIP_SENDERS = re.compile(
    r"(no-reply|noreply|donotreply|notifications?@|mailer@|newsletter@"
    r"|support@|info@|hello@|admin@|bot@|automated@)",
    re.IGNORECASE,
)


def get_recent_emails(
    hours: int = 24,
    max_emails: int = 40,
    skip_automated: bool = True,
) -> list[dict]:
    """Fetch recent emails from the Outlook inbox.

    Args:
      hours: how many hours back to look (default 24)
      max_emails: max emails to return (default 40)
      skip_automated: skip newsletters, no-reply senders (default True)

    Returns: list of {subject, from, date, snippet, body_preview, message_id, thread_id}
    """
    since = (datetime.now(timezone.utc) - timedelta(hours=hours)).strftime("%Y-%m-%dT%H:%M:%SZ")

    result = gauth.request(
        "GET",
        "/me/mailFolders/inbox/messages",
        params={
            "$filter": f"receivedDateTime ge {since}",
            "$orderby": "receivedDateTime desc",
            "$top": max_emails,
            "$select": "id,conversationId,subject,from,receivedDateTime,bodyPreview",
        },
    )

    emails = []
    for msg in result.get("value", []):
        sender_addr = msg.get("from", {}).get("emailAddress", {}).get("address", "")
        sender_name = msg.get("from", {}).get("emailAddress", {}).get("name", "")
        sender = f"{sender_name} <{sender_addr}>" if sender_name else sender_addr
        subject = msg.get("subject") or "(no subject)"
        preview = msg.get("bodyPreview", "")

        if skip_automated and _SKIP_SENDERS.search(sender):
            continue

        emails.append({
            "message_id": msg.get("id", ""),
            "thread_id": msg.get("conversationId", ""),
            "subject": subject,
            "from": sender,
            "date": msg.get("receivedDateTime", ""),
            "snippet": preview[:200],
            "body_preview": preview[:500].strip(),
        })

    log.info("gmail: fetched %d emails (hours=%d)", len(emails), hours)
    return emails
