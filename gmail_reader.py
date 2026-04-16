"""
Gmail Reader - Read emails from ceslep@gmail.com via Gmail API (OAuth2).

Setup:
  1. Go to https://console.cloud.google.com/
  2. Create project (or select existing)
  3. Enable "Gmail API"
  4. Create OAuth 2.0 credentials (Desktop App)
  5. Download JSON → save as 'credentials.json' in this folder
  6. pip install -r requirements.txt
  7. python gmail_reader.py
"""

import os
import base64
from datetime import datetime
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# Read-only scope
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
TOKEN_FILE = "token.json"
CREDENTIALS_FILE = "credentials.json"


def authenticate():
    """Authenticate with Gmail API via OAuth2. Opens browser on first run."""
    creds = None

    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_FILE):
                print(f"ERROR: '{CREDENTIALS_FILE}' not found.")
                print("Download it from Google Cloud Console → APIs → Credentials")
                raise SystemExit(1)
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(TOKEN_FILE, "w") as token:
            token.write(creds.to_json())

    return creds


def get_message_detail(service, msg_id):
    """Fetch full message and extract headers + body snippet."""
    msg = service.users().messages().get(userId="me", id=msg_id, format="full").execute()

    headers = msg.get("payload", {}).get("headers", [])
    header_map = {h["name"]: h["value"] for h in headers}

    subject = header_map.get("Subject", "(no subject)")
    sender = header_map.get("From", "(unknown)")
    date_str = header_map.get("Date", "")

    # Try to extract plain text body
    body = ""
    payload = msg.get("payload", {})
    parts = payload.get("parts", [])

    if parts:
        for part in parts:
            if part.get("mimeType") == "text/plain":
                data = part.get("body", {}).get("data", "")
                if data:
                    body = base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")
                break
    else:
        data = payload.get("body", {}).get("data", "")
        if data:
            body = base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")

    if not body:
        body = msg.get("snippet", "")

    return {
        "subject": subject,
        "from": sender,
        "date": date_str,
        "body": body[:500],  # Truncate long bodies
        "labels": msg.get("labelIds", []),
    }


def list_messages(service, query="", max_results=10):
    """List messages matching query. Default: latest 10 from inbox."""
    results = (
        service.users()
        .messages()
        .list(userId="me", q=query, maxResults=max_results)
        .execute()
    )
    return results.get("messages", [])


def main():
    print("Authenticating with Gmail (ceslep@gmail.com)...")
    creds = authenticate()
    service = build("gmail", "v1", credentials=creds)

    # Get profile info
    profile = service.users().getProfile(userId="me").execute()
    print(f"Connected: {profile['emailAddress']}")
    print(f"Total messages: {profile.get('messagesTotal', '?')}")
    print("-" * 60)

    # Fetch latest emails
    print("\nLatest 10 emails:\n")
    messages = list_messages(service, max_results=10)

    if not messages:
        print("No messages found.")
        return

    for i, msg_ref in enumerate(messages, 1):
        msg = get_message_detail(service, msg_ref["id"])
        print(f"[{i}] {msg['subject']}")
        print(f"    From: {msg['from']}")
        print(f"    Date: {msg['date']}")
        print(f"    Preview: {msg['body'][:100]}...")
        print()


if __name__ == "__main__":
    main()
