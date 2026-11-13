"""
Download emails from your Microsoft 365 mailbox (works with New Outlook)
and save them as .eml files in inbox/, so check_orders.py can read them.

Read-only: it only uses the Mail.Read permission. Nothing is sent, moved or deleted.

Setup (once):
  1. pip install O365
  2. Register an app in https://entra.microsoft.com  (App registrations > New registration)
       - Supported account types: "Accounts in this organizational directory only"
       - Redirect URI: platform "Public client/native (mobile & desktop)",
         value  https://login.microsoftonline.com/common/oauth2/nativeclient
       - API permissions: Microsoft Graph > Delegated > Mail.Read  (User.Read is there already)
  3. Copy "Application (client) ID" and "Directory (tenant) ID" from the app's Overview page
     into CLIENT_ID and TENANT_ID below (or set env vars O365_CLIENT_ID / O365_TENANT_ID).

Run:
  python fetch_emails.py                       # latest 10 emails from Inbox
  python fetch_emails.py --folder "Orders" -n 25
  python fetch_emails.py --subject "order"     # only emails whose subject contains "order"

The first run prints a link: open it, sign in, then paste the URL you land on
(a blank page) back into the terminal. After that the login is remembered.
"""

import argparse
import os
import re
import sys
from pathlib import Path

from O365 import Account, FileSystemTokenBackend

CLIENT_ID = os.environ.get("O365_CLIENT_ID", "PASTE-CLIENT-ID-HERE")
TENANT_ID = os.environ.get("O365_TENANT_ID", "PASTE-TENANT-ID-HERE")

INBOX_DIR = Path(__file__).parent / "inbox"
# Keep the login token outside the OneDrive-synced project folder.
TOKEN_DIR = Path.home() / ".order_checker"


def get_account() -> Account:
    if "PASTE" in CLIENT_ID or "PASTE" in TENANT_ID:
        sys.exit("Set CLIENT_ID and TENANT_ID at the top of fetch_emails.py first (see the setup notes).")

    TOKEN_DIR.mkdir(exist_ok=True)
    token_backend = FileSystemTokenBackend(token_path=TOKEN_DIR, token_filename="o365_token.txt")
    account = Account(
        CLIENT_ID,
        auth_flow_type="public",
        tenant_id=TENANT_ID,
        token_backend=token_backend,
    )
    if not account.is_authenticated:
        if not account.authenticate(requested_scopes=["basic", "mailbox"]):
            sys.exit("Sign-in failed.")
    return account


def safe_name(text: str, limit: int = 60) -> str:
    text = re.sub(r"[^\w\- ]+", "", text or "no subject").strip().replace(" ", "_")
    return text[:limit] or "no_subject"


def main():
    parser = argparse.ArgumentParser(description="Save mailbox emails as .eml into inbox/")
    parser.add_argument("--folder", help="Subfolder of Inbox to read (default: Inbox itself)")
    parser.add_argument("-n", "--limit", type=int, default=10, help="How many emails (default 10)")
    parser.add_argument("--subject", help="Only emails whose subject contains this text")
    args = parser.parse_args()

    account = get_account()
    mailbox = account.mailbox()
    folder = mailbox.inbox_folder()
    if args.folder:
        folder = folder.get_folder(folder_name=args.folder)
        if folder is None:
            sys.exit(f'Folder "{args.folder}" not found under Inbox.')

    query = None
    if args.subject:
        query = mailbox.new_query().on_attribute("subject").contains(args.subject)

    INBOX_DIR.mkdir(exist_ok=True)
    saved = skipped = 0
    for msg in folder.get_messages(limit=args.limit, query=query, order_by="receivedDateTime desc"):
        stamp = msg.received.strftime("%Y%m%d_%H%M%S") if msg.received else "nodate"
        path = INBOX_DIR / f"mail_{stamp}_{safe_name(msg.subject)}.eml"
        if path.exists():
            skipped += 1
            continue
        if msg.save_as_eml(to_path=path):
            saved += 1
            print(f"saved  {path.name}")
        else:
            print(f"failed {msg.subject!r}")

    print(f"\n{saved} saved, {skipped} already there. Now run: python check_orders.py")


if __name__ == "__main__":
    main()
