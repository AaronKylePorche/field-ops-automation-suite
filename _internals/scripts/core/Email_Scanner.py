# Email_Scanner.py - Real-time Outlook Email Monitor
# Refactored for New User Package (Configuration-driven)
# Based on InboxTrigger_WOP.py
#
# Real-time Outlook watcher (non-blocking):
# - On NewMailEx: enqueue work and return immediately (no UI freeze)
# - Worker loop: finds original with attachments, moves to Claims, queues WOP
# - WOP launched in a new console (non-blocking) so watcher keeps handling mail
#
# All settings come from config.py - No hardcoding!
# Users only customize config.py
#
# Run: python Email_Scanner.py (Outlook must be open; requires pywin32)

import os
import sys
import time
import re
import traceback
from pathlib import Path
import subprocess
from collections import deque
from datetime import datetime

# Import config from parent directory
sys.path.insert(0, str(Path(__file__).parent.parent.parent / "config"))
import config

try:
    import pythoncom
    import win32com.client as win32
except Exception as e:
    print("pywin32 is required. Install with:")
    print("  python -m pip install pywin32")
    raise

# === CONFIG (from config.py) ================================================================

PM_WHITELIST = set(config.OUTLOOK_SETTINGS["whitelist_emails"])
TARGET_FOLDER_PATH = config.OUTLOOK_SETTINGS["target_folder_path"]
QUEUE_FOLDER = config.OUTLOOK_SETTINGS["queue_folder"]

# Globals shared between event sink and worker loop
pending_ids = deque()
g_session = None
g_dest_folder = None
g_whitelist = set()
g_processed_original_entryids = set()

RECEIVED_LINE_REGEX = re.compile(r'^\s*received\s*[\.!\?"]*\s*$', re.I)


# ==========================================================================

def set_console_title(title: str):
    if os.name == "nt":
        os.system(f"title {title}")

def print_banner(name: str, description: str):
    print("=" * 70)
    print(f"🚀  {name}")
    print("=" * 70)
    print(description.strip())
    print("=" * 70 + "\n")


def get_timestamp():
    """Return current time in human-readable format: MM/DD/YYYY H:MMAm/Pm"""
    return datetime.now().strftime("%m/%d/%Y %I:%M%p")


def get_sender_smtp(mail):
    try:
        sender = mail.Sender
        if sender is not None:
            if getattr(sender, "AddressEntryUserType", None) is not None:
                try:
                    ex = sender.GetExchangeUser()
                    if ex is not None and getattr(ex, "PrimarySmtpAddress", None):
                        return ex.PrimarySmtpAddress.lower()
                except Exception:
                    pass
        addr = getattr(mail, "SenderEmailAddress", None)
        if addr:
            return addr.lower()
    except Exception:
        pass
    return None

def first_real_line_is_received(mail):
    try:
        body = getattr(mail, "Body", "") or ""
        for raw in body.splitlines():
            line = raw.strip()
            if not line:
                continue
            if line.startswith(">"):
                continue
            return bool(RECEIVED_LINE_REGEX.match(line))
        return False
    except Exception:
        return False

def get_first_real_line(mail):
    try:
        body = getattr(mail, "Body", "") or ""
        for raw in body.splitlines():
            line = raw.strip()
            if not line:
                continue
            if line.startswith(">"):
                continue
            return line
        return ""
    except Exception:
        return ""

def is_reply_message(mail):
    try:
        pa = mail.PropertyAccessor
        HDR = "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/"
        in_reply_to = None
        try:
            in_reply_to = pa.GetProperty(HDR + "In-Reply-To")
        except Exception:
            in_reply_to = None
        if in_reply_to:
            return True
        subj = (mail.Subject or "").strip().lower()
        if subj.startswith("re:"):
            return True
    except Exception:
        pass
    return False

def resolve_folder(session, path_parts):
    if not path_parts or path_parts[0].lower() != "inbox":
        raise ValueError('TARGET_FOLDER_PATH must start with "Inbox"')
    inbox = session.GetDefaultFolder(6)  # olFolderInbox
    folder = inbox
    for name in path_parts[1:]:
        folder = folder.Folders.Item(name)
    return folder

def flatten_conversation(conv):
    try:
        roots = conv.GetRootItems()
    except Exception:
        roots = None
    if not roots:
        return
    stack = [roots.Item(i) for i in range(1, roots.Count + 1)]
    while stack:
        it = stack.pop(0)
        yield it
        try:
            kids = conv.GetChildren(it)
            if kids:
                for j in range(1, kids.Count + 1):
                    stack.append(kids.Item(j))
        except Exception:
            pass

def pick_oldest_with_attachments(conv):
    cands = []
    for it in flatten_conversation(conv):
        try:
            if getattr(it, "Class", None) != 43:
                continue
            atts = getattr(it, "Attachments", None)
            if atts is not None and atts.Count > 0:
                ts = getattr(it, "ReceivedTime", None) or getattr(it, "CreationTime", None)
                cands.append((ts, it))
        except Exception:
            pass
    if not cands:
        return None
    cands.sort(key=lambda x: x[0])
    return cands[0][1]

def get_message_id(mail):
    try:
        pa = mail.PropertyAccessor
        return pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F")
    except Exception:
        try:
            return pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001E")
        except Exception:
            return None

def get_folder_path(folder):
    try:
        return getattr(folder, "FolderPath", str(folder))
    except Exception:
        return "UNKNOWN"

def count_with_message_id(folder, message_id):
    try:
        if not message_id or not folder:
            return 0
        dasl = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
        safe = message_id.replace("'", "''")
        items = folder.Items
        try:
            restricted = items.Restrict(f'@SQL="{dasl}" = \'{safe}\'')
            return restricted.Count
        except Exception:
            items.Sort("[ReceivedTime]", True)
            total = min(500, items.Count)
            count = 0
            for i in range(1, total + 1):
                try:
                    it = items.Item(i)
                    if get_message_id(it) == message_id:
                        count += 1
                except Exception:
                    pass
            return count
    except Exception:
        return 0

def move_item_to_folder(item, dest_folder):
    try:
        parent = getattr(item, "Parent", None)
        if parent and getattr(parent, "EntryID", None) == getattr(dest_folder, "EntryID", None):
            print(f"[{get_timestamp()}]   - Original already in target folder; skipping move.")
            return item
        moved = item.Move(dest_folder)
        print(f"[{get_timestamp()}]   - Moved original with attachments to: {dest_folder.FolderPath}")
        return moved
    except Exception as e:
        print(f"[{get_timestamp()}]   ! Move failed:", e)
        return None

def write_wop_ticket():
    """Create a ticket file in queue folder to request a WOP run."""
    try:
        qdir = Path(QUEUE_FOLDER).resolve()
        qdir.mkdir(parents=True, exist_ok=True)
        import time, random
        ticket = qdir / f"wop_ticket_{int(time.time())}_{random.randint(1000,9999)}.txt"
        ticket.write_text("run", encoding="utf-8")
        print(f"[{get_timestamp()}]   -> Queued WOP run: {ticket.name}")
    except Exception as e:
        print(f"[{get_timestamp()}]   ! Failed to queue WOP run:", e)


class OutlookEventSink:
    """Enqueue-only event sink to avoid blocking Outlook UI."""
    def __init__(self):
        self.enqueued = 0

    def OnNewMailEx(self, entry_id_collection):
        try:
            ids = [e.strip() for e in entry_id_collection.split(',') if e.strip()]
            print(f"[{get_timestamp()}] NewMailEx fired: {len(ids)} item(s)")
            for eid in ids:
                pending_ids.append(eid)
                self.enqueued += 1
        except Exception:
            traceback.print_exc()

def process_entry(entry_id):
    """Heavy work: resolved off the event thread to keep Outlook responsive."""
    global g_session, g_dest_folder, g_whitelist, g_processed_original_entryids
    try:
        item = g_session.GetItemFromID(entry_id)
    except Exception:
        return  # not found

    try:
        if getattr(item, "Class", None) != 43:  # olMail
            return

        sender = get_sender_smtp(item) or "unknown"
        subj = (getattr(item, "Subject", "") or "").strip()
        print(f'[{get_timestamp()}] [NewMail] From: {sender} | Subject: "{subj}"')

        if sender.lower() not in g_whitelist:
            print(f"[{get_timestamp()}]   SKIP: not in whitelist")
            return

        first_line = get_first_real_line(item)
        if not RECEIVED_LINE_REGEX.match(first_line or ""):
            print(f'[{get_timestamp()}]   SKIP: first line not "received" -> "{first_line}"')
            return

        if not is_reply_message(item):
            print(f"[{get_timestamp()}]   SKIP: not a reply (no In-Reply-To and no RE:)")
            return

        conv = None
        try:
            conv = item.GetConversation()
        except Exception:
            conv = None

        if not conv:
            print(f"[{get_timestamp()}]   SKIP: no conversation available")
            return

        original = pick_oldest_with_attachments(conv)
        if not original:
            print(f"[{get_timestamp()}]   SKIP: no original with attachments found in your mailbox")
            return

        orig_eid = getattr(original, "EntryID", None)
        o_subj = (getattr(original, "Subject", "") or "").strip()
        o_time = getattr(original, "ReceivedTime", None) or getattr(original, "CreationTime", None)
        print(f'[{get_timestamp()}]   Selecting original with attachments -> Subject: "{o_subj}" | Received: {o_time}')

        if orig_eid in g_processed_original_entryids:
            print(f"[{get_timestamp()}]   SKIP: original already processed in this session")
            return

        # Debug before move
        src_folder = getattr(original, "Parent", None)
        src_path = get_folder_path(src_folder)
        msgid = get_message_id(original)
        eid_before = (orig_eid or "")[-8:]
        print(f"[{get_timestamp()}]   Original parent before move: {src_path}")
        print(f"[{get_timestamp()}]   Message-ID: {msgid}")
        print(f"[{get_timestamp()}]   EntryID(before): ...{eid_before}")
        try:
            src_count_pre = count_with_message_id(src_folder, msgid)
            print(f"[{get_timestamp()}]   Pre-move -- Source items with same Message-ID: {src_count_pre}")
        except Exception:
            pass

        moved = move_item_to_folder(original, g_dest_folder)
        if moved:
            g_processed_original_entryids.add(orig_eid)
            try:
                dst_count = count_with_message_id(g_dest_folder, msgid)
                src_count_post = count_with_message_id(src_folder, msgid)
                dst_path = get_folder_path(g_dest_folder)
                eid_after = (getattr(moved, "EntryID", "") or "")[-8:]
                print(f"[{get_timestamp()}]   Destination items with same Message-ID: {dst_count}")
                print(f"[{get_timestamp()}]   Post-move -- Source items with same Message-ID: {src_count_post}")
                print(f"[{get_timestamp()}]   EntryID(after): ...{eid_after}")
                print(f"[{get_timestamp()}]   Moved to: {dst_path}")
            except Exception:
                pass
            write_wop_ticket()
        else:
            print(f"[{get_timestamp()}]   SKIP: failed to move original (see above)")

    except Exception:
        traceback.print_exc()

def main():
    global g_session, g_dest_folder, g_whitelist
    print(f"[{get_timestamp()}] Starting Monitor...")
    print(f"[{get_timestamp()}] Keep this window open. Press Ctrl+C to stop.\n")
    if not PM_WHITELIST:
        print(f"[{get_timestamp()}] WARNING: PM_WHITELIST is empty. Add PM email(s) to config.py to enable triggers.\n")

    pythoncom.CoInitialize()

    app = win32.Dispatch("Outlook.Application")
    session = app.GetNamespace("MAPI")
    dest_folder = resolve_folder(session, TARGET_FOLDER_PATH)

    sink = win32.DispatchWithEvents(app, OutlookEventSink)

    g_session = session
    g_dest_folder = dest_folder
    g_whitelist = {e.lower() for e in PM_WHITELIST}

    print(f"[{get_timestamp()}] Monitoring Email...")
    print(f"[{get_timestamp()}]  - Destination folder: {dest_folder.FolderPath}")
    print(f"[{get_timestamp()}]  - Whitelist size: {len(g_whitelist)}")
    print(f"[{get_timestamp()}]  - Queue folder: {QUEUE_FOLDER}\n")

    try:
        while True:
            # Pump COM messages so NewMailEx fires
            pythoncom.PumpWaitingMessages()
            # Drain queue without blocking Outlook
            while pending_ids:
                eid = pending_ids.popleft()
                process_entry(eid)
            time.sleep(0.1)
    except KeyboardInterrupt:
        print(f"\n[{get_timestamp()}] Stopping. Bye.")

if __name__ == "__main__":
    set_console_title("Email Scanner - Outlook Listener")
    print_banner(
        "Email Scanner",
        """
        Monitors your Outlook inbox (configured in config.py) for "received" emails from
        whitelisted PM addresses and queues them for processing by WOP22.

        - Filters by PM whitelist (from config.OUTLOOK_SETTINGS["whitelist_emails"])
        - Checks for 'received' as first line (must be reply to existing message)
        - Moves original message with attachments to Claims folder
        - Writes a WOP ticket to queue folder to trigger processing

        This is part of the Claim Watcher Suite. Ensure Outlook is running.
        """
    )
    main()
