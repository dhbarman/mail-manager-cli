"""
mail.py — Yahoo Mail manager via IMAP.

Connects to Yahoo Mail using IMAP (SSL). Requires an App Password
(Yahoo account → Security → Generate app password).

Setup:
  export YAHOO_EMAIL="you@yahoo.com"
  export YAHOO_APP_PASSWORD="xxxx-xxxx-xxxx-xxxx"

Usage:
  python3 mail.py --inbox                        list inbox (last 20)
  python3 mail.py --inbox --limit 50             list last 50
  python3 mail.py --read ID                      read email by UID
  python3 mail.py --search "subject:invoice"     search by subject keyword
  python3 mail.py --search "from:amazon"         search by sender
  python3 mail.py --unread                       list unread only
  python3 mail.py --folders                      list all folders/labels
  python3 mail.py --folder "Sent"                list sent folder
  python3 mail.py --save ID                      save email to mails/ dir
  python3 mail.py --save-all                     save all inbox emails locally
  python3 mail.py --delete ID                    move email to Trash
  python3 mail.py --mark-read ID                 mark email as read
  python3 mail.py --send --to addr --subject S --body B            send plain email
  python3 mail.py --send --to addr --subject S --body B --html      send HTML email
  python3 mail.py --send --to addr --subject S --body B --attach f1 f2  send with attachments

  Bulk delete by filter (add --dry-run to preview without deleting):
  python3 mail.py --bulk-delete --from-addr "spam@example.com"
  python3 mail.py --bulk-delete --subject-has "unsubscribe"
  python3 mail.py --bulk-delete --body-has "click here"
  python3 mail.py --bulk-delete --older-than 90          (days)
  python3 mail.py --bulk-delete --before 2024-01-01
  python3 mail.py --bulk-delete --unread-only
  python3 mail.py --bulk-delete --from-addr "news@" --subject-has "deal" --older-than 30
"""

import os
import sys
import time
import imaplib
import smtplib
import email
import unicodedata
import json
import datetime
import argparse
import re
import yaml
from email.header      import decode_header
from email.mime.text   import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base   import MIMEBase
from email.mime.application import MIMEApplication
from email             import encoders
import mimetypes

# ─────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────

YAHOO_EMAIL        = os.getenv("YAHOO_EMAIL")
YAHOO_APP_PASSWORD = os.getenv("YAHOO_APP_PASSWORD")

IMAP_HOST = "imap.mail.yahoo.com"
IMAP_PORT = 993
SMTP_HOST = "smtp.mail.yahoo.com"
SMTP_PORT = 465

SAVE_DIR       = os.path.join(os.path.dirname(__file__), "mails")
HISTORY_FILE   = os.path.join(os.path.dirname(__file__), "delete_history.json")
TEMPLATES_FILE = os.path.join(os.path.dirname(__file__), "email_templates.yaml")


# ─────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────

def _strip_invisible(text):
    """Remove invisible Unicode chars spammers insert to defeat keyword matching."""
    return ''.join(
        c for c in text
        if unicodedata.category(c) not in ('Cf', 'Co', 'Cc')
    )


def _decode_str(value):
    """Decode email header string (handles encoded words)."""
    if value is None:
        return ""
    parts = decode_header(value)
    decoded = []
    _FALLBACK_ENCS = {"unknown-8bit", "x-unknown", "unknown", ""}
    for part, enc in parts:
        if isinstance(part, bytes):
            charset = (enc or "").lower()
            if charset in _FALLBACK_ENCS:
                charset = "latin-1"
            try:
                decoded.append(part.decode(charset, errors="replace"))
            except (LookupError, TypeError):
                decoded.append(part.decode("latin-1", errors="replace"))
        else:
            decoded.append(str(part))
    return " ".join(decoded)


def _get_body(msg):
    """Extract plain text body from email message."""
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            cd = str(part.get("Content-Disposition", ""))
            if ct == "text/plain" and "attachment" not in cd:
                charset = part.get_content_charset() or "utf-8"
                body = part.get_payload(decode=True).decode(charset, errors="replace")
                break
            elif ct == "text/html" and "attachment" not in cd and not body:
                charset = part.get_content_charset() or "utf-8"
                raw_html = part.get_payload(decode=True).decode(charset, errors="replace")
                body = re.sub(r"<[^>]+>", " ", raw_html)
                body = re.sub(r"\s+", " ", body).strip()
    else:
        charset = msg.get_content_charset() or "utf-8"
        body = msg.get_payload(decode=True).decode(charset, errors="replace")
    return body.strip()


def _connect():
    """Connect and authenticate to Yahoo IMAP."""
    if not YAHOO_EMAIL or not YAHOO_APP_PASSWORD:
        print("Error: set YAHOO_EMAIL and YAHOO_APP_PASSWORD env vars")
        print("  Generate an app password at: Yahoo → Account Security → App passwords")
        sys.exit(1)
    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        mail.login(YAHOO_EMAIL, YAHOO_APP_PASSWORD)
        return mail
    except imaplib.IMAP4.error as e:
        print(f"[error] Login failed: {e}")
        print("  Make sure you're using an App Password, not your Yahoo account password.")
        sys.exit(1)


def _select_folder(mail, folder="INBOX"):
    status, data = mail.select(folder)
    if status != "OK":
        print(f"[error] Could not select folder '{folder}': {data}")
        sys.exit(1)
    return int(data[0])


def _fetch_message(mail, uid):
    """Fetch and parse a single email by UID."""
    status, data = mail.uid("fetch", uid, "(RFC822)")
    if status != "OK" or not data or data[0] is None:
        return None
    raw = data[0][1]
    return email.message_from_bytes(raw)


def _parse_summary(msg, uid):
    """Return a summary dict from a parsed email."""
    return {
        "uid":     uid.decode() if isinstance(uid, bytes) else str(uid),
        "from":    _decode_str(msg.get("From", "")),
        "to":      _decode_str(msg.get("To", "")),
        "subject": _decode_str(msg.get("Subject", "(no subject)")),
        "date":    msg.get("Date", ""),
    }


# ─────────────────────────────────────────
# COMMANDS
# ─────────────────────────────────────────

def list_folders():
    mail = _connect()
    status, folders = mail.list()
    mail.logout()
    print(f"\n  FOLDERS")
    print(f"  {'─'*50}")
    for f in folders:
        decoded = f.decode() if isinstance(f, bytes) else str(f)
        # extract folder name from: (\HasNoChildren) "/" "INBOX"
        parts = decoded.split('"')
        name  = parts[-2] if len(parts) >= 2 else decoded
        print(f"  {name}")
    print()


def list_inbox(folder="INBOX", limit=20, unread_only=False):
    mail  = _connect()
    count = _select_folder(mail, folder)

    search_criteria = "UNSEEN" if unread_only else "ALL"
    status, data    = mail.uid("search", None, search_criteria)
    if status != "OK":
        print("[error] Search failed")
        mail.logout(); return

    uids  = data[0].split()
    total = len(uids)
    uids  = uids[-limit:]   # most recent N

    label = "unread" if unread_only else "total"
    print(f"\n  {folder}  |  {total} {label} emails  |  showing last {len(uids)}")
    print(f"\n  {'UID':<8} {'Date':<22} {'From':<30} Subject")
    print(f"  {'─'*100}")

    for uid in reversed(uids):
        msg     = _fetch_message(mail, uid)
        if not msg:
            continue
        summary = _parse_summary(msg, uid)
        date    = summary["date"][:22] if summary["date"] else "—"
        sender  = summary["from"][:28]
        subject = summary["subject"][:50]
        print(f"  {summary['uid']:<8} {date:<22} {sender:<30} {subject}")

    print()
    mail.logout()


def read_email(uid, folder="INBOX"):
    mail = _connect()
    _select_folder(mail, folder)
    msg  = _fetch_message(mail, uid.encode() if isinstance(uid, str) else uid)
    if not msg:
        print(f"[error] Email UID {uid} not found.")
        mail.logout(); return

    summary = _parse_summary(msg, uid)
    body    = _get_body(msg)

    print(f"\n{'='*70}")
    print(f"  From    : {summary['from']}")
    print(f"  To      : {summary['to']}")
    print(f"  Subject : {summary['subject']}")
    print(f"  Date    : {summary['date']}")
    print(f"{'='*70}\n")
    print(body[:5000])
    if len(body) > 5000:
        print(f"\n  ... [{len(body) - 5000} more characters — use --save {uid} to get full email]")
    print()
    mail.logout()


def search_emails(query, folder="INBOX", limit=20):
    """
    query examples:
      'subject:invoice'  → search by subject keyword
      'from:amazon'      → search by sender
      'keyword'          → full text search (subject + from)
    """
    mail = _connect()
    _select_folder(mail, folder)

    if query.startswith("subject:"):
        keyword  = query[8:].strip()
        criteria = f'SUBJECT "{keyword}"'
    elif query.startswith("from:"):
        keyword  = query[5:].strip()
        criteria = f'FROM "{keyword}"'
    else:
        criteria = f'TEXT "{query}"'

    status, data = mail.uid("search", None, criteria)
    if status != "OK":
        print("[error] Search failed")
        mail.logout(); return

    uids  = data[0].split()
    total = len(uids)
    uids  = uids[-limit:]

    print(f"\n  Search: '{query}'  |  {total} results  |  showing last {len(uids)}")
    print(f"\n  {'UID':<8} {'Date':<22} {'From':<30} Subject")
    print(f"  {'─'*100}")

    for uid in reversed(uids):
        msg = _fetch_message(mail, uid)
        if not msg:
            continue
        summary = _parse_summary(msg, uid)
        date    = summary["date"][:22] if summary["date"] else "—"
        print(f"  {summary['uid']:<8} {date:<22} {summary['from'][:28]:<30} {summary['subject'][:50]}")

    print()
    mail.logout()


def save_email(uid, folder="INBOX"):
    """Save a single email as JSON + txt to the mails/ directory."""
    os.makedirs(SAVE_DIR, exist_ok=True)
    mail = _connect()
    _select_folder(mail, folder)

    uid_bytes = uid.encode() if isinstance(uid, str) else uid
    msg       = _fetch_message(mail, uid_bytes)
    if not msg:
        print(f"[error] Email UID {uid} not found.")
        mail.logout(); return

    summary = _parse_summary(msg, uid)
    body    = _get_body(msg)

    safe_subj = re.sub(r"[^\w\s-]", "", summary["subject"])[:40].strip().replace(" ", "_")
    filename  = f"{uid}_{safe_subj}"

    # save plain text
    txt_path = os.path.join(SAVE_DIR, f"{filename}.txt")
    with open(txt_path, "w") as f:
        f.write(f"From   : {summary['from']}\n")
        f.write(f"To     : {summary['to']}\n")
        f.write(f"Subject: {summary['subject']}\n")
        f.write(f"Date   : {summary['date']}\n")
        f.write(f"UID    : {summary['uid']}\n")
        f.write(f"\n{'='*60}\n\n")
        f.write(body)

    # save metadata as JSON
    json_path = os.path.join(SAVE_DIR, f"{filename}.json")
    with open(json_path, "w") as f:
        json.dump({**summary, "body_preview": body[:500]}, f, indent=2)

    print(f"[save] {txt_path}")
    print(f"[save] {json_path}")
    mail.logout()


def save_all(folder="INBOX", limit=100):
    """Save last N emails from a folder to the mails/ directory."""
    os.makedirs(SAVE_DIR, exist_ok=True)
    mail = _connect()
    _select_folder(mail, folder)

    status, data = mail.uid("search", None, "ALL")
    uids         = data[0].split()[-limit:]
    print(f"[save-all] Saving {len(uids)} emails from {folder}...")

    for i, uid in enumerate(reversed(uids), 1):
        msg = _fetch_message(mail, uid)
        if not msg:
            continue
        summary   = _parse_summary(msg, uid)
        body      = _get_body(msg)
        safe_subj = re.sub(r"[^\w\s-]", "", summary["subject"])[:40].strip().replace(" ", "_")
        filename  = f"{uid.decode()}_{safe_subj}"
        txt_path  = os.path.join(SAVE_DIR, f"{filename}.txt")
        with open(txt_path, "w") as f:
            f.write(f"From   : {summary['from']}\n")
            f.write(f"To     : {summary['to']}\n")
            f.write(f"Subject: {summary['subject']}\n")
            f.write(f"Date   : {summary['date']}\n\n")
            f.write(body)
        print(f"  [{i}/{len(uids)}] {summary['subject'][:50]}")

    print(f"[save-all] Done. Files saved to {SAVE_DIR}/")
    mail.logout()


def delete_email(uid, folder="INBOX"):
    mail = _connect()
    _select_folder(mail, folder)
    mail.uid("store", uid.encode(), "+FLAGS", "\\Deleted")
    mail.expunge()
    print(f"[delete] Email UID {uid} moved to Trash.")
    mail.logout()


def _build_cmd(folder, from_addr, subject_has, body_has, older_than, before, unread_only, parallel):
    """Reconstruct the shell command from bulk-delete args."""
    parts = ["python3 mail.py --bulk-delete"]
    if folder and folder != "INBOX":
        parts.append(f'--folder "{folder}"')
    if from_addr:
        parts.append("--from-addr " + " ".join(f'"{a}"' for a in from_addr))
    if subject_has:
        parts.append("--subject-has " + " ".join(f'"{s}"' for s in subject_has))
    if body_has:
        parts.append(f'--body-has "{body_has}"')
    if older_than is not None:
        parts.append(f"--older-than {older_than}")
    if before:
        parts.append(f"--before {before}")
    if unread_only:
        parts.append("--unread-only")
    if parallel:
        parts.append("--parallel")
    return " ".join(parts)


def _save_history(cmd, deleted, dry_run):
    """Append a bulk-delete run to the history file."""
    history = []
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE) as f:
            try:
                history = json.load(f)
            except Exception:
                history = []
    entry = {
        "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "cmd":       cmd,
        "deleted":   deleted,
        "dry_run":   dry_run,
    }
    history.append(entry)
    with open(HISTORY_FILE, "w") as f:
        json.dump(history, f, indent=2)


def export_filters():
    """Read delete_history.json and print Yahoo Mail filter instructions."""
    if not os.path.exists(HISTORY_FILE):
        print("  No delete history found.")
        return
    with open(HISTORY_FILE) as f:
        try:
            history = json.load(f)
        except Exception:
            print("  [error] Could not read history file.")
            return

    if not history:
        print("  No history entries found.")
        return

    # Parse CLI flags from saved command string
    def _parse_cmd(cmd):
        criteria = {}
        m = re.search(r'--from-addr\s+((?:"[^"]+"\s*)+)', cmd)
        if m:
            criteria["from"] = re.findall(r'"([^"]+)"', m.group(1))
        subjects = re.findall(r'--subject-has\s+((?:"[^"]+"\s*)+)', cmd)
        if subjects:
            criteria["subject"] = re.findall(r'"([^"]+)"', subjects[0])
        m = re.search(r'--body-has\s+"([^"]+)"', cmd)
        if m:
            criteria["body"] = m.group(1)
        m = re.search(r'--older-than\s+(\d+)', cmd)
        if m:
            criteria["older_than"] = m.group(1)
        m = re.search(r'--before\s+(\S+)', cmd)
        if m:
            criteria["before"] = m.group(1)
        if "--unread-only" in cmd:
            criteria["unread_only"] = True
        return criteria

    # Deduplicate: same criteria = same filter
    seen   = []
    unique = []
    for entry in history:
        c = _parse_cmd(entry.get("cmd", ""))
        key = json.dumps(c, sort_keys=True)
        if key not in seen and c:
            seen.append(key)
            unique.append((entry, c))

    print("\n" + "═" * 60)
    print("  YAHOO MAIL FILTER EXPORT")
    print("═" * 60)
    print("  How to add: Yahoo Mail → Settings (gear icon)")
    print("  → More Settings → Filters → Add new filter")
    print("═" * 60 + "\n")

    for i, (entry, c) in enumerate(unique, 1):
        ts      = entry.get("timestamp", "")
        deleted = entry.get("deleted", "?")
        dry     = " [dry-run]" if entry.get("dry_run") else ""

        print(f"  Filter #{i}  (from history: {ts}{dry}, matched {deleted} email(s))")
        print(f"  {'-'*56}")

        if c.get("from"):
            for addr in c["from"]:
                print(f"  Sender contains     : {addr}")
        if c.get("subject"):
            for s in (c["subject"] if isinstance(c["subject"], list) else [c["subject"]]):
                print(f"  Subject contains    : {s}")
        if c.get("body"):
            print(f"  Email contains text : {c['body']}")
        if c.get("older_than"):
            print(f"  Note                : older than {c['older_than']} days (Yahoo filters don't support age — set action below and run script for old emails)")
        if c.get("before"):
            print(f"  Note                : before {c['before']} (same — Yahoo has no date filter, handle old ones via script)")
        if c.get("unread_only"):
            print(f"  Status              : Unread only")

        print(f"  Action              : Delete")
        print()

    print("═" * 60)
    print(f"  Total unique filters: {len(unique)}")
    print("═" * 60)
    print("\n  Steps in Yahoo Mail:")
    print("  1. Open mail.yahoo.com → gear icon → More Settings")
    print("  2. Click 'Filters' in the left sidebar")
    print("  3. Click 'Add new filter'")
    print("  4. Enter filter name, paste criteria above, set Action = Delete")
    print("  5. Save — Yahoo will apply it to all future incoming emails\n")


def show_history():
    """Print saved bulk-delete history."""
    if not os.path.exists(HISTORY_FILE):
        print("  No delete history found.")
        return
    with open(HISTORY_FILE) as f:
        try:
            history = json.load(f)
        except Exception:
            print("  [error] Could not read history file.")
            return
    if not history:
        print("  No delete history found.")
        return
    print(f"\n  {'#':<4} {'Timestamp':<22} {'Del':>5}  Command")
    print(f"  {'─'*100}")
    for i, entry in enumerate(history, 1):
        ts      = entry.get("timestamp", "")
        deleted = entry.get("deleted", "—")
        dry     = " [dry-run]" if entry.get("dry_run") else ""
        cmd     = entry.get("cmd", "")
        print(f"  {i:<4} {ts:<22} {str(deleted):>5}{dry}  {cmd}")
    print(f"\n  Run again with: python3 mail.py --replay N")
    print()


def replay_history(index):
    """Re-run a saved bulk-delete command by 1-based index."""
    if not os.path.exists(HISTORY_FILE):
        print("  No delete history found.")
        return
    with open(HISTORY_FILE) as f:
        history = json.load(f)
    if index < 1 or index > len(history):
        print(f"  [error] Index {index} out of range (1–{len(history)}).")
        return
    cmd = history[index - 1]["cmd"]
    print(f"  Replaying: {cmd}\n")
    os.system(cmd)


def list_senders(folder="INBOX", limit=None, sort_by="count"):
    """Fetch all headers and print unique From addresses with email count."""
    mail = _connect()
    _select_folder(mail, folder)

    status, data = mail.uid("search", None, "ALL")
    if status != "OK":
        print("[error] Search failed")
        mail.logout()
        return

    all_uids = data[0].split()
    if not all_uids:
        print("  Mailbox is empty.")
        mail.logout()
        return

    print(f"\n  Fetching headers from {len(all_uids)} email(s) in '{folder}'...")

    CHUNK   = 50
    RETRIES = 3

    def _parse_chunk(hdata):
        result = {}
        for item in hdata:
            if not isinstance(item, tuple):
                continue
            m = re.search(rb"UID (\d+)", item[0], re.IGNORECASE)
            if m:
                result[m.group(1)] = email.message_from_bytes(item[1])
        return result

    header_map = {}
    chunks = [all_uids[i: i + CHUNK] for i in range(0, len(all_uids), CHUNK)]
    for i, chunk in enumerate(chunks, 1):
        for attempt in range(RETRIES):
            uid_set = b",".join(chunk)
            st, hdata = mail.uid("fetch", uid_set, "(RFC822.HEADER)")
            if st == "OK":
                header_map.update(_parse_chunk(hdata))
                break
        print(f"  [{i}/{len(chunks)}] fetched {len(header_map)} headers", end="\r")

    mail.logout()
    print()

    # count emails per sender
    from collections import Counter
    import email.utils as _eutils

    counts  = Counter()
    display = {}   # normalised addr → display string "Name <addr>"
    for msg in header_map.values():
        raw  = _decode_str(msg.get("From", ""))
        name, addr = _eutils.parseaddr(raw)
        addr = addr.lower().strip()
        if not addr:
            addr = raw.strip().lower()
        counts[addr] += 1
        if addr not in display:
            display[addr] = f"{name} <{addr}>" if name else addr

    if not counts:
        print("  No senders found.")
        return

    if sort_by == "count":
        ranked = counts.most_common()
    else:
        ranked = sorted(counts.items(), key=lambda x: x[0])

    if limit:
        ranked = ranked[:limit]

    col = max(len(display[a]) for a, _ in ranked)
    print(f"\n  {'Sender':<{col}}  Count")
    print(f"  {'─' * col}  ─────")
    for addr, cnt in ranked:
        print(f"  {display[addr]:<{col}}  {cnt}")
    print(f"\n  {len(counts)} unique sender(s) across {len(all_uids)} email(s).")


def bulk_delete(
    folder="INBOX",
    from_addr=None,
    subject_has=None,
    body_has=None,
    older_than=None,
    before=None,
    unread_only=False,
    dry_run=False,
    parallel=False,
    match_any=False,
    auto_confirm=False,
):
    """
    Delete emails matching filters.

    from_addr   – list of substrings; email matches if From contains ANY of them
    subject_has – list of substrings; email matches if Subject contains ANY of them (OR within)
    body_has    – substring match on plain-text body (case-insensitive)
    older_than  – integer days; delete emails older than N days from today (always AND)
    before      – date string "YYYY-MM-DD"; delete emails sent before this date (always AND)
    unread_only – only delete emails that have not been read (always AND)
    match_any   – if True, from_addr/subject_has/body_has are OR; default is AND
    dry_run     – list matches but do not delete
    """
    import email.utils as _eutils

    mail = _connect()
    _select_folder(mail, folder)

    # Yahoo IMAP server-side search is unreliable — fetch all UIDs, filter client-side.
    base_criteria = "UNSEEN" if unread_only else "ALL"
    status, data  = mail.uid("search", None, base_criteria)
    if status != "OK":
        print("[error] Search failed")
        mail.logout()
        return

    all_uids = data[0].split()
    if not all_uids:
        print("  Mailbox is empty.")
        mail.logout()
        return

    # Compute date cutoff if needed
    cutoff_date = None
    if older_than is not None:
        cutoff_date = datetime.date.today() - datetime.timedelta(days=int(older_than))
    elif before:
        cutoff_date = datetime.date.fromisoformat(before)

    print(f"\n  Scanning {len(all_uids)} email(s) in '{folder}'...")

    # ── Phase 1: batch-fetch headers only ──
    CHUNK   = 50    # smaller chunks — Yahoo IMAP drops large requests silently
    RETRIES = 3
    header_map = {}   # uid (bytes) → parsed email message (headers only)

    def _parse_chunk_response(hdata):
        result = {}
        for item in hdata:
            if not isinstance(item, tuple):
                continue
            m = re.search(rb"UID (\d+)", item[0], re.IGNORECASE)
            if not m:
                continue
            result[m.group(1)] = email.message_from_bytes(item[1])
        return result

    def _fetch_header_chunk(chunk):
        for attempt in range(RETRIES):
            try:
                conn    = _connect()
                conn.select(folder)
                uid_set = b",".join(chunk)
                st, hdata = conn.uid("fetch", uid_set, "(RFC822.HEADER)")
                conn.logout()
                if st == "OK":
                    return _parse_chunk_response(hdata)
            except Exception as e:
                if attempt == RETRIES - 1:
                    print(f"  [warn] chunk failed after {RETRIES} attempts: {e}")
        return {}

    chunks = [all_uids[i: i + CHUNK] for i in range(0, len(all_uids), CHUNK)]
    print(f"  Fetching headers in {len(chunks)} chunk(s) of {CHUNK}...")

    if parallel and len(chunks) > 1:
        from concurrent.futures import ThreadPoolExecutor, as_completed
        workers = min(8, len(chunks))
        with ThreadPoolExecutor(max_workers=workers) as pool:
            futures = [pool.submit(_fetch_header_chunk, c) for c in chunks]
            for fut in as_completed(futures):
                header_map.update(fut.result())
    else:
        # reuse existing connection — avoids Yahoo throttling new connections
        for i, chunk in enumerate(chunks, 1):
            uid_set = b",".join(chunk)
            st, hdata = mail.uid("fetch", uid_set, "(RFC822.HEADER)")
            if st != "OK":
                print(f"  [warn] chunk {i} fetch failed (status={st})")
                continue
            result = _parse_chunk_response(hdata)
            header_map.update(result)
            print(f"  [{i}/{len(chunks)}] fetched {len(result)}/{len(chunk)} headers", end="\r")

    print()
    if len(header_map) < len(all_uids):
        print(f"  [warn] Only fetched {len(header_map)}/{len(all_uids)} headers — some may have been missed")

    # debug: show sample subjects to verify headers parsed correctly
    if header_map:
        sample = [_strip_invisible(_decode_str(msg.get("Subject", "(none)"))) for msg in list(header_map.values())[:5]]
        print(f"  [debug] sample subjects (cleaned): {sample}")

    # ── Phase 2: filter by headers + date ──
    candidates = []   # (uid_bytes, summary) that pass header filters
    for uid_bytes, msg in header_map.items():
        summary = _parse_summary(msg, uid_bytes)

        clean_from    = _strip_invisible(summary["from"].lower())
        clean_subject = _strip_invisible(summary["subject"].lower())
        from_match    = from_addr   and any(f.lower() in clean_from    for f in from_addr)
        subject_match = subject_has and any(s.lower() in clean_subject for s in subject_has)

        if match_any:
            # OR: pass if any content filter matches (only consider filters that were specified)
            active = [x for x in [
                from_match    if from_addr    else None,
                subject_match if subject_has  else None,
            ] if x is not None]
            if active and not any(active):
                continue
        else:
            # AND: all specified content filters must match
            if from_addr    and not from_match:
                continue
            if subject_has  and not subject_match:
                continue
        if cutoff_date:
            try:
                raw      = re.sub(r"\s*\(.*?\)\s*$", "", summary["date"]).strip()
                parsed   = _eutils.parsedate(raw)
                msg_date = datetime.date(*parsed[:3])
                if msg_date >= cutoff_date:
                    continue
            except Exception:
                continue  # skip emails with unparseable dates when date filter is active

        candidates.append((uid_bytes, summary))

    # ── Phase 3: body filter only for remaining candidates (if needed) ──
    to_delete = []
    if body_has:
        print(f"  Header scan done — checking body in {len(candidates)} candidate(s)...")

        def _check_body(uid_bytes, summary):
            conn = _connect()
            conn.select(folder)
            st2, fdata = conn.uid("fetch", uid_bytes, "(RFC822)")
            conn.logout()
            if st2 != "OK" or not fdata or fdata[0] is None:
                return None
            full_msg  = email.message_from_bytes(fdata[0][1])
            body_text = _get_body(full_msg)
            body_match = body_has.lower() in body_text.lower()
            if match_any:
                # In OR mode, candidates already passed at least one header filter,
                # so body match alone is enough to include; but if body_has is the
                # ONLY filter, require it to match.
                only_body = not from_addr and not subject_has
                return (uid_bytes, summary) if (body_match or not only_body) else None
            else:
                return (uid_bytes, summary) if body_match else None

        if parallel:
            from concurrent.futures import ThreadPoolExecutor, as_completed
            workers = min(10, len(candidates))
            with ThreadPoolExecutor(max_workers=workers) as pool:
                futures = {pool.submit(_check_body, u, s): (u, s) for u, s in candidates}
                for fut in as_completed(futures):
                    result = fut.result()
                    if result:
                        to_delete.append(result)
        else:
            for uid_bytes, summary in candidates:
                result = _check_body(uid_bytes, summary)
                if result:
                    to_delete.append(result)
    else:
        to_delete = candidates

    if not to_delete:
        print("  No emails passed all filter criteria.")
        mail.logout()
        return

    print(f"\n  {'UID':<8} {'Date':<22} {'From':<30} Subject")
    print(f"  {'─'*100}")
    for uid, s in to_delete:
        date    = s["date"][:22] if s["date"] else "—"
        sender  = s["from"][:28]
        subject = s["subject"][:50]
        print(f"  {s['uid']:<8} {date:<22} {sender:<30} {subject}")

    print(f"\n  Total: {len(to_delete)} email(s) will be deleted.")

    if dry_run:
        print("  [dry-run] No emails were deleted.")
        mail.logout()
        cmd = _build_cmd(folder, from_addr, subject_has, body_has, older_than, before, unread_only, parallel)
        _save_history(cmd, len(to_delete), dry_run=True)
        print(f"  [history] Dry-run saved to {HISTORY_FILE}")
        return

    if auto_confirm:
        print(f"\n  [--yes] Auto-confirming deletion of {len(to_delete)} email(s).")
        confirm = "y"
    else:
        confirm = input(f"\n  Delete {len(to_delete)} email(s)? [y/N] ").strip().lower()
    if confirm != "y":
        print("  Aborted.")
        mail.logout()
        return

    uid_set = b",".join(uid for uid, _ in to_delete)
    mail.uid("store", uid_set, "+FLAGS", "\\Deleted")
    deleted = len(to_delete)

    mail.expunge()
    print(f"\n  [bulk-delete] {deleted} email(s) deleted from '{folder}'.")
    mail.logout()
    cmd = _build_cmd(folder, from_addr, subject_has, body_has, older_than, before, unread_only, parallel)
    _save_history(cmd, deleted, dry_run=False)
    print(f"  [history] Command saved to {HISTORY_FILE}")


def mark_read(uid, folder="INBOX"):
    mail = _connect()
    _select_folder(mail, folder)
    mail.uid("store", uid.encode(), "+FLAGS", "\\Seen")
    print(f"[mark-read] Email UID {uid} marked as read.")
    mail.logout()


def load_template(name, template_file=None):
    """Load a named template from the YAML templates file."""
    path = template_file or TEMPLATES_FILE
    if not os.path.exists(path):
        print(f"[error] Templates file not found: {path}")
        print(f"  Create it at {TEMPLATES_FILE} or pass --template-file PATH")
        sys.exit(1)
    with open(path) as f:
        data = yaml.safe_load(f)
    if name not in data:
        available = ", ".join(data.keys())
        print(f"[error] Template '{name}' not found. Available: {available}")
        sys.exit(1)
    return data[name]


def list_templates(template_file=None):
    """Print all templates defined in the YAML file."""
    path = template_file or TEMPLATES_FILE
    if not os.path.exists(path):
        print(f"  No templates file found at {path}")
        return
    with open(path) as f:
        data = yaml.safe_load(f) or {}
    if not data:
        print("  Templates file is empty.")
        return
    print(f"\n  Templates in {path}\n")
    for name, tpl in data.items():
        subject     = tpl.get("subject", "(no subject)")
        html_flag   = "HTML" if tpl.get("html") else "plain"
        attachments = tpl.get("attachments") or []
        print(f"  [{name}]")
        print(f"    Subject    : {subject}")
        print(f"    Body type  : {html_flag}")
        if attachments:
            for a in attachments:
                print(f"    Attachment : {a}")
        print()


def _find_folder(mail, keywords):
    """Find a folder whose name contains any of the given keywords (case-insensitive)."""
    status, folders = mail.list()
    for f in folders:
        raw  = f.decode() if isinstance(f, bytes) else str(f)
        name_lower = raw.lower()
        if any(kw in name_lower for kw in keywords):
            parts = raw.split('"')
            if len(parts) >= 2:
                return parts[-2]
    return None


def _find_sent_folder(mail):
    return _find_folder(mail, ["sent"]) or "Sent"


# Allowed clearable folders and their search keywords
_CLEARABLE = {
    "sent":  ["sent"],
    "trash": ["trash", "deleted"],
    "spam":  ["spam", "bulk", "junk"],
}


def clear_system_folder(target):
    """Delete all emails in Sent, Trash, or Spam — no other folder allowed."""
    target = target.lower()
    if target not in _CLEARABLE:
        print(f"[error] '{target}' is not clearable. Choose from: sent, trash, spam")
        return

    mail = _connect()
    folder_name = _find_folder(mail, _CLEARABLE[target])
    if not folder_name:
        print(f"[error] Could not find a '{target}' folder on the server.")
        mail.logout()
        return

    try:
        status, data = mail.select(f'"{folder_name}"')
        if status != "OK":
            print(f"[error] Could not open folder '{folder_name}'")
            mail.logout()
            return
    except Exception as e:
        print(f"[error] {e}")
        mail.logout()
        return

    status, data = mail.uid("search", None, "ALL")
    uids = data[0].split() if status == "OK" else []

    if not uids:
        print(f"  '{folder_name}' is already empty.")
        mail.logout()
        return

    print(f"\n  Folder  : {folder_name}")
    print(f"  Emails  : {len(uids)}")
    confirm = input(f"\n  Permanently delete all {len(uids)} email(s) from '{folder_name}'? [y/N] ").strip().lower()
    if confirm != "y":
        print("  Aborted.")
        mail.logout()
        return

    # Mark all as deleted in one batch, then expunge
    uid_set = b",".join(uids)
    mail.uid("store", uid_set, "+FLAGS", "\\Deleted")
    mail.expunge()
    print(f"  [clear] {len(uids)} email(s) permanently deleted from '{folder_name}'.")
    mail.logout()


def delete_from_sent(subject):
    """Delete the most recent email in Sent folder matching the subject."""
    mail = _connect()
    sent_folder = _find_sent_folder(mail)
    try:
        status, _ = mail.select(f'"{sent_folder}"')
        if status != "OK":
            print(f"[warn] Could not open Sent folder '{sent_folder}'")
            mail.logout()
            return
    except Exception as e:
        print(f"[warn] Could not open Sent folder: {e}")
        mail.logout()
        return

    status, data = mail.uid("search", None, f'SUBJECT "{subject}"')
    if status != "OK" or not data[0].split():
        print(f"[warn] Could not find sent email with subject: {subject}")
        mail.logout()
        return

    uids     = data[0].split()
    latest   = uids[-1]   # most recently sent match
    mail.uid("store", latest, "+FLAGS", "\\Deleted")
    mail.expunge()
    print(f"[delete-sent] Removed from '{sent_folder}' folder.")
    mail.logout()


def send_email(to, subject, body, attachments=None, html=False, delete_sent=False):
    """
    Send an email with optional attachments.

    to           – recipient address (single string or comma-separated)
    subject      – subject line
    body         – plain text or HTML body
    attachments  – list of file paths to attach
    html         – if True, body is sent as text/html
    delete_sent  – if True, delete the copy from Sent folder after sending
    """
    if not YAHOO_EMAIL or not YAHOO_APP_PASSWORD:
        print("Error: set YAHOO_EMAIL and YAHOO_APP_PASSWORD env vars")
        sys.exit(1)

    msg            = MIMEMultipart()
    msg["From"]    = YAHOO_EMAIL
    msg["To"]      = to
    msg["Subject"] = subject

    # Body — plain or HTML
    body_type = "html" if html else "plain"
    msg.attach(MIMEText(body, body_type, "utf-8"))

    # Attachments
    for path in (attachments or []):
        path = os.path.expanduser(path)
        if not os.path.isfile(path):
            print(f"[warn] Attachment not found, skipping: {path}")
            continue

        filename  = os.path.basename(path)
        mime_type, _ = mimetypes.guess_type(path)
        if mime_type is None:
            mime_type = "application/octet-stream"
        main_type, sub_type = mime_type.split("/", 1)

        with open(path, "rb") as f:
            data = f.read()

        if main_type == "text":
            part = MIMEText(data.decode("utf-8", errors="replace"), _subtype=sub_type)
        elif main_type == "image":
            part = MIMEBase(main_type, sub_type)
            part.set_payload(data)
            encoders.encode_base64(part)
        else:
            part = MIMEApplication(data, Name=filename)

        part["Content-Disposition"] = f'attachment; filename="{filename}"'
        msg.attach(part)
        print(f"  [attach] {filename}  ({len(data)//1024 or 1} KB)")

    try:
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
            server.login(YAHOO_EMAIL, YAHOO_APP_PASSWORD)
            server.sendmail(YAHOO_EMAIL, to, msg.as_string())
        attach_count = len(attachments) if attachments else 0
        print(f"[send] Email sent to {to}  |  Subject: {subject}"
              + (f"  |  {attach_count} attachment(s)" if attach_count else ""))
        if delete_sent:
            confirm = input("  Delete copy from Sent folder? [y/N] ").strip().lower()
            if confirm == "y":
                delete_from_sent(subject)
            else:
                print("  Kept in Sent folder.")
    except Exception as e:
        print(f"[error] Failed to send: {e}")


# ─────────────────────────────────────────
# CLI
# ─────────────────────────────────────────

def print_examples():
    examples = """
╔══════════════════════════════════════════════════════════════════╗
║              mail.py — usage examples                           ║
╚══════════════════════════════════════════════════════════════════╝

── INBOX ───────────────────────────────────────────────────────────
  List last 20 emails:
    python3 mail.py --inbox

  List last 50 emails from a folder:
    python3 mail.py --inbox --folder Spam --limit 50

  List unread only:
    python3 mail.py --unread

  List all folders:
    python3 mail.py --folders

── READ / SEARCH ────────────────────────────────────────────────────
  Read email by UID:
    python3 mail.py --read 12345

  Search by sender:
    python3 mail.py --search "from:boss@company.com"

  Search by subject:
    python3 mail.py --search "subject:invoice"

  Search by keyword:
    python3 mail.py --search "meeting"

── SEND ────────────────────────────────────────────────────────────
  Send a plain text email:
    python3 mail.py --send --to "alice@example.com" \\
        --subject "Hello" --body "Hi Alice!"

  Send with attachment:
    python3 mail.py --send --to "alice@example.com" \\
        --subject "Report" --body "See attached." \\
        --attach /path/to/report.pdf

  Send HTML email:
    python3 mail.py --send --to "alice@example.com" \\
        --subject "Hello" --body "<h1>Hi!</h1>" --html

  Send using a template:
    python3 mail.py --send --to "recruiter@company.com" \\
        --template job_application

  Send using template, override subject:
    python3 mail.py --send --to "recruiter@company.com" \\
        --template job_application \\
        --subject "Application for Senior Engineer role"

  List saved templates:
    python3 mail.py --list-templates

── SENDERS ─────────────────────────────────────────────────────────
  List all unique senders (sorted by email count):
    python3 mail.py --list-senders

  From a specific folder:
    python3 mail.py --list-senders --folder Spam

  Top 20 senders only:
    python3 mail.py --list-senders --limit 20

  Sorted alphabetically by address:
    python3 mail.py --list-senders --sort-by addr

── DELETE ──────────────────────────────────────────────────────────
  Delete single email by UID:
    python3 mail.py --delete 12345

  Clear all sent mail:
    python3 mail.py --clear sent

  Clear spam folder:
    python3 mail.py --clear spam

── BULK DELETE ─────────────────────────────────────────────────────
  Dry-run (preview matches, no deletion):
    python3 mail.py --bulk-delete --subject-has "promo" --dry-run

  Delete by subject keyword:
    python3 mail.py --bulk-delete --subject-has "sale offer"

  Delete by multiple subject keywords (OR):
    python3 mail.py --bulk-delete --subject-has "promo" "sale" "offer"

  Delete by sender:
    python3 mail.py --bulk-delete --from-addr "noreply@spam.com"

  Delete by sender OR subject (match-any):
    python3 mail.py --bulk-delete \\
        --from-addr "noreply@spam.com" --subject-has "promo" \\
        --match-any

  Delete emails older than 30 days:
    python3 mail.py --bulk-delete --older-than 30

  Delete emails before a date:
    python3 mail.py --bulk-delete --before 2024-01-01

  Delete from Spam folder:
    python3 mail.py --bulk-delete --folder Spam --subject-has "promo"

  Delete unread only:
    python3 mail.py --bulk-delete --subject-has "promo" --unread-only

  Faster scan with parallel fetch:
    python3 mail.py --bulk-delete --subject-has "promo" --parallel

  Auto-confirm (no prompt):
    python3 mail.py --bulk-delete --subject-has "promo" --yes

── REPEAT (scheduled / unattended) ────────────────────────────────
  Run every 5 minutes until Ctrl-C:
    python3 mail.py --bulk-delete --subject-has "promo" \\
        --repeat --yes

  Run every 10 minutes:
    python3 mail.py --bulk-delete --subject-has "promo" \\
        --repeat --interval 600 --yes

  Dry-run repeat (monitor matches without deleting):
    python3 mail.py --bulk-delete --subject-has "promo" \\
        --repeat --interval 120 --dry-run --yes

── HISTORY ─────────────────────────────────────────────────────────
  Show past bulk-delete runs:
    python3 mail.py --history

  Re-run command #3 from history:
    python3 mail.py --replay 3

  Export history as Yahoo Mail filter rules:
    python3 mail.py --export-filters
"""
    print(examples)


def main():
    ap = argparse.ArgumentParser(description="Yahoo Mail manager via IMAP")
    ap.add_argument("--inbox",     action="store_true", help="List inbox")
    ap.add_argument("--unread",    action="store_true", help="List unread emails only")
    ap.add_argument("--folders",   action="store_true", help="List all folders")
    ap.add_argument("--folder",    default="INBOX",     help="Folder to use (default: INBOX)")
    ap.add_argument("--limit",     type=int, default=20,help="Max emails to list (default: 20)")
    ap.add_argument("--read",      metavar="UID",       help="Read email by UID")
    ap.add_argument("--search",    metavar="QUERY",     help="Search emails (from:X / subject:X / keyword)")
    ap.add_argument("--save",      metavar="UID",       help="Save email to mails/ directory")
    ap.add_argument("--save-all",  action="store_true", help="Save all inbox emails locally")
    ap.add_argument("--delete",    metavar="UID",       help="Delete email by UID")
    ap.add_argument("--mark-read", metavar="UID",       help="Mark email as read")
    ap.add_argument("--send",        action="store_true", help="Send an email")
    ap.add_argument("--to",          metavar="ADDRESS",   help="Recipient email address")
    ap.add_argument("--subject",     metavar="SUBJECT",   help="Email subject")
    ap.add_argument("--body",        metavar="BODY",      help="Email body text")
    ap.add_argument("--html",          action="store_true", help="Send body as HTML")
    ap.add_argument("--attach",        metavar="FILE", nargs="+", help="File(s) to attach")
    ap.add_argument("--delete-sent",   action="store_true", help="Prompt to delete the copy from Sent folder after sending")
    ap.add_argument("--clear",         metavar="FOLDER", choices=["sent", "trash", "spam"],
                                       help="Clear all emails from: sent, trash, or spam")
    ap.add_argument("--template",      metavar="NAME",  help="Use a named template from email_templates.yaml")
    ap.add_argument("--template-file", metavar="PATH",  help="Path to templates YAML (default: email_templates.yaml)")
    ap.add_argument("--list-templates",action="store_true", help="List all saved email templates")

    # bulk-delete flags
    ap.add_argument("--list-senders", action="store_true", help="List unique From addresses with email count")
    ap.add_argument("--sort-by",      metavar="FIELD", choices=["count", "addr"], default="count",
                                      help="Sort --list-senders by 'count' (default) or 'addr'")
    ap.add_argument("--bulk-delete",  action="store_true", help="Bulk delete emails by filter")
    ap.add_argument("--from-addr",    metavar="SENDER", nargs="+", help="Delete emails from these senders (any match, substring)")
    ap.add_argument("--subject-has",  metavar="TEXT", nargs="+", help="Delete emails whose subject contains any of these (OR within)")
    ap.add_argument("--body-has",     metavar="TEXT",      help="Delete emails whose body contains TEXT")
    ap.add_argument("--older-than",   metavar="DAYS",      type=int, help="Delete emails older than N days")
    ap.add_argument("--before",       metavar="YYYY-MM-DD",help="Delete emails sent before this date")
    ap.add_argument("--unread-only",  action="store_true", help="Restrict bulk-delete to unread emails")
    ap.add_argument("--dry-run",      action="store_true", help="Preview matches without deleting")
    ap.add_argument("--parallel",     action="store_true", help="Parallel body fetch (faster with --body-has)")
    ap.add_argument("--match-any",    action="store_true", help="OR logic: delete if any of from-addr/subject-has/body-has matches (default is AND)")
    ap.add_argument("--history",        action="store_true", help="Show bulk-delete history")
    ap.add_argument("--replay",         metavar="N", type=int, help="Re-run bulk-delete command #N from history")
    ap.add_argument("--export-filters", action="store_true", help="Export history as Yahoo Mail filter instructions")
    ap.add_argument("--repeat",         action="store_true", help="Re-run bulk-delete on an interval until Ctrl-C")
    ap.add_argument("--interval",       metavar="SECONDS", type=int, default=300,
                                        help="Seconds between repeats (default: 300). Requires --repeat.")
    ap.add_argument("--yes",            action="store_true", help="Auto-confirm deletion (required with --repeat)")
    ap.add_argument("--examples",       action="store_true", help="Show usage examples and exit")

    args = ap.parse_args()

    if args.examples:
        print_examples()
    elif args.list_senders:
        list_senders(folder=args.folder, limit=args.limit, sort_by=args.sort_by)
    elif args.export_filters:
        export_filters()
    elif args.clear:
        clear_system_folder(args.clear)
    elif args.list_templates:
        list_templates(args.template_file)
    elif args.history:
        show_history()
    elif args.replay:
        replay_history(args.replay)
    elif args.folders:
        list_folders()
    elif args.unread:
        list_inbox(folder=args.folder, limit=args.limit, unread_only=True)
    elif args.inbox:
        list_inbox(folder=args.folder, limit=args.limit)
    elif args.read:
        read_email(args.read, folder=args.folder)
    elif args.search:
        search_emails(args.search, folder=args.folder, limit=args.limit)
    elif args.save:
        save_email(args.save, folder=args.folder)
    elif args.save_all:
        save_all(folder=args.folder, limit=args.limit)
    elif args.delete:
        delete_email(args.delete, folder=args.folder)
    elif args.mark_read:
        mark_read(args.mark_read, folder=args.folder)
    elif args.send:
        # Load template first, then let CLI args override
        subject     = args.subject
        body        = args.body
        html        = args.html
        attachments = args.attach or []

        if args.template:
            tpl         = load_template(args.template, args.template_file)
            subject     = subject     or tpl.get("subject", "")
            body        = body        or tpl.get("body", "")
            html        = html        or tpl.get("html", False)
            attachments = attachments or [os.path.expanduser(a) for a in (tpl.get("attachments") or [])]

        if not args.to:
            print("Error: --send requires --to")
            sys.exit(1)
        if not subject:
            print("Error: no subject — provide --subject or define it in the template")
            sys.exit(1)
        if not body:
            print("Error: no body — provide --body or define it in the template")
            sys.exit(1)

        send_email(args.to, subject, body, attachments=attachments or None,
                   html=html, delete_sent=args.delete_sent)
    elif args.bulk_delete:
        if not any([args.from_addr, args.subject_has, args.body_has,
                    args.older_than, args.before, args.unread_only]):
            print("Error: --bulk-delete requires at least one filter flag.")
            print("  Options: --from-addr, --subject-has, --body-has, --older-than, --before, --unread-only")
            sys.exit(1)
        if args.repeat and not args.yes:
            print("Error: --repeat requires --yes (auto-confirm) to run unattended.")
            sys.exit(1)

        run = 0
        while True:
            run += 1
            if args.repeat:
                print(f"\n{'─'*50}")
                print(f"  [repeat] run #{run}  —  {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                print(f"{'─'*50}")
            bulk_delete(
                folder       = args.folder,
                from_addr    = args.from_addr,
                subject_has  = args.subject_has,
                body_has     = args.body_has,
                older_than   = args.older_than,
                before       = args.before,
                unread_only  = args.unread_only,
                dry_run      = args.dry_run,
                parallel     = args.parallel,
                match_any    = args.match_any,
                auto_confirm = args.yes,
            )
            if not args.repeat:
                break
            print(f"\n  [repeat] next run in {args.interval}s — Ctrl-C to stop")
            try:
                time.sleep(args.interval)
            except KeyboardInterrupt:
                print("\n  [repeat] stopped.")
                break
    else:
        ap.print_help()


if __name__ == "__main__":
    main()
