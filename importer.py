"""MBOX to Proton Mail Bridge Importer — GUI application."""

import email
import email.policy
import hashlib
import json
import mailbox
import os
import queue
import sys
import threading
import time
import tkinter as tk
from datetime import datetime
from email.mime.text import MIMEText
from email.utils import mktime_tz, parsedate_tz
from tkinter import filedialog, messagebox, ttk
import imaplib

from imap_client import BridgeIMAP, MessageRejected

# Allow large IMAP response lines (some emails are huge)
imaplib._MAXLINE = 10_000_000

APP_DIR = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_FILE = os.path.join(APP_DIR, "credentials.json")
STATE_FILE = os.path.join(APP_DIR, "lastState.json")
SKIPPED_FILE = os.path.join(APP_DIR, "skipped.json")
MAX_MESSAGE_SIZE = 25 * 1024 * 1024  # 25 MB


# ---------------------------------------------------------------------------
# Import state (resume tracking)
# ---------------------------------------------------------------------------

class ImportState:
    """Tracks import progress in lastState.json for resume capability."""

    @staticmethod
    def fingerprint(mbox_path):
        """SHA-256 of first 8KB + file size — detects file changes."""
        size = os.path.getsize(mbox_path)
        with open(mbox_path, "rb") as f:
            head = f.read(8192)
        h = hashlib.sha256(head + str(size).encode()).hexdigest()
        return h, size

    @staticmethod
    def load():
        """Load saved state, or return None if no state file."""
        if not os.path.exists(STATE_FILE):
            return None
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return None

    @staticmethod
    def save(state):
        """Atomically write state to disk."""
        tmp = STATE_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(state, f, indent=2)
        os.replace(tmp, STATE_FILE)

    @staticmethod
    def delete():
        """Remove state file on successful completion."""
        try:
            os.remove(STATE_FILE)
        except OSError:
            pass


# ---------------------------------------------------------------------------
# Skipped message tracking
# ---------------------------------------------------------------------------

class SkippedStore:
    """Persists skipped message details to skipped.json."""

    @staticmethod
    def load():
        if not os.path.exists(SKIPPED_FILE):
            return None
        try:
            with open(SKIPPED_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return None

    @staticmethod
    def save(data):
        tmp = SKIPPED_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        os.replace(tmp, SKIPPED_FILE)

    @staticmethod
    def delete():
        try:
            os.remove(SKIPPED_FILE)
        except OSError:
            pass

    @staticmethod
    def add_entry(mbox_path, fingerprint, target_folder, entry):
        """Append a skip entry, creating the file if needed."""
        data = SkippedStore.load()
        if data is None or data.get("mbox_fingerprint") != fingerprint:
            data = {
                "mbox_path": mbox_path,
                "mbox_fingerprint": fingerprint,
                "target_folder": target_folder,
                "messages": [],
            }
        data["messages"].append(entry)
        SkippedStore.save(data)


# ---------------------------------------------------------------------------
# Retry pipeline — increasingly aggressive fix-up steps
# ---------------------------------------------------------------------------

_SAFE_HEADERS = ("From", "To", "Cc", "Subject", "Date", "Message-ID",
                 "In-Reply-To", "References")


def _parse_msg(raw_bytes):
    """Parse raw bytes into an email.message.Message, or None."""
    try:
        return email.message_from_bytes(raw_bytes, policy=email.policy.compat32)
    except Exception:
        return None


def _copy_headers(src, dst):
    """Copy safe headers from src message to dst message."""
    for hdr in _SAFE_HEADERS:
        val = src.get(hdr)
        if val:
            dst[hdr] = val


def _extract_text(msg):
    """Pull text/plain (or text/html fallback) from a message."""
    if msg.is_multipart():
        for ct in ("text/plain", "text/html"):
            for part in msg.walk():
                if part.get_content_type() == ct:
                    payload = part.get_payload(decode=True)
                    if payload:
                        charset = part.get_content_charset("utf-8")
                        return payload.decode(charset, errors="replace")
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            charset = msg.get_content_charset("utf-8")
            return payload.decode(charset, errors="replace")
    return None


# -- Individual retry steps --

def _step_reserialize(raw_bytes):
    """Step 1: Parse and re-emit. Fixes line folding, boundaries, encoding."""
    msg = _parse_msg(raw_bytes)
    if msg is None:
        return None
    try:
        return msg.as_bytes()
    except Exception:
        return None


def _step_fix_headers(raw_bytes):
    """Step 2: Fix corrupted sender/recipient/date headers."""
    msg = _parse_msg(raw_bytes)
    if msg is None:
        return None
    try:
        changed = False
        # Fix From
        from_val = msg.get("From", "")
        if not from_val or "@" not in from_val:
            del msg["From"]
            msg["From"] = "unknown@unknown.com"
            changed = True
        # Fix To
        to_val = msg.get("To", "")
        if not to_val or "@" not in to_val:
            del msg["To"]
            msg["To"] = "unknown@unknown.com"
            changed = True
        # Fix Date
        date_val = msg.get("Date", "")
        if not date_val or parsedate_tz(date_val) is None:
            del msg["Date"]
            msg["Date"] = datetime.now().strftime("%a, %d %b %Y %H:%M:%S %z") or \
                          "Sat, 01 Jan 2000 00:00:00 +0000"
            changed = True
        if not changed:
            return None  # headers were fine, skip this step
        return msg.as_bytes()
    except Exception:
        return None


def _step_strip_attachments(raw_bytes):
    """Step 3: Keep only text/* parts, remove attachments."""
    msg = _parse_msg(raw_bytes)
    if msg is None or not msg.is_multipart():
        return None
    try:
        text_body = _extract_text(msg)
        if not text_body:
            return None
        new_msg = MIMEText(text_body, "plain", "utf-8")
        _copy_headers(msg, new_msg)
        new_msg["X-Import-Note"] = "Attachments stripped during import (original was corrupted)"
        return new_msg.as_bytes()
    except Exception:
        return None


def _step_flatten_to_text(raw_bytes):
    """Step 4: Extract whatever text we can into a fresh simple message."""
    msg = _parse_msg(raw_bytes)
    if msg is None:
        return None
    try:
        body = _extract_text(msg) or "[Could not extract message body]"
        new_msg = MIMEText(body, "plain", "utf-8")
        _copy_headers(msg, new_msg)
        # Fix headers on the new message too
        if not new_msg.get("From") or "@" not in new_msg.get("From", ""):
            del new_msg["From"]
            new_msg["From"] = "unknown@unknown.com"
        if not new_msg.get("To") or "@" not in new_msg.get("To", ""):
            del new_msg["To"]
            new_msg["To"] = "unknown@unknown.com"
        new_msg["X-Import-Note"] = "Reconstructed from corrupted original; attachments lost"
        return new_msg.as_bytes()
    except Exception:
        return None


def _step_stub_message(raw_bytes):
    """Step 5: Last resort — headers only, placeholder body."""
    msg = _parse_msg(raw_bytes)
    if msg is None:
        # Can't even parse — build from scratch with minimal info
        new_msg = MIMEText("[Message body could not be imported]", "plain", "utf-8")
        new_msg["From"] = "unknown@unknown.com"
        new_msg["To"] = "unknown@unknown.com"
        new_msg["Subject"] = "[Corrupted message]"
        new_msg["Date"] = "Sat, 01 Jan 2000 00:00:00 +0000"
        new_msg["X-Import-Note"] = "Stub message — original was completely unreadable"
        try:
            return new_msg.as_bytes()
        except Exception:
            return None
    try:
        new_msg = MIMEText("[Message body could not be imported]", "plain", "utf-8")
        _copy_headers(msg, new_msg)
        if not new_msg.get("From") or "@" not in new_msg.get("From", ""):
            del new_msg["From"]
            new_msg["From"] = "unknown@unknown.com"
        if not new_msg.get("To") or "@" not in new_msg.get("To", ""):
            del new_msg["To"]
            new_msg["To"] = "unknown@unknown.com"
        new_msg["X-Import-Note"] = "Stub message — original body was corrupted"
        return new_msg.as_bytes()
    except Exception:
        return None


# Ordered pipeline: least destructive → most destructive
RETRY_STEPS = [
    ("reserialize", _step_reserialize),
    ("fix_headers", _step_fix_headers),
    ("strip_attachments", _step_strip_attachments),
    ("flatten_to_text", _step_flatten_to_text),
    ("stub_message", _step_stub_message),
]


# ---------------------------------------------------------------------------
# Background import thread
# ---------------------------------------------------------------------------

class MboxImporter(threading.Thread):
    """Iterates an mbox file and APPENDs messages to IMAP."""

    def __init__(self, imap_client, mbox_path, target_folder, start_index,
                 total_messages, msg_queue, cancel_event):
        super().__init__(daemon=True)
        self.imap = imap_client
        self.mbox_path = mbox_path
        self.target_folder = target_folder
        self.start_index = start_index
        self.total_messages = total_messages
        self.q = msg_queue
        self.cancel_event = cancel_event

    # -- flag conversion helpers --

    @staticmethod
    def _mbox_flags_to_imap(message):
        """Convert mbox Status/X-Status headers to IMAP flags."""
        flags = []
        status = message.get("Status", "")
        xstatus = message.get("X-Status", "")
        if "R" in status:
            flags.append("\\Seen")
        if "A" in xstatus or "A" in status:
            flags.append("\\Answered")
        if "F" in xstatus:
            flags.append("\\Flagged")
        if "D" in xstatus:
            flags.append("\\Deleted")
        if "T" in xstatus:
            flags.append("\\Draft")
        return "(" + " ".join(flags) + ")"

    @staticmethod
    def _parse_date(message):
        """Extract Date header as epoch timestamp for IMAP INTERNALDATE.

        Always returns a float — falls back to current time if the header
        is missing or unparseable (Proton Bridge requires a date).
        """
        date_str = message.get("Date")
        if date_str:
            try:
                parsed = parsedate_tz(date_str)
                if parsed is not None:
                    return mktime_tz(parsed)
            except Exception:
                pass
        return time.time()

    @staticmethod
    def _message_to_bytes(message):
        """Convert mailbox.mboxMessage to bytes, with fallback."""
        try:
            return message.as_bytes()
        except Exception:
            pass
        try:
            return message.as_string().encode("utf-8", errors="replace")
        except Exception:
            return None

    @staticmethod
    def _skip_entry(idx, message, reason, detail="", retryable=True):
        """Build a skip record dict with message metadata."""
        subject = ""
        message_id = ""
        date = ""
        try:
            subject = str(message.get("Subject", ""))[:200]
            message_id = str(message.get("Message-ID", ""))
            date = str(message.get("Date", ""))
        except Exception:
            pass
        return {
            "index": idx,
            "message_id": message_id,
            "subject": subject,
            "date": date,
            "reason": reason,
            "detail": detail[:500],
            "retryable": retryable,
        }

    def run(self):
        mbox = mailbox.mbox(self.mbox_path)
        skipped = []
        imported = 0
        fingerprint, mbox_size = ImportState.fingerprint(self.mbox_path)

        try:
            for idx, message in enumerate(mbox):
                if self.cancel_event.is_set():
                    self.q.put(("cancelled", imported, skipped))
                    return

                if idx < self.start_index:
                    continue

                # Get raw bytes
                raw = self._message_to_bytes(message)
                if raw is None:
                    entry = self._skip_entry(idx, message, "encoding_failure",
                                             "Could not serialize message to bytes",
                                             retryable=False)
                    skipped.append(entry)
                    SkippedStore.add_entry(self.mbox_path, fingerprint,
                                          self.target_folder, entry)
                    self.q.put(("progress", idx, self.total_messages, len(skipped)))
                    continue

                if len(raw) > MAX_MESSAGE_SIZE:
                    entry = self._skip_entry(idx, message, "too_large",
                                             f"Message size {len(raw)} exceeds limit",
                                             retryable=False)
                    skipped.append(entry)
                    SkippedStore.add_entry(self.mbox_path, fingerprint,
                                          self.target_folder, entry)
                    self.q.put(("progress", idx, self.total_messages, len(skipped)))
                    continue

                flags = self._mbox_flags_to_imap(message)
                date = self._parse_date(message)

                # Attempt APPEND — skip rejected messages, retry on connection failure
                try:
                    self.imap.append_message(self.target_folder, flags, date, raw)
                except MessageRejected as e:
                    entry = self._skip_entry(idx, message, "rejected", str(e))
                    skipped.append(entry)
                    SkippedStore.add_entry(self.mbox_path, fingerprint,
                                          self.target_folder, entry)
                    self.q.put(("progress", idx, self.total_messages, len(skipped)))
                    continue
                except Exception:
                    try:
                        self.imap.reconnect()
                        self.imap.append_message(self.target_folder, flags, date, raw)
                    except MessageRejected as e:
                        entry = self._skip_entry(idx, message, "rejected", str(e))
                        skipped.append(entry)
                        SkippedStore.add_entry(self.mbox_path, fingerprint,
                                              self.target_folder, entry)
                        self.q.put(("progress", idx, self.total_messages, len(skipped)))
                        continue
                    except Exception as e:
                        self.q.put(("error", f"Connection lost at message {idx + 1}: {e}"))
                        return

                imported += 1

                # Save progress after each successful append
                ImportState.save({
                    "mbox_path": self.mbox_path,
                    "mbox_fingerprint": fingerprint,
                    "mbox_size": mbox_size,
                    "total_messages": self.total_messages,
                    "last_imported_index": idx,
                    "target_folder": self.target_folder,
                    "timestamp": datetime.now().isoformat(timespec="seconds"),
                })

                self.q.put(("progress", idx, self.total_messages, len(skipped)))

                # Rate limit
                time.sleep(0.5)

        except Exception as e:
            self.q.put(("error", f"Import error: {e}"))
            return

        # Success — clean up state file
        ImportState.delete()
        self.q.put(("done", imported, skipped))


# ---------------------------------------------------------------------------
# Retry thread for skipped messages
# ---------------------------------------------------------------------------

class SkippedRetryImporter(threading.Thread):
    """Retries skipped messages using alternative serialization strategies."""

    def __init__(self, imap_client, mbox_path, target_folder, skipped_entries,
                 msg_queue, cancel_event):
        super().__init__(daemon=True)
        self.imap = imap_client
        self.mbox_path = mbox_path
        self.target_folder = target_folder
        self.entries = [e for e in skipped_entries if e.get("retryable", True)]
        self.q = msg_queue
        self.cancel_event = cancel_event

    def run(self):
        mbox = mailbox.mbox(self.mbox_path)
        recovered = 0
        still_skipped = []
        total = len(self.entries)

        for i, entry in enumerate(self.entries):
            if self.cancel_event.is_set():
                still_skipped.extend(self.entries[i:])
                self.q.put(("retry_done", recovered, still_skipped))
                return

            idx = entry["index"]
            try:
                message = mbox[idx]
            except (KeyError, IndexError):
                entry["retryable"] = False
                entry["detail"] = "Message no longer accessible at this index"
                still_skipped.append(entry)
                self.q.put(("retry_progress", i + 1, total))
                continue

            raw = MboxImporter._message_to_bytes(message)
            if raw is None:
                entry["retryable"] = False
                still_skipped.append(entry)
                self.q.put(("retry_progress", i + 1, total))
                continue

            flags = MboxImporter._mbox_flags_to_imap(message)
            date = MboxImporter._parse_date(message)

            # Try each fix-up step from least to most destructive
            success = False
            tried_steps = []
            for step_name, step_fn in RETRY_STEPS:
                fixed = step_fn(raw)
                if fixed is None:
                    continue
                if len(fixed) > MAX_MESSAGE_SIZE:
                    continue
                tried_steps.append(step_name)
                if self._try_append(fixed, flags, date):
                    recovered += 1
                    entry["recovered_by"] = step_name
                    success = True
                    break

            if not success:
                entry["retryable"] = False
                entry["detail"] = f"All steps failed: {', '.join(tried_steps) or 'none applicable'}"
                still_skipped.append(entry)

            self.q.put(("retry_progress", i + 1, total))
            if success:
                time.sleep(0.5)

        # Update skipped.json
        if still_skipped:
            data = SkippedStore.load()
            if data:
                data["messages"] = still_skipped
                SkippedStore.save(data)
        else:
            SkippedStore.delete()

        self.q.put(("retry_done", recovered, still_skipped))

    def _try_append(self, message_bytes, flags, date):
        """Attempt IMAP APPEND with one reconnect retry. Returns True on success."""
        try:
            self.imap.append_message(self.target_folder, flags, date, message_bytes)
            return True
        except MessageRejected:
            return False
        except Exception:
            try:
                self.imap.reconnect()
                self.imap.append_message(self.target_folder, flags, date, message_bytes)
                return True
            except Exception:
                return False


# ---------------------------------------------------------------------------
# GUI application
# ---------------------------------------------------------------------------

class App:
    """Main tkinter application."""

    def __init__(self, root):
        self.root = root
        self.root.title("MBOX to Proton Mail Importer")
        self.root.resizable(False, False)

        self.imap = BridgeIMAP()
        self.mbox_path = None
        self.total_messages = 0
        self.msg_queue = queue.Queue()
        self.cancel_event = threading.Event()
        self.importer_thread = None
        self.saved_state = None

        self._build_ui()
        self._load_credentials()
        self._set_state("disconnected")

    # -- UI construction --

    def _build_ui(self):
        pad = {"padx": 8, "pady": 4}

        # Connection frame
        conn_frame = ttk.LabelFrame(self.root, text="Connection", padding=8)
        conn_frame.pack(fill="x", **pad)

        row1 = ttk.Frame(conn_frame)
        row1.pack(fill="x", pady=2)
        ttk.Label(row1, text="Host:").pack(side="left")
        self.host_var = tk.StringVar(value="127.0.0.1")
        ttk.Entry(row1, textvariable=self.host_var, width=20).pack(side="left", padx=(4, 12))
        ttk.Label(row1, text="Port:").pack(side="left")
        self.port_var = tk.StringVar(value="1143")
        ttk.Entry(row1, textvariable=self.port_var, width=8).pack(side="left", padx=4)

        row2 = ttk.Frame(conn_frame)
        row2.pack(fill="x", pady=2)
        ttk.Label(row2, text="User:").pack(side="left")
        self.user_var = tk.StringVar()
        ttk.Entry(row2, textvariable=self.user_var, width=20).pack(side="left", padx=(4, 12))
        ttk.Label(row2, text="Pass:").pack(side="left")
        self.pass_var = tk.StringVar()
        ttk.Entry(row2, textvariable=self.pass_var, width=20, show="*").pack(side="left", padx=4)

        row3 = ttk.Frame(conn_frame)
        row3.pack(fill="x", pady=(4, 0))
        self.connect_btn = ttk.Button(row3, text="Connect", command=self._on_connect)
        self.connect_btn.pack(side="right")
        self.conn_status = ttk.Label(row3, text="")
        self.conn_status.pack(side="left")

        # File frame
        file_frame = ttk.LabelFrame(self.root, text="MBOX File", padding=8)
        file_frame.pack(fill="x", **pad)

        self.drop_label = ttk.Label(
            file_frame,
            text="Drop MBOX file here or click Browse",
            anchor="center",
            relief="sunken",
            padding=18,
        )
        self.drop_label.pack(fill="x", pady=(0, 4))

        # Try to register drag-and-drop
        self._setup_dnd(self.drop_label)

        file_btn_row = ttk.Frame(file_frame)
        file_btn_row.pack(fill="x")
        self.browse_btn = ttk.Button(file_btn_row, text="Browse...", command=self._on_browse)
        self.browse_btn.pack(side="left")
        self.file_info = ttk.Label(file_btn_row, text="")
        self.file_info.pack(side="left", padx=8)

        # Import frame
        imp_frame = ttk.LabelFrame(self.root, text="Import", padding=8)
        imp_frame.pack(fill="x", **pad)

        folder_row = ttk.Frame(imp_frame)
        folder_row.pack(fill="x", pady=2)
        ttk.Label(folder_row, text="Import To:").pack(side="left")
        self.folder_var = tk.StringVar()
        self.folder_combo = ttk.Combobox(folder_row, textvariable=self.folder_var,
                                         state="readonly", width=35)
        self.folder_combo.pack(side="left", padx=4)

        prog_row = ttk.Frame(imp_frame)
        prog_row.pack(fill="x", pady=2)
        self.progress_bar = ttk.Progressbar(prog_row, mode="determinate")
        self.progress_bar.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.progress_label = ttk.Label(prog_row, text="0/0")
        self.progress_label.pack(side="left")

        status_row = ttk.Frame(imp_frame)
        status_row.pack(fill="x", pady=2)
        self.status_label = ttk.Label(status_row, text="Ready")
        self.status_label.pack(side="left")
        self.skipped_label = ttk.Label(status_row, text="")
        self.skipped_label.pack(side="left", padx=12)

        btn_row = ttk.Frame(imp_frame)
        btn_row.pack(fill="x", pady=(4, 0))
        self.import_btn = ttk.Button(btn_row, text="Import", command=self._on_import)
        self.import_btn.pack(side="left", padx=(0, 8))
        self.cancel_btn = ttk.Button(btn_row, text="Cancel", command=self._on_cancel)
        self.cancel_btn.pack(side="left", padx=(0, 8))
        self.retry_btn = ttk.Button(btn_row, text="Retry Skipped (0)",
                                    command=self._on_retry)
        self.retry_btn.pack(side="left")

        # Resume info label
        self.resume_label = ttk.Label(imp_frame, text="", foreground="blue")
        self.resume_label.pack(fill="x", pady=(4, 0))

    def _setup_dnd(self, widget):
        """Register drag-and-drop if tkinterdnd2 is available."""
        try:
            widget.drop_target_register("DND_Files")
            widget.dnd_bind("<<Drop>>", self._on_drop)
        except Exception:
            pass

    # -- credentials --

    def _load_credentials(self):
        if not os.path.exists(CREDENTIALS_FILE):
            return
        try:
            with open(CREDENTIALS_FILE, "r", encoding="utf-8") as f:
                creds = json.load(f)
            self.host_var.set(creds.get("host", "127.0.0.1"))
            self.port_var.set(str(creds.get("port", 1143)))
            self.user_var.set(creds.get("username", ""))
            self.pass_var.set(creds.get("password", ""))
        except (json.JSONDecodeError, OSError):
            pass

    def _save_credentials(self):
        creds = {
            "host": self.host_var.get(),
            "port": int(self.port_var.get()),
            "username": self.user_var.get(),
            "password": self.pass_var.get(),
        }
        try:
            with open(CREDENTIALS_FILE, "w", encoding="utf-8") as f:
                json.dump(creds, f, indent=2)
        except OSError:
            pass

    # -- state machine --

    def _set_state(self, state):
        """Enable/disable controls based on current state."""
        self._ui_state = state

        is_disconnected = state == "disconnected"
        is_connected = state in ("connected", "file_loaded")
        is_importing = state == "importing"

        # Connection fields
        conn_state = "normal" if is_disconnected else "disabled"
        for child in self.root.winfo_children():
            if isinstance(child, ttk.LabelFrame) and child.cget("text") == "Connection":
                for frame in child.winfo_children():
                    for widget in frame.winfo_children():
                        if isinstance(widget, ttk.Entry):
                            widget.configure(state=conn_state)

        self.connect_btn.configure(
            state="normal" if is_disconnected else "disabled"
        )
        self.browse_btn.configure(state="normal" if is_connected or state == "file_loaded" else "disabled")
        self.folder_combo.configure(state="readonly" if (is_connected or state == "file_loaded") else "disabled")
        self.import_btn.configure(state="normal" if state == "file_loaded" else "disabled")
        self.cancel_btn.configure(state="normal" if is_importing else "disabled")
        self.retry_btn.configure(state="disabled")

        if state == "file_loaded":
            self._update_retry_button()

    def _update_retry_button(self):
        """Enable retry button if there are retryable skipped messages."""
        if not self.mbox_path:
            self.retry_btn.configure(text="Retry Skipped (0)", state="disabled")
            return
        data = SkippedStore.load()
        if data is None:
            self.retry_btn.configure(text="Retry Skipped (0)", state="disabled")
            return
        try:
            fp, _ = ImportState.fingerprint(self.mbox_path)
        except OSError:
            self.retry_btn.configure(text="Retry Skipped (0)", state="disabled")
            return
        if fp != data.get("mbox_fingerprint"):
            self.retry_btn.configure(text="Retry Skipped (0)", state="disabled")
            return
        retryable = [e for e in data.get("messages", []) if e.get("retryable", True)]
        count = len(retryable)
        self.retry_btn.configure(
            text=f"Retry Skipped ({count})",
            state="normal" if count > 0 else "disabled",
        )

    # -- connection --

    def _on_connect(self):
        host = self.host_var.get().strip()
        port_str = self.port_var.get().strip()
        user = self.user_var.get().strip()
        password = self.pass_var.get().strip()

        if not all([host, port_str, user, password]):
            messagebox.showwarning("Missing fields", "Please fill in all connection fields.")
            return

        try:
            port = int(port_str)
        except ValueError:
            messagebox.showerror("Invalid port", "Port must be a number.")
            return

        self.conn_status.configure(text="Connecting...", foreground="orange")
        self.root.update_idletasks()

        try:
            self.imap.connect(host, port, user, password)
        except imaplib.IMAP4.error as e:
            self.conn_status.configure(text="", foreground="black")
            messagebox.showerror("Auth failed",
                                 f"Login failed: {e}\n\nCheck your Bridge password.")
            return
        except Exception as e:
            self.conn_status.configure(text="", foreground="black")
            messagebox.showerror("Connection error",
                                 f"Could not connect to {host}:{port}\n\n{e}\n\n"
                                 "Is Proton Mail Bridge running?")
            return

        # Populate folders
        try:
            folders = self.imap.list_folders()
        except Exception as e:
            messagebox.showerror("Error", f"Connected but failed to list folders:\n{e}")
            folders = ["INBOX"]

        self.folder_combo["values"] = folders
        if "INBOX" in folders:
            self.folder_var.set("INBOX")
        elif folders:
            self.folder_var.set(folders[0])

        self._save_credentials()
        self.conn_status.configure(text="Connected", foreground="green")
        self._set_state("connected")

        # Check for saved import state
        self._check_resume_state()

    # -- resume --

    def _check_resume_state(self):
        state = ImportState.load()
        if state is None:
            self.resume_label.configure(text="")
            self.saved_state = None
            return

        self.saved_state = state
        last = state.get("last_imported_index", 0)
        total = state.get("total_messages", 0)
        fname = os.path.basename(state.get("mbox_path", "?"))
        self.resume_label.configure(
            text=f"Previous import interrupted at message {last + 1}/{total} of {fname}"
        )

    def _try_resume(self, mbox_path):
        """Check if selected file matches saved state; offer resume if so."""
        if self.saved_state is None:
            return 0  # start from beginning

        fp, size = ImportState.fingerprint(mbox_path)
        if (fp == self.saved_state.get("mbox_fingerprint") and
                size == self.saved_state.get("mbox_size")):
            last = self.saved_state["last_imported_index"]
            total = self.saved_state["total_messages"]
            folder = self.saved_state.get("target_folder", "INBOX")

            # Set the target folder from saved state
            if folder in list(self.folder_combo["values"]):
                self.folder_var.set(folder)

            answer = messagebox.askyesnocancel(
                "Resume Import",
                f"This file has a saved import state.\n\n"
                f"Imported {last + 1} of {total} messages.\n\n"
                f"Yes = Resume from message {last + 2}\n"
                f"No = Start over\n"
                f"Cancel = Do nothing",
            )
            if answer is None:
                return -1  # user cancelled
            if answer:
                return last + 1  # resume
            else:
                ImportState.delete()
                self.saved_state = None
                self.resume_label.configure(text="")
                return 0  # start over
        return 0

    # -- file selection --

    def _on_browse(self):
        path = filedialog.askopenfilename(
            title="Select MBOX file",
            filetypes=[("MBOX files", "*.mbox"), ("All files", "*.*")],
        )
        if path:
            self._load_mbox(path)

    def _on_drop(self, event):
        path = event.data.strip()
        # tkdnd may wrap paths in braces
        if path.startswith("{") and path.endswith("}"):
            path = path[1:-1]
        if os.path.isfile(path):
            self._load_mbox(path)

    def _load_mbox(self, path):
        self.drop_label.configure(text="Counting messages...")
        self.root.update_idletasks()

        try:
            mbox = mailbox.mbox(path)
            count = len(mbox)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open MBOX file:\n{e}")
            self.drop_label.configure(text="Drop MBOX file here or click Browse")
            return

        self.mbox_path = path
        self.total_messages = count
        fname = os.path.basename(path)
        self.drop_label.configure(text=fname)
        self.file_info.configure(text=f"{count} messages")
        self.progress_bar["maximum"] = count
        self.progress_bar["value"] = 0
        self.progress_label.configure(text=f"0/{count}")
        self._set_state("file_loaded")

        # Check resume
        result = self._try_resume(path)
        if result == -1:
            # User cancelled — reset file selection
            self.mbox_path = None
            self.total_messages = 0
            self.drop_label.configure(text="Drop MBOX file here or click Browse")
            self.file_info.configure(text="")
            self._set_state("connected")
            return

        self._resume_index = result
        if result > 0:
            self.progress_bar["value"] = result
            self.progress_label.configure(text=f"{result}/{count}")
            self.status_label.configure(text=f"Will resume from message {result + 1}")

    # -- import --

    def _on_import(self):
        if not self.mbox_path or not self.folder_var.get():
            return

        self.cancel_event.clear()
        start = getattr(self, "_resume_index", 0)

        self.importer_thread = MboxImporter(
            imap_client=self.imap,
            mbox_path=self.mbox_path,
            target_folder=self.folder_var.get(),
            start_index=start,
            total_messages=self.total_messages,
            msg_queue=self.msg_queue,
            cancel_event=self.cancel_event,
        )
        self.importer_thread.start()
        self._set_state("importing")
        self.status_label.configure(text="Importing...")
        self.skipped_label.configure(text="")
        self._poll_queue()

    def _on_cancel(self):
        self.cancel_event.set()
        self.status_label.configure(text="Cancelling...")

    def _on_retry(self):
        """Retry skipped messages with alternative strategies."""
        data = SkippedStore.load()
        if data is None:
            return

        self.cancel_event.clear()
        entries = data.get("messages", [])
        retryable = [e for e in entries if e.get("retryable", True)]

        self.progress_bar["maximum"] = len(retryable)
        self.progress_bar["value"] = 0
        self.progress_label.configure(text=f"0/{len(retryable)}")

        self.importer_thread = SkippedRetryImporter(
            imap_client=self.imap,
            mbox_path=self.mbox_path,
            target_folder=data.get("target_folder", self.folder_var.get()),
            skipped_entries=entries,
            msg_queue=self.msg_queue,
            cancel_event=self.cancel_event,
        )
        self.importer_thread.start()
        self._set_state("importing")
        self.status_label.configure(text="Retrying skipped messages...")
        self.skipped_label.configure(text="")
        self._poll_queue()

    def _poll_queue(self):
        """Process messages from the import thread."""
        try:
            while True:
                msg = self.msg_queue.get_nowait()
                kind = msg[0]

                if kind == "progress":
                    _, idx, total, skipped_count = msg
                    self.progress_bar["value"] = idx + 1
                    self.progress_label.configure(text=f"{idx + 1}/{total}")
                    if skipped_count:
                        self.skipped_label.configure(text=f"Skipped: {skipped_count}")

                elif kind == "retry_progress":
                    _, current, total = msg
                    self.progress_bar["value"] = current
                    self.progress_label.configure(text=f"{current}/{total}")

                elif kind == "done":
                    _, imported, skipped = msg
                    self.status_label.configure(text="Import complete!")
                    self.progress_label.configure(
                        text=f"{self.total_messages}/{self.total_messages}"
                    )
                    self.resume_label.configure(text="")
                    self.saved_state = None
                    skip_msg = f"\nSkipped {len(skipped)} message(s)." if skipped else ""
                    messagebox.showinfo(
                        "Done",
                        f"Successfully imported {imported} message(s).{skip_msg}",
                    )
                    self._set_state("file_loaded")
                    self._resume_index = 0
                    return

                elif kind == "retry_done":
                    _, recovered, still_skipped = msg
                    self.status_label.configure(text="Retry complete!")
                    still_count = len(still_skipped)
                    detail = ""
                    if still_count:
                        detail = (f"\n{still_count} message(s) could not be imported"
                                  " (corrupted beyond repair).")
                    messagebox.showinfo(
                        "Retry Complete",
                        f"Recovered {recovered} message(s).{detail}",
                    )
                    self._set_state("file_loaded")
                    return

                elif kind == "cancelled":
                    _, imported, skipped = msg
                    self.status_label.configure(text="Cancelled")
                    messagebox.showinfo(
                        "Cancelled",
                        f"Import cancelled. {imported} message(s) imported.\n"
                        "Progress saved — you can resume later.",
                    )
                    self._set_state("file_loaded")
                    self._check_resume_state()
                    return

                elif kind == "error":
                    _, error_msg = msg
                    self.status_label.configure(text="Error — stopped")
                    messagebox.showerror(
                        "Import Error",
                        f"{error_msg}\n\nProgress saved — you can resume later.",
                    )
                    self._set_state("file_loaded")
                    self._check_resume_state()
                    return

        except queue.Empty:
            pass

        if self._ui_state == "importing":
            self.root.after(100, self._poll_queue)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    # Try tkinterdnd2 for drag & drop support
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except ImportError:
        root = tk.Tk()

    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
