"""IMAP connection wrapper for Proton Mail Bridge."""

import imaplib
import re
import ssl

_CRLF_RE = re.compile(br'\r\n|\r|\n')


class MessageRejected(Exception):
    """Server rejected the message (NO response) — not a connection error."""
    pass


class BridgeIMAP:
    """Manages IMAP connection to Proton Mail Bridge's local interface."""

    def __init__(self):
        self._conn = None
        self._credentials = None

    def connect(self, host, port, user, password):
        """Connect to Proton Mail Bridge via IMAP + STARTTLS.

        Bridge uses self-signed certs, so we disable certificate verification.
        """
        self._credentials = {
            "host": host,
            "port": port,
            "user": user,
            "password": password,
        }
        conn = imaplib.IMAP4(host, port)
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        conn.starttls(ssl_context=ctx)
        conn.login(user, password)
        self._conn = conn

    def reconnect(self):
        """Re-establish connection using stored credentials."""
        if self._credentials is None:
            raise RuntimeError("No stored credentials to reconnect with")
        try:
            self.disconnect()
        except Exception:
            pass
        self.connect(
            self._credentials["host"],
            self._credentials["port"],
            self._credentials["user"],
            self._credentials["password"],
        )

    def list_folders(self):
        """Return list of folder names from the server."""
        status, data = self._conn.list()
        if status != "OK":
            raise RuntimeError(f"LIST failed: {status}")
        folders = []
        for item in data:
            if item is None:
                continue
            line = item.decode("utf-8", errors="replace") if isinstance(item, bytes) else str(item)
            # Format: '(\\flags) "delimiter" "name"' or '(\\flags) "delimiter" name'
            parts = line.rsplit('"', 2)
            if len(parts) >= 2:
                # Last quoted string is the folder name
                name = parts[-2].strip()
                if name:
                    folders.append(name)
            else:
                # Fallback: take everything after the last space
                name = line.rsplit(" ", 1)[-1].strip().strip('"')
                if name:
                    folders.append(name)
        return sorted(folders)

    def append_message(self, folder, flags, date_time, message_bytes):
        """Append a message to the given folder via IMAP APPEND.

        Bypasses imaplib.append() to control the exact wire format —
        Proton Bridge is strict about quoting and date-time presence.

        Args:
            folder: Target mailbox name.
            flags: IMAP flag string, e.g. '(\\Seen \\Flagged)'.
            date_time: Epoch timestamp (float) for IMAP INTERNALDATE.
            message_bytes: Raw email as bytes.
        """
        # Quote the folder name (required for names with spaces/special chars)
        quoted_folder = '"' + folder.replace('\\', '\\\\').replace('"', '\\"') + '"'
        # Format date — always provide one (Bridge requires it)
        date_str = imaplib.Time2Internaldate(date_time)
        # Normalize line endings to CRLF
        literal = _CRLF_RE.sub(b'\r\n', message_bytes)
        # Set literal for _command to pick up, then issue the command
        self._conn.literal = literal
        try:
            status, resp = self._conn._simple_command(
                'APPEND', quoted_folder, flags, date_str
            )
        except self._conn.error as e:
            # BAD responses (e.g. invalid RFC5322) are raised by imaplib
            # before we can check status — treat as message rejection
            raise MessageRejected(str(e)) from e
        if status != "OK":
            raise MessageRejected(f"APPEND {status}: {resp}")

    def disconnect(self):
        """Logout and close the connection."""
        if self._conn is not None:
            try:
                self._conn.logout()
            except Exception:
                pass
            self._conn = None
