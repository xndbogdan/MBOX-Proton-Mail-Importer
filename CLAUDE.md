# MBOX to Proton Mail Bridge Importer

## Overview
GUI utility (Python/tkinter) to import MBOX email archives into Proton Mail via Proton Mail Bridge's local IMAP interface. Supports resume on failure.

## Structure
- `importer.py` — Main app: GUI (`App`), import thread (`MboxImporter`), retry thread (`SkippedRetryImporter`), resume tracking (`ImportState`), skip tracking (`SkippedStore`)
- `imap_client.py` — IMAP connection wrapper (`BridgeIMAP`) for Proton Bridge
- `requirements.txt` — Single dependency: `tkinterdnd2` (drag & drop, optional)

## Runtime files (auto-created in app directory, gitignored)
- `credentials.json` — Saved IMAP connection settings
- `lastState.json` — Import progress for resume capability
- `skipped.json` — Skipped message details with retry support

## Commands
- Run: `python importer.py`
- Install deps: `pip install tkinterdnd2`

## Key details
- Connects via IMAP4 + STARTTLS with `ssl.CERT_NONE` (Bridge uses self-signed certs)
- APPEND command is constructed manually (bypasses `imaplib.append()`) because Proton Bridge is strict about quoting and requires date-time
- Rate limited: 0.5s between messages
- Skips messages >25MB
