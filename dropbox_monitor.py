"""
Dropbox folder monitor — watches /Acq/FY 26/ for changes and posts to Slack.

Uses Dropbox's list_folder/longpoll endpoint to detect:
- New deal folders created
- New files added to existing deal folders
- Files modified/updated

Posts notifications to Slack when changes occur.
"""

from __future__ import annotations

import json
import os
import threading
import time
from typing import Callable

import requests

import dropbox_client

POLL_INTERVAL = int(os.getenv("DROPBOX_POLL_INTERVAL", "60"))  # seconds


def _get_cursor(path: str) -> str:
    """Get a cursor for the current state of a folder."""
    headers = {**dropbox_client._headers(), "Content-Type": "application/json"}
    resp = requests.post(
        "https://api.dropboxapi.com/2/files/list_folder",
        headers=headers,
        json={"path": path, "recursive": True},
        timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()

    # Consume all pages to get final cursor
    while data.get("has_more"):
        resp = requests.post(
            "https://api.dropboxapi.com/2/files/list_folder/continue",
            headers=headers,
            json={"cursor": data["cursor"]},
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()

    return data["cursor"]


def _get_changes(cursor: str) -> tuple[list[dict], str]:
    """Get changes since cursor. Returns (entries, new_cursor)."""
    headers = {**dropbox_client._headers(), "Content-Type": "application/json"}
    all_entries = []

    resp = requests.post(
        "https://api.dropboxapi.com/2/files/list_folder/continue",
        headers=headers,
        json={"cursor": cursor},
        timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()
    all_entries.extend(data.get("entries", []))

    while data.get("has_more"):
        resp = requests.post(
            "https://api.dropboxapi.com/2/files/list_folder/continue",
            headers=headers,
            json={"cursor": data["cursor"]},
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        all_entries.extend(data.get("entries", []))

    return all_entries, data["cursor"]


def _format_change_notification(entries: list[dict]) -> str | None:
    """Format Dropbox changes into a Slack message."""
    base = dropbox_client.BASE_PATH.lower()
    new_folders = []
    new_files = []
    modified_files = []

    for entry in entries:
        tag = entry.get(".tag", "")
        path = entry.get("path_display", "")
        name = entry.get("name", "")

        # Only care about items in our deal folder
        if not path.lower().startswith(base):
            continue

        # Get relative path after /Acq/FY 26/
        rel = path[len(dropbox_client.BASE_PATH):].strip("/")
        parts = rel.split("/")

        if tag == "folder" and len(parts) == 1:
            # New deal folder created
            new_folders.append(name)
        elif tag == "file":
            # Determine if this is a deal subfolder file
            if len(parts) >= 2:
                deal_name = parts[0]
                file_name = parts[-1]
                new_files.append((deal_name, file_name))

    if not new_folders and not new_files:
        return None

    lines = ["📂 *Dropbox Update — Acq/FY 26*"]

    if new_folders:
        for folder in new_folders:
            lines.append(f"  📁 New deal folder: *{folder}*")

    if new_files:
        # Group by deal
        by_deal: dict[str, list[str]] = {}
        for deal, fname in new_files:
            by_deal.setdefault(deal, []).append(fname)

        for deal, files in by_deal.items():
            if len(files) == 1:
                lines.append(f"  📄 *{deal}*: new file — {files[0]}")
            else:
                lines.append(f"  📄 *{deal}*: {len(files)} new files")
                for f in files[:5]:
                    lines.append(f"      • {f}")
                if len(files) > 5:
                    lines.append(f"      • ...and {len(files) - 5} more")

    return "\n".join(lines)


def start_monitoring(slack_say: Callable[[str], None]):
    """Start background thread that monitors Dropbox for changes."""

    def _monitor_loop():
        cursor = None
        while True:
            try:
                if cursor is None:
                    # Get initial cursor (snapshot of current state)
                    cursor = _get_cursor(dropbox_client.BASE_PATH)
                    print(f"[DROPBOX MONITOR] Initial cursor acquired")
                    time.sleep(POLL_INTERVAL)
                    continue

                entries, new_cursor = _get_changes(cursor)
                cursor = new_cursor

                if entries:
                    msg = _format_change_notification(entries)
                    if msg:
                        slack_say(msg)
                        print(f"[DROPBOX MONITOR] Notified: {len(entries)} changes")

            except Exception as e:
                print(f"[DROPBOX MONITOR] Error: {e}")
                cursor = None  # Reset on error

            time.sleep(POLL_INTERVAL)

    thread = threading.Thread(target=_monitor_loop, daemon=True, name="dropbox-monitor")
    thread.start()
    print(f"[DROPBOX MONITOR] Started — polling every {POLL_INTERVAL}s")
    return thread
