"""
Dropbox API client for M&A deal bot.
All deal outputs go to Dropbox/Acq/FY 26/[Deal Name]/

Uses refresh token for auto-renewing access tokens.
Env vars:
  DROPBOX_REFRESH_TOKEN — long-lived refresh token (never expires)
  DROPBOX_APP_KEY — app key from Dropbox App Console
  DROPBOX_APP_SECRET — app secret from Dropbox App Console
"""

from __future__ import annotations

import json
import os
import re
import threading
import time
from typing import Any

import requests

DROPBOX_REFRESH_TOKEN = os.getenv("DROPBOX_REFRESH_TOKEN", "")
DROPBOX_APP_KEY = os.getenv("DROPBOX_APP_KEY", "")
DROPBOX_APP_SECRET = os.getenv("DROPBOX_APP_SECRET", "")
BASE_PATH = "/Acq/FY 26"

# Token cache
_token_lock = threading.Lock()
_cached_token: str = ""
_token_expires_at: float = 0.0


def _get_access_token() -> str:
    """Get a valid access token, refreshing if needed."""
    global _cached_token, _token_expires_at

    with _token_lock:
        # Return cached if still valid (with 5 min buffer)
        if _cached_token and time.time() < _token_expires_at - 300:
            return _cached_token

        # Refresh
        resp = requests.post(
            "https://api.dropboxapi.com/oauth2/token",
            auth=(DROPBOX_APP_KEY, DROPBOX_APP_SECRET),
            data={
                "grant_type": "refresh_token",
                "refresh_token": DROPBOX_REFRESH_TOKEN,
            },
            timeout=15,
        )
        resp.raise_for_status()
        data = resp.json()
        _cached_token = data["access_token"]
        _token_expires_at = time.time() + data.get("expires_in", 14400)
        print(f"[DROPBOX] Token refreshed (expires in {data.get('expires_in', '?')}s)")
        return _cached_token


def _headers() -> dict[str, str]:
    return {"Authorization": f"Bearer {_get_access_token()}"}


def _safe_name(name: str) -> str:
    """Sanitize a company name for use as a folder/file name."""
    name = re.sub(r'[<>:"/\\|?*]', '', name)
    name = name.strip(". ")
    return name


def _deal_folder(deal_name: str) -> str:
    return f"{BASE_PATH}/{_safe_name(deal_name)}"


# ── Dedup: list existing deal folders ─────────────────────────────────────

def list_deal_folders() -> list[str]:
    """List all existing deal folder names in /Acq/FY 26/."""
    folders: list[str] = []
    resp = requests.post(
        "https://api.dropboxapi.com/2/files/list_folder",
        headers={**_headers(), "Content-Type": "application/json"},
        json={"path": BASE_PATH, "recursive": False},
        timeout=15,
    )
    if resp.status_code != 200:
        return folders

    data = resp.json()
    for entry in data.get("entries", []):
        if entry.get(".tag") == "folder":
            folders.append(entry["name"])

    # Handle pagination
    while data.get("has_more"):
        cursor = data["cursor"]
        resp = requests.post(
            "https://api.dropboxapi.com/2/files/list_folder/continue",
            headers={**_headers(), "Content-Type": "application/json"},
            json={"cursor": cursor},
            timeout=15,
        )
        if resp.status_code != 200:
            break
        data = resp.json()
        for entry in data.get("entries", []):
            if entry.get(".tag") == "folder":
                folders.append(entry["name"])

    return folders


def check_duplicate(company_name: str, openai_client) -> dict[str, Any] | None:
    """Check if a company name matches an existing deal folder.

    Uses LLM fuzzy matching against existing folder names.
    Returns {"match": "folder name", "confidence": "high/medium/low"} or None.
    """
    existing = list_deal_folders()
    if not existing:
        return None

    safe = _safe_name(company_name)

    # Exact match (case-insensitive)
    for folder in existing:
        if folder.lower() == safe.lower():
            return {"match": folder, "confidence": "high"}

    # LLM fuzzy match
    prompt = (
        f"I have a new deal for a company called: \"{company_name}\"\n\n"
        f"Here are existing deal folders:\n"
        + "\n".join(f"- {f}" for f in existing)
        + "\n\nDoes the new company name match any existing folder? "
        "Consider abbreviations, project codenames, DBA names, or slight variations.\n\n"
        "Respond with ONLY this JSON (no markdown):\n"
        '{"is_match": true/false, "matched_folder": "folder name or null", '
        '"confidence": "high/medium/low", "reason": "one sentence"}'
    )

    try:
        response = openai_client.chat.completions.create(
            model="gpt-4.1-mini",
            messages=[{"role": "user", "content": prompt}],
            temperature=0,
        )
        text = response.choices[0].message.content.strip()
        if text.startswith("```"):
            text = text.split("\n", 1)[1] if "\n" in text else text[3:]
        if text.endswith("```"):
            text = text[:-3]
        result = json.loads(text.strip())
        if result.get("is_match"):
            return {
                "match": result.get("matched_folder", ""),
                "confidence": result.get("confidence", "low"),
                "reason": result.get("reason", ""),
            }
    except Exception as e:
        print(f"[DROPBOX DEDUP] LLM match error: {e}")

    return None


# ── List files in a deal folder ───────────────────────────────────────────

def list_deal_files(deal_name: str) -> list[dict[str, Any]]:
    """List files in a deal folder. Returns list of {name, path, size}."""
    path = _deal_folder(deal_name)
    resp = requests.post(
        "https://api.dropboxapi.com/2/files/list_folder",
        headers={**_headers(), "Content-Type": "application/json"},
        json={"path": path, "recursive": False},
        timeout=15,
    )
    if resp.status_code != 200:
        return []

    files = []
    for entry in resp.json().get("entries", []):
        if entry.get(".tag") == "file":
            files.append({
                "name": entry["name"],
                "path": entry["path_display"],
                "size": entry.get("size", 0),
            })
    return files


def download_file(path: str) -> bytes | None:
    """Download a file from Dropbox by path."""
    resp = requests.post(
        "https://content.dropboxapi.com/2/files/download",
        headers={
            **_headers(),
            "Dropbox-API-Arg": json.dumps({"path": path}),
        },
        timeout=60,
    )
    if resp.status_code == 200:
        return resp.content
    return None


# ── Folder & file operations ─────────────────────────────────────────────

def ensure_folder(deal_name: str) -> str:
    """Create the deal folder if it doesn't exist. Returns the path."""
    path = _deal_folder(deal_name)
    resp = requests.post(
        "https://api.dropboxapi.com/2/files/create_folder_v2",
        headers={**_headers(), "Content-Type": "application/json"},
        json={"path": path, "autorename": False},
        timeout=15,
    )
    if resp.status_code not in (200, 409):
        try:
            err = resp.json()
            if err.get("error", {}).get(".tag") == "path" and \
               err["error"].get("path", {}).get(".tag") == "conflict":
                return path
        except Exception:
            pass
        resp.raise_for_status()
    return path


def upload_file(deal_name: str, file_name: str, content: bytes) -> dict:
    """Upload a file to the deal folder. Returns Dropbox metadata."""
    folder = ensure_folder(deal_name)
    safe_file = _safe_name(file_name)
    path = f"{folder}/{safe_file}"

    resp = requests.post(
        "https://content.dropboxapi.com/2/files/upload",
        headers={
            **_headers(),
            "Content-Type": "application/octet-stream",
            "Dropbox-API-Arg": json.dumps({
                "path": path,
                "mode": "overwrite",
                "autorename": False,
            }),
        },
        data=content,
        timeout=60,
    )
    resp.raise_for_status()
    return resp.json()


def create_shared_link(path: str) -> str:
    """Get or create a shared link for a file/folder. Returns the URL."""
    resp = requests.post(
        "https://api.dropboxapi.com/2/sharing/create_shared_link_with_settings",
        headers={**_headers(), "Content-Type": "application/json"},
        json={"path": path},
        timeout=15,
    )
    if resp.status_code == 200:
        return resp.json().get("url", "")

    # If already exists, list and return
    if resp.status_code == 409:
        list_resp = requests.post(
            "https://api.dropboxapi.com/2/sharing/list_shared_links",
            headers={**_headers(), "Content-Type": "application/json"},
            json={"path": path, "direct_only": True},
            timeout=15,
        )
        if list_resp.status_code == 200:
            links = list_resp.json().get("links", [])
            if links:
                return links[0].get("url", "")

    return ""


def get_deal_folder_link(deal_name: str) -> str:
    """Get a shared link for the deal folder."""
    folder = _deal_folder(deal_name)
    return create_shared_link(folder)


def upload_deal_outputs(deal_name: str, files: dict[str, bytes]) -> dict[str, str]:
    """Upload multiple files to a deal folder.

    Args:
        deal_name: Company name for the folder
        files: dict of {filename: bytes_content}

    Returns:
        dict of {filename: dropbox_path}
    """
    paths = {}
    for filename, content in files.items():
        meta = upload_file(deal_name, filename, content)
        paths[filename] = meta.get("path_display", "")
    return paths
