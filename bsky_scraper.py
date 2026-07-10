"""
Bluesky -> Google Drive + Google Sheets scraper

Scrapes ORIGINAL posts (no reposts) from a single Bluesky username, downloads
images/videos, uploads them to a named Google Drive folder (created if it
doesn't already exist), and logs the filename + caption + hashtags into a
Google Sheet.

All configuration is passed in via environment variables so this can be
driven entirely from a GitHub Actions workflow_dispatch run.

Environment variables:
    BSKY_HANDLE                 - your Bluesky login handle
    BSKY_APP_PASSWORD           - your Bluesky app password (NOT your main password)
    TARGET_USERNAME             - the handle whose posts you want to scrape
    MODE                        - "timeline" or "media"
    CONTENT_TYPE                - "images", "videos", or "both"
    MAX_POSTS                   - how many of the user's posts to scan (default 100)
    HASHTAG_COUNT               - max hashtags to save per post (default 3)
    DRIVE_FOLDER_NAME           - Google Drive folder name to upload into
    DRIVE_ID                    - (optional) Shared Drive ID, recommended - see README
    GOOGLE_APPLICATION_CREDENTIALS - path to the Google OAuth user token JSON file
    GOOGLE_SHEET_ID             - target Google Sheet ID

Google credentials format (GOOGLE_APPLICATION_CREDENTIALS file):
    This is a USER OAUTH TOKEN JSON (not a service account key), shaped like:
    {
        "token": "...",
        "refresh_token": "...",
        "token_uri": "https://oauth2.googleapis.com/token",
        "client_id": "...",
        "client_secret": "...",
        "scopes": ["https://www.googleapis.com/auth/drive",
                    "https://www.googleapis.com/auth/spreadsheets"]
    }
"""

import json
import os
import re
import sys
import time
from datetime import datetime

import requests
from atproto import Client

from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# ---------------------------------------------------------------------------
# Config from environment
# ---------------------------------------------------------------------------
BSKY_HANDLE = os.environ.get("BSKY_HANDLE", "").strip()
BSKY_APP_PASSWORD = os.environ.get("BSKY_APP_PASSWORD", "").strip()
TARGET_USERNAME = os.environ.get("TARGET_USERNAME", "").strip()
MODE = os.environ.get("MODE", "timeline").strip().lower()
CONTENT_TYPE = os.environ.get("CONTENT_TYPE", "both").strip().lower()
MAX_POSTS = int(os.environ.get("MAX_POSTS", "100") or 100)
HASHTAG_COUNT = int(os.environ.get("HASHTAG_COUNT", "3") or 3)
DRIVE_FOLDER_NAME = os.environ.get("DRIVE_FOLDER_NAME", "").strip()
DRIVE_ID = os.environ.get("DRIVE_ID", "").strip() or None
GOOGLE_CREDS_PATH = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "google_creds.json")
GOOGLE_SHEET_ID = os.environ.get("GOOGLE_SHEET_ID", "").strip()

DOWNLOAD_DIR = "downloads"
SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]

SHEET_HEADER = ["File Name", "Type", "Caption", "Hashtags"]


def log(msg):
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)


def fail(msg):
    log(f"FATAL: {msg}")
    sys.exit(1)


# ---------------------------------------------------------------------------
# Google auth / helpers
# ---------------------------------------------------------------------------
def get_google_services():
    if not os.path.exists(GOOGLE_CREDS_PATH):
        fail(f"Google credentials file not found at {GOOGLE_CREDS_PATH}")

    with open(GOOGLE_CREDS_PATH) as f:
        info = json.load(f)

    creds = Credentials(
        token=info.get("token"),
        refresh_token=info.get("refresh_token"),
        token_uri=info.get("token_uri"),
        client_id=info.get("client_id"),
        client_secret=info.get("client_secret"),
        scopes=info.get("scopes", SCOPES),
    )

    # Refresh if expired (or if no access token was stored, just a refresh_token)
    if not creds.valid:
        if creds.refresh_token:
            creds.refresh(Request())
            log("🔄 Refreshed Google OAuth access token")
        else:
            fail("Google credentials are invalid/expired and no refresh_token is present")

    drive_service = build("drive", "v3", credentials=creds)
    sheets_service = build("sheets", "v4", credentials=creds)
    return drive_service, sheets_service


def get_or_create_drive_folder(drive_service, folder_name):
    query = (
        f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder' "
        f"and trashed=false"
    )
    list_kwargs = dict(q=query, spaces="drive", fields="files(id, name)")
    if DRIVE_ID:
        list_kwargs.update(
            corpora="drive",
            driveId=DRIVE_ID,
            includeItemsFromAllDrives=True,
            supportsAllDrives=True,
        )
    results = drive_service.files().list(**list_kwargs).execute()
    items = results.get("files", [])
    if items:
        log(f"📁 Found existing Drive folder '{folder_name}' ({items[0]['id']})")
        return items[0]["id"]

    metadata = {"name": folder_name, "mimeType": "application/vnd.google-apps.folder"}
    if DRIVE_ID:
        metadata["parents"] = [DRIVE_ID]
    create_kwargs = dict(body=metadata, fields="id")
    if DRIVE_ID:
        create_kwargs["supportsAllDrives"] = True
    folder = drive_service.files().create(**create_kwargs).execute()
    log(f"📁 Created new Drive folder '{folder_name}' ({folder['id']})")
    return folder["id"]


def upload_file_to_drive(drive_service, filepath, folder_id):
    metadata = {"name": os.path.basename(filepath), "parents": [folder_id]}
    media = MediaFileUpload(filepath, resumable=True)
    kwargs = dict(body=metadata, media_body=media, fields="id")
    if DRIVE_ID:
        kwargs["supportsAllDrives"] = True
    file = drive_service.files().create(**kwargs).execute()
    return file.get("id")


def get_first_sheet_title(sheets_service, sheet_id):
    meta = sheets_service.spreadsheets().get(spreadsheetId=sheet_id).execute()
    return meta["sheets"][0]["properties"]["title"]


def ensure_sheet_header(sheets_service, sheet_id, sheet_title):
    existing = (
        sheets_service.spreadsheets()
        .values()
        .get(spreadsheetId=sheet_id, range=f"{sheet_title}!A1:D1")
        .execute()
        .get("values", [])
    )
    if not existing:
        sheets_service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=f"{sheet_title}!A1:D1",
            valueInputOption="RAW",
            body={"values": [SHEET_HEADER]},
        ).execute()


def append_sheet_row(sheets_service, sheet_id, sheet_title, row):
    sheets_service.spreadsheets().values().append(
        spreadsheetId=sheet_id,
        range=f"{sheet_title}!A:D",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": [row]},
    ).execute()


# ---------------------------------------------------------------------------
# Bluesky helpers
# ---------------------------------------------------------------------------
def get_video_cid_and_did(post_view):
    try:
        embed = getattr(post_view, "embed", None)
        if not embed:
            return None, None
        did = getattr(post_view.author, "did", None)
        embed_type = getattr(embed, "$type", "") or getattr(embed, "py_type", "") or str(type(embed))
        if "app.bsky.embed.video#view" in embed_type or "video" in embed_type.lower():
            cid = getattr(embed, "cid", None)
            if cid:
                return str(cid), did
        if hasattr(embed, "video") and hasattr(embed.video, "ref"):
            ref = embed.video.ref
            cid = getattr(ref, "$link", None) or (ref.get("$link") if isinstance(ref, dict) else None)
            if cid:
                return str(cid), did
    except Exception:
        pass
    return None, None


def get_image_urls(post_view):
    urls = []
    try:
        embed = getattr(post_view, "embed", None)
        if not embed:
            return urls
        embed_type = getattr(embed, "$type", "") or getattr(embed, "py_type", "") or str(type(embed))

        if "images" in embed_type.lower():
            for img in getattr(embed, "images", []) or []:
                url = getattr(img, "fullsize", None) or getattr(img, "thumb", None)
                if url:
                    urls.append(url)

        media = getattr(embed, "media", None)
        if media:
            media_type = getattr(media, "$type", "") or getattr(media, "py_type", "") or str(type(media))
            if "images" in media_type.lower():
                for img in getattr(media, "images", []) or []:
                    url = getattr(img, "fullsize", None) or getattr(img, "thumb", None)
                    if url:
                        urls.append(url)
    except Exception:
        pass
    return urls


def download_binary(url, filepath, timeout=30):
    try:
        r = requests.get(url, stream=True, timeout=timeout)
        if r.status_code == 200:
            with open(filepath, "wb") as f:
                for chunk in r.iter_content(chunk_size=8192):
                    f.write(chunk)
            return True
    except Exception:
        pass
    return False


def download_video(client, did, cid, filepath):
    if not did or not cid:
        return False
    cdn_url = f"https://cdn.bsky.app/img/feed_fullsize/plain/{did}/{cid}"
    if download_binary(cdn_url, filepath, timeout=45):
        if os.path.getsize(filepath) > 10000:
            return True
        os.remove(filepath)
    try:
        blob_data = client.com.atproto.sync.get_blob(params={"did": did, "cid": cid})
        with open(filepath, "wb") as f:
            f.write(blob_data)
        return True
    except Exception:
        pass
    return False


def extract_hashtags(text, limit):
    tags = re.findall(r"#(\w+)", text or "")
    if limit <= 0:
        return tags
    return tags[:limit]


def is_repost(feed_item):
    """A feed item from get_author_feed represents a repost if it has a
    'reason' of type app.bsky.feed.defs#reasonRepost."""
    reason = getattr(feed_item, "reason", None)
    if not reason:
        return False
    reason_type = getattr(reason, "$type", "") or getattr(reason, "py_type", "") or str(type(reason))
    return "reasonRepost" in reason_type


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    if not BSKY_HANDLE or not BSKY_APP_PASSWORD:
        fail("BSKY_HANDLE / BSKY_APP_PASSWORD are not set")
    if not TARGET_USERNAME:
        fail("TARGET_USERNAME is not set")
    if not DRIVE_FOLDER_NAME:
        fail("DRIVE_FOLDER_NAME is not set")
    if not GOOGLE_SHEET_ID:
        fail("GOOGLE_SHEET_ID is not set")
    if MODE not in ("timeline", "media"):
        fail("MODE must be 'timeline' or 'media'")
    if CONTENT_TYPE not in ("images", "videos", "both"):
        fail("CONTENT_TYPE must be 'images', 'videos', or 'both'")

    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    log("🔑 Setting up Google Drive / Sheets services...")
    drive_service, sheets_service = get_google_services()
    folder_id = get_or_create_drive_folder(drive_service, DRIVE_FOLDER_NAME)
    sheet_title = get_first_sheet_title(sheets_service, GOOGLE_SHEET_ID)
    ensure_sheet_header(sheets_service, GOOGLE_SHEET_ID, sheet_title)

    log(f"🔑 Logging in to Bluesky as {BSKY_HANDLE}...")
    client = Client()
    client.login(BSKY_HANDLE, BSKY_APP_PASSWORD)
    log("✅ Bluesky login successful!")

    try:
        profile = client.app.bsky.actor.get_profile(params={"actor": TARGET_USERNAME})
        target_did = profile.did
    except Exception as e:
        fail(f"Could not resolve target username '{TARGET_USERNAME}': {e}")

    log(f"🔍 Scraping '{TARGET_USERNAME}' | mode={MODE} | content={CONTENT_TYPE} "
        f"| max_posts={MAX_POSTS} | hashtags={HASHTAG_COUNT}")

    feed_filter = "posts_with_media" if MODE == "media" else "posts_no_replies"

    cursor = None
    scanned = 0
    saved_count = 0

    while scanned < MAX_POSTS:
        try:
            resp = client.app.bsky.feed.get_author_feed(
                params={
                    "actor": TARGET_USERNAME,
                    "filter": feed_filter,
                    "limit": 30,
                    "cursor": cursor,
                }
            )
        except Exception as e:
            log(f"⚠️ Feed fetch error: {e}")
            break

        if not resp.feed:
            break

        for item in resp.feed:
            if scanned >= MAX_POSTS:
                break
            scanned += 1

            # Skip reposts entirely - only original posts by this user
            if is_repost(item):
                continue

            post_view = item.post
            author_did = getattr(post_view.author, "did", None)
            if author_did != target_did:
                continue

            record = getattr(post_view, "record", None)
            text = getattr(record, "text", "") or ""
            post_cid = getattr(post_view, "cid", None)
            safe_name = re.sub(r"[^a-zA-Z0-9_-]", "_", post_cid or str(scanned))
            hashtags = extract_hashtags(text, HASHTAG_COUNT)
            hashtags_str = ", ".join(f"#{h}" for h in hashtags)

            saved_files = []

            if CONTENT_TYPE in ("videos", "both"):
                cid, did = get_video_cid_and_did(post_view)
                if cid:
                    fname = f"{safe_name}.mp4"
                    fpath = os.path.join(DOWNLOAD_DIR, fname)
                    log(f"⬇️ Downloading video: {text[:50]}...")
                    if download_video(client, did, cid, fpath):
                        saved_files.append(("video", fpath))
                    else:
                        log(f"  ❌ Failed to download video for post {safe_name}")

            if CONTENT_TYPE in ("images", "both"):
                img_urls = get_image_urls(post_view)
                for i, url in enumerate(img_urls):
                    fname = f"{safe_name}_img{i + 1}.jpg"
                    fpath = os.path.join(DOWNLOAD_DIR, fname)
                    log(f"⬇️ Downloading image {i + 1}: {text[:50]}...")
                    if download_binary(url, fpath):
                        saved_files.append(("image", fpath))
                    else:
                        log(f"  ❌ Failed to download image for post {safe_name}")

            # Upload each saved file to Drive + log a row in Sheets
            for media_type, fpath in saved_files:
                try:
                    upload_file_to_drive(drive_service, fpath, folder_id)
                    append_sheet_row(
                        sheets_service,
                        GOOGLE_SHEET_ID,
                        sheet_title,
                        [os.path.basename(fpath), media_type, text, hashtags_str],
                    )
                    saved_count += 1
                    log(f"  ✅ Uploaded + logged: {os.path.basename(fpath)}")
                except Exception as e:
                    log(f"  ⚠️ Upload/log failed for {fpath}: {e}")
                finally:
                    if os.path.exists(fpath):
                        os.remove(fpath)

        cursor = getattr(resp, "cursor", None)
        if not cursor:
            break
        time.sleep(0.5)

    log(f"🎉 Done! Scanned {scanned} posts, uploaded {saved_count} files to "
        f"Drive folder '{DRIVE_FOLDER_NAME}', logged to sheet '{sheet_title}'.")


if __name__ == "__main__":
    main()
