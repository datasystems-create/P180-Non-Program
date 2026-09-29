"""
sync_staff_photos.py
=====================
Pull staff badge photos from a local/network folder, match each photo's
*name* to the Staff Contact List master sheet, rename it to that person's
email, and upload it into a Google Drive folder that the org-chart app reads.

WHAT IT DOES
  1. Reads every image in SOURCE_DIR.
  2. Parses the staff name from the file name and looks it up in the master
     sheet to find the person's email.
  3. Uploads the image to DRIVE_FOLDER_ID with the file title set to the email
     (exactly the convention the app expects). Anyone who already has a photo
     in the folder is skipped, so each run only uploads new photos. Nothing in
     Drive is ever replaced or deleted.
  4. Photos with no confident name match are skipped and listed in a report
     (flip UPLOAD_UNMATCHED to True if you want them uploaded as-is instead).

RUNS: locally on a machine that can see SOURCE_DIR (e.g. your PC with S: mapped),
      authenticating as your own Google account. Nothing else needs to be hosted.

--------------------------------------------------------------------------
ONE-TIME SETUP
  1. Install the libraries:
       pip install google-api-python-client google-auth google-auth-oauthlib Pillow
     (Pillow is optional; without it, images upload at full size.)
  2. In Google Cloud Console (any project):
       - Enable the "Google Drive API" and "Google Sheets API".
       - APIs & Services > Credentials > Create credentials > OAuth client ID.
       - Application type: "Desktop app".
       - Download the JSON and save it next to this script as client_secret.json.
  3. Run the script once from a terminal; a browser opens for you to sign in
     with the SAME Google account that owns the Drive photo folder. After that a
     token.json is cached and future runs are non-interactive.
--------------------------------------------------------------------------
"""

import os
import re
import sys
import mimetypes
import unicodedata
from collections import defaultdict

# ============================ CONFIG ============================
# Source folder holding the badge photos (named by staff member name).
# Tip: if you ever schedule this, use the UNC path (\\server\share\...) instead
# of the S: drive letter -- mapped letters don't exist in background tasks.
SOURCE_DIR = r"S:\Administrative Tools\Staff Pics"

# Destination Drive folder. NOTE: your app currently reads photos from
# 1IeoJJvFsaPLFKXTAxHP3nb00h-P3FkP9 (PHOTO_FOLDER_ID in Code.gs). To have these
# show up in the app, either set this to that folder OR update the app to read
# the folder below.
DRIVE_FOLDER_ID = "1TfIp3zfl_Tzq_iDbN-H-S-lvIReRFs8Q"

# Master sheet + tab used to translate names -> emails.
SPREADSHEET_ID = "1MEmNcvYkQ-uO5Gj-Bequ2_Y4wJrqPNJ4gG1cBkEjPI0"
SHEET_NAME     = "Staff Contact List"

CLIENT_SECRET_FILE = "client_secret.json"
TOKEN_FILE         = "token.json"

# Behaviour toggles
UPLOAD_UNMATCHED = False   # True = upload photos with no name match, using a
                           #        cleaned original name (app will ignore them
                           #        since they aren't emails).
DRY_RUN          = False   # True = match + report only, upload nothing.
RESIZE_MAX_PX    = 1024    # Downscale longest side to this before upload (0 = off).
                           # Keeps the app's detail panel fast. Needs Pillow.
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp"}

# Tokens that appear in badge file names but aren't part of a person's name.
STOPWORDS = {"badge", "photo", "photos", "headshot", "headshots", "pic", "pics",
             "picture", "pictures", "id", "staff", "copy", "final", "new"}

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets.readonly",
]
# ================================================================


# ---------------------------------------------------------------------------
# Name normalisation + matching
# ---------------------------------------------------------------------------
def norm(s):
    """Lowercase, strip accents/punctuation/parentheticals, collapse spaces."""
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    s = s.lower()
    s = re.sub(r"\(.*?\)", " ", s)      # drop "(2)", "(HR)", etc.
    s = re.sub(r"[^a-z0-9]+", " ", s)   # punctuation/underscores -> space
    return re.sub(r"\s+", " ", s).strip()


def name_tokens_from_filename(base):
    """Turn a photo file's base name into cleaned name tokens.
    Handles 'Last, First' ordering and strips badge/id noise + digit runs."""
    base = base.strip()
    if "," in base:                     # "Reyes, Abigail" -> "Abigail Reyes"
        left, right = base.split(",", 1)
        base = right + " " + left
    toks = norm(base).split()
    toks = [t for t in toks if t not in STOPWORDS and not t.isdigit()]
    return toks


def build_lookups(rows, hdr_idx, col):
    """From sheet rows, build name lookups -> set of emails."""
    by_full = defaultdict(set)          # "abigail reyes" -> {email}
    by_first_last = defaultdict(set)    # ("abigail","reyes") -> {email}
    email_to_name = {}
    sn, pn, em = col["staff"], col["preferred"], col["email"]

    def get(row, i):
        return row[i].strip() if i is not None and i < len(row) else ""

    for row in rows[hdr_idx + 1:]:
        name = get(row, sn)
        email = get(row, em).lower()
        if not name or not email:
            continue
        email_to_name.setdefault(email, name)
        for label in (name, get(row, pn)):
            full = norm(label)
            if not full:
                continue
            by_full[full].add(email)
            t = full.split()
            if len(t) >= 2:
                by_first_last[(t[0], t[-1])].add(email)
    return by_full, by_first_last, email_to_name


def match_email(toks, by_full, by_first_last):
    """Return (email, reason) or (None, reason). Only confident, unique matches."""
    if not toks:
        return None, "empty name"
    full = " ".join(toks)
    if full in by_full:
        cands = by_full[full]
        if len(cands) == 1:
            return next(iter(cands)), "full-name match"
        return None, "ambiguous full-name (%d matches)" % len(cands)
    if len(toks) >= 2:
        key = (toks[0], toks[-1])
        if key in by_first_last:
            cands = by_first_last[key]
            if len(cands) == 1:
                return next(iter(cands)), "first+last match"
            return None, "ambiguous first+last (%d matches)" % len(cands)
    return None, "no match"


# ---------------------------------------------------------------------------
# Google auth + services
# ---------------------------------------------------------------------------
def get_services():
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build

    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CLIENT_SECRET_FILE):
                sys.exit("Missing %s -- see the ONE-TIME SETUP notes at the top." % CLIENT_SECRET_FILE)
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())

    drive = build("drive", "v3", credentials=creds, cache_discovery=False)
    sheets = build("sheets", "v4", credentials=creds, cache_discovery=False)
    return drive, sheets


def read_sheet(sheets):
    """Read the whole tab, locate the header row, map the columns we need."""
    resp = sheets.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range="'%s'" % SHEET_NAME,
        majorDimension="ROWS",
    ).execute()
    rows = resp.get("values", [])
    hdr_idx = None
    for i, r in enumerate(rows[:20]):
        cells = [c.strip() for c in r]
        if "Staff Name" in cells and "Email Address" in cells:
            hdr_idx = i
            break
    if hdr_idx is None:
        sys.exit("Could not find the header row (Staff Name / Email Address) in the sheet.")
    hdr = [c.strip() for c in rows[hdr_idx]]

    def idx(label):
        return hdr.index(label) if label in hdr else None

    col = {"staff": idx("Staff Name"), "preferred": idx("Preferred Name"), "email": idx("Email Address")}
    if col["staff"] is None or col["email"] is None:
        sys.exit("Sheet is missing 'Staff Name' or 'Email Address'.")
    return rows, hdr_idx, col


# ---------------------------------------------------------------------------
# Drive upload (upsert by exact file name = email)
# ---------------------------------------------------------------------------
def load_image_bytes(path):
    """Return (bytes, mimetype), optionally downscaled via Pillow."""
    mt = mimetypes.guess_type(path)[0] or "image/jpeg"
    if RESIZE_MAX_PX and RESIZE_MAX_PX > 0:
        try:
            from PIL import Image, ImageOps
            import io
            img = Image.open(path)
            # Keep the photo's orientation exactly as it appears in the source
            # (applies the camera's EXIF rotation, which re-saving would drop).
            img = ImageOps.exif_transpose(img)
            img = img.convert("RGB") if img.mode not in ("RGB", "L") else img
            w, h = img.size
            if max(w, h) > RESIZE_MAX_PX:
                scale = RESIZE_MAX_PX / float(max(w, h))
                img = img.resize((max(1, int(w * scale)), max(1, int(h * scale))), Image.LANCZOS)
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=88)
            return buf.getvalue(), "image/jpeg"
        except ImportError:
            pass  # Pillow not installed -> fall through to raw bytes
        except Exception as e:
            print("   ! could not process image, uploading original (%s)" % e)
    with open(path, "rb") as f:
        return f.read(), mt

def list_drive_names(drive):
    """Return the set of file names already in DRIVE_FOLDER_ID."""
    names, token = set(), None
    q = "'%s' in parents and trashed = false" % DRIVE_FOLDER_ID
    while True:
        resp = drive.files().list(
            q=q, spaces="drive", fields="nextPageToken, files(name)",
            pageSize=1000, pageToken=token,
            supportsAllDrives=True, includeItemsFromAllDrives=True,
        ).execute()
        names.update(f["name"].lower() for f in resp.get("files", []))
        token = resp.get("nextPageToken")
        if not token:
            return names


def create_photo(drive, title, data, mimetype):
    from googleapiclient.http import MediaInMemoryUpload
    media = MediaInMemoryUpload(data, mimetype=mimetype, resumable=False)
    drive.files().create(
        body={"name": title, "parents": [DRIVE_FOLDER_ID]},
        media_body=media, fields="id", supportsAllDrives=True,
    ).execute()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    if not os.path.isdir(SOURCE_DIR):
        sys.exit("Source folder not found (is the share connected?): %s" % SOURCE_DIR)

    print("Authenticating with Google ...")
    drive, sheets = get_services()

    print("Reading master sheet ...")
    rows, hdr_idx, col = read_sheet(sheets)
    by_full, by_first_last, email_to_name = build_lookups(rows, hdr_idx, col)
    print("  %d staff emails loaded." % len(email_to_name))

    files = sorted(
        f for f in os.listdir(SOURCE_DIR)
        if os.path.splitext(f)[1].lower() in IMAGE_EXTS
        and os.path.isfile(os.path.join(SOURCE_DIR, f))
    )
    print("Found %d image files in %s\n" % (len(files), SOURCE_DIR))

    print("Checking what is already in Drive ...")
    in_drive = list_drive_names(drive)
    print("  %d photos already in the Drive folder.\n" % len(in_drive))

    # ---- pass 1: match names and build the list of NEW photos to upload ----
    skipped, already, unmatched, dup = 0, 0, [], []
    queue = []     # (fname, path, title, who) still needing upload
    claimed = {}   # email -> filename already used this run (dedupe)

    for fname in files:
        base = os.path.splitext(fname)[0]
        path = os.path.join(SOURCE_DIR, fname)
        toks = name_tokens_from_filename(base)
        email, reason = match_email(toks, by_full, by_first_last)

        if not email:
            print("SKIP  %-40s (%s)" % (fname, reason))
            unmatched.append((fname, reason))
            if UPLOAD_UNMATCHED:
                title = norm(base).replace(" ", "_") or "unnamed"
                if title.lower() in in_drive:
                    already += 1
                else:
                    queue.append((fname, path, title, base))
            else:
                skipped += 1
            continue

        if email in claimed:
            print("DUP   %-40s -> %s (already from '%s')" % (fname, email, claimed[email]))
            dup.append((fname, email, claimed[email]))
            continue
        claimed[email] = fname

        if email in in_drive:
            already += 1
            continue
        queue.append((fname, path, email, email_to_name.get(email, "")))

    print("\n%d already in Drive, %d new photo(s) to upload.\n" % (already, len(queue)))

    # ---- pass 2: upload the new photos ----
    uploaded = 0
    for i, (fname, path, title, who) in enumerate(queue, 1):
        remaining = len(queue) - i
        if DRY_RUN:
            print("[%d/%d] Would upload photo for %s (%s) -- %d more after this"
                  % (i, len(queue), who, title, remaining))
            uploaded += 1
            continue
        print("[%d/%d] Uploading photo for %s (%s) -- %d more after this"
              % (i, len(queue), who, title, remaining))
        try:
            data, mt = load_image_bytes(path)
            create_photo(drive, title, data, mt)
            uploaded += 1
        except Exception as e:
            print("   ERROR %s -> %s : %s" % (fname, title, e))
            unmatched.append((fname, "upload error: %s" % e))

    # ---- report ----
    print("\n" + "=" * 60)
    print("Done. %s%d uploaded, %d already in Drive, %d skipped, %d duplicate names."
          % ("(dry run) " if DRY_RUN else "", uploaded, already, skipped, len(dup)))
    if unmatched:
        report = os.path.join(os.path.dirname(os.path.abspath(__file__)), "unmatched_photos.txt")
        with open(report, "w", encoding="utf-8") as f:
            f.write("Photos with no confident match (rename these or add the person to the sheet):\n\n")
            for fn, why in unmatched:
                f.write("%s\t%s\n" % (fn, why))
            if dup:
                f.write("\nDuplicate name collisions (only the first was used):\n")
                for fn, em, first in dup:
                    f.write("%s\t%s\tfirst used: %s\n" % (fn, em, first))
        print("Details for the %d unmatched file(s) written to: %s" % (len(unmatched), report))


if __name__ == "__main__":
    main()