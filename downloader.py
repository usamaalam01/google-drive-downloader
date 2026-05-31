import io
import os
import random
import sys
import time
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload
from tqdm import tqdm

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
CREDENTIALS_FILE = Path(__file__).parent / "credentials.json"
TOKEN_FILE = Path(__file__).parent / "token.json"
DOWNLOAD_ROOT = Path(__file__).parent.parent  # E:\Google Drive Download\

EXPORT_MIME_MAP = {
    "application/vnd.google-apps.document": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ".docx",
    ),
    "application/vnd.google-apps.spreadsheet": (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".xlsx",
    ),
    "application/vnd.google-apps.presentation": (
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        ".pptx",
    ),
    "application/vnd.google-apps.drawing": ("image/png", ".png"),
    "application/vnd.google-apps.form": ("application/pdf", ".pdf"),
    "application/vnd.google-apps.script": ("application/vnd.google-apps.script+json", ".json"),
    "application/vnd.google-apps.notebook": ("application/x-ipynb+json", ".ipynb"),
}
GOOGLE_APPS_PREFIX = "application/vnd.google-apps."
GOOGLE_APPS_FOLDER = "application/vnd.google-apps.folder"


def authenticate():
    creds = None
    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CREDENTIALS_FILE.exists():
                print(
                    f"ERROR: {CREDENTIALS_FILE} not found.\n"
                    "Download your OAuth credentials from Google Cloud Console and save as credentials.json.",
                    file=sys.stderr,
                )
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
            creds = flow.run_local_server(port=0)
        TOKEN_FILE.write_text(creds.to_json())
    print("Authenticated successfully.")
    return build("drive", "v3", credentials=creds)


def sanitize_name(name: str) -> str:
    illegal = r'\/:*?"<>|'
    result = "".join("_" if c in illegal or ord(c) < 32 else c for c in name)
    result = result.rstrip(". ")
    if not result:
        result = "_"
    return result[:200]


def get_root_id(service):
    return service.files().get(fileId="root", fields="id").execute()["id"]


def list_root_folders(service, root_id):
    """Return list of (id, name) for folders directly under My Drive root."""
    folders = []
    page_token = None
    while True:
        resp = service.files().list(
            q=f"'{root_id}' in parents and mimeType = '{GOOGLE_APPS_FOLDER}' and trashed = false",
            pageSize=1000,
            fields="nextPageToken, files(id, name)",
        ).execute()
        folders.extend(resp.get("files", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return sorted(folders, key=lambda f: f["name"].lower())


def list_root_files(service, root_id):
    """Return list of file dicts directly under My Drive root (non-folders)."""
    files = []
    page_token = None
    while True:
        resp = service.files().list(
            q=f"'{root_id}' in parents and mimeType != '{GOOGLE_APPS_FOLDER}' and trashed = false",
            pageSize=1000,
            fields="nextPageToken, files(id, name, mimeType, size)",
        ).execute()
        files.extend(resp.get("files", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return files


def scan_folder_tree(service, folder_id, folder_path: Path):
    """
    Fetch ALL files under a folder subtree in one paginated scan using
    a Drive corpus query, then resolve paths via parent relationships.
    Much faster than recursive per-folder API calls.
    """
    print(f"  Fetching file list...", end="", flush=True)

    # Collect every non-trashed file whose ancestry includes folder_id.
    # We fetch all files from My Drive and filter to those inside our target.
    id_map = {}      # file_id -> file dict
    children_map = {}  # parent_id -> [child_id, ...]
    page_token = None
    total = 0

    while True:
        resp = service.files().list(
            q="trashed = false",
            pageSize=1000,
            fields="nextPageToken, files(id, name, mimeType, size, parents)",
            spaces="drive",
        ).execute()
        for f in resp.get("files", []):
            id_map[f["id"]] = f
            for p in f.get("parents", []):
                children_map.setdefault(p, []).append(f["id"])
        total += len(resp.get("files", []))
        print(f"\r  Fetching file list... {total} items", end="", flush=True)
        page_token = resp.get("nextPageToken")
        if not page_token:
            break

    print(f"\r  Fetching file list... {total} items. Resolving paths...", flush=True)

    results = []

    def walk(fid, path):
        seen = {}
        for child_id in children_map.get(fid, []):
            child = id_map.get(child_id)
            if child is None:
                continue
            raw = sanitize_name(child["name"])
            if raw in seen:
                seen[raw] += 1
                unique = f"{raw}_({child_id[:8]})"
            else:
                seen[raw] = 1
                unique = raw
            child_path = path / unique
            if child["mimeType"] == GOOGLE_APPS_FOLDER:
                walk(child_id, child_path)
            else:
                results.append((child, child_path))

    walk(folder_id, folder_path)
    return results


def ensure_dir(path: Path):
    path.mkdir(parents=True, exist_ok=True)


def download_file(service, file: dict, local_path: Path, pbar) -> int:
    """Download or export a single file. Returns bytes written (0 if skipped)."""
    mime = file["mimeType"]
    is_google_app = mime.startswith(GOOGLE_APPS_PREFIX) and mime != GOOGLE_APPS_FOLDER

    if is_google_app:
        if mime in EXPORT_MIME_MAP:
            export_mime, ext = EXPORT_MIME_MAP[mime]
        else:
            print(f"\n[SKIP] {file['name']}: unsupported Google Apps type ({mime})", file=sys.stderr)
            return 0

        if not local_path.suffix or local_path.suffix != ext:
            local_path = local_path.with_name(local_path.name + ext)

        if local_path.exists():
            return 0

        try:
            request = service.files().export_media(fileId=file["id"], mimeType=export_mime)
        except HttpError as e:
            if e.resp.status == 403:
                print(f"\n[SKIP] {file['name']}: export not permitted ({e})", file=sys.stderr)
                return 0
            raise
    else:
        if local_path.exists():
            return 0
        request = service.files().get_media(fileId=file["id"])

    part_path = local_path.with_name(local_path.name + ".part")
    ensure_dir(local_path.parent)

    bytes_written = 0
    try:
        with open(part_path, "wb") as fh:
            downloader = MediaIoBaseDownload(fh, request, chunksize=10 * 1024 * 1024)
            done = False
            while not done:
                status, done = downloader.next_chunk()
                if status:
                    chunk_bytes = int(status.total_size * status.progress()) - bytes_written
                    bytes_written += max(chunk_bytes, 0)
                    pbar.set_postfix(MB=f"{bytes_written / 1_048_576:.1f}", refresh=False)
        part_path.rename(local_path)
    except HttpError as e:
        if part_path.exists():
            part_path.unlink()
        if e.resp.status == 403:
            print(f"\n[SKIP] {file['name']}: download not permitted", file=sys.stderr)
            return 0
        raise

    return bytes_written


def download_with_retry(service, file: dict, local_path: Path, pbar, max_retries: int = 5) -> int:
    for attempt in range(max_retries + 1):
        try:
            return download_file(service, file, local_path, pbar)
        except HttpError as e:
            status = e.resp.status
            if status == 429 or status >= 500:
                if attempt == max_retries:
                    break
                wait = (2 ** attempt) + random.uniform(0, 1)
                time.sleep(wait)
                continue
            elif status == 403:
                wait = 60
                print(f"\n[QUOTA] 403 quota error, waiting {wait}s...", file=sys.stderr)
                time.sleep(wait)
                continue
            else:
                print(f"\n[FAIL] {file['name']}: HTTP {status} — {e}", file=sys.stderr)
                return 0
        except Exception as e:
            if attempt == max_retries:
                break
            wait = (2 ** attempt) + random.uniform(0, 1)
            time.sleep(wait)

    print(f"\n[FAIL] {file['name']}: failed after {max_retries} retries", file=sys.stderr)
    return 0


def pick_folders(folders):
    """Interactive prompt to select one or more folders by number."""
    print("\nTop-level folders in My Drive:\n")
    print("  [0] -- Loose files in root (not in any folder)")
    for i, f in enumerate(folders, 1):
        print(f"  [{i}] {f['name']}")
    print()
    raw = input("Enter numbers to download (e.g. 1  or  1,3,5  or  0,2): ").strip()
    chosen = set()
    for part in raw.replace(" ", ",").split(","):
        part = part.strip()
        if not part:
            continue
        try:
            n = int(part)
        except ValueError:
            print(f"Skipping invalid entry: {part!r}")
            continue
        if n < 0 or n > len(folders):
            print(f"Skipping out-of-range number: {n}")
            continue
        chosen.add(n)
    return sorted(chosen)


def run_downloads(service, file_list, label):
    """Download a list of (file_dict, local_path) tuples with a progress bar."""
    print(f"\n--- {label} ---")
    print(f"Files to process: {len(file_list)}")

    downloaded = skipped = failed = 0
    total_bytes = 0

    with tqdm(total=len(file_list), unit="file", dynamic_ncols=True) as pbar:
        for file, local_path in file_list:
            pbar.set_description(file["name"][:40])

            # Quick skip check
            mime = file["mimeType"]
            is_google_app = mime.startswith(GOOGLE_APPS_PREFIX) and mime != GOOGLE_APPS_FOLDER
            check_path = local_path
            if is_google_app:
                ext = EXPORT_MIME_MAP.get(mime, ("application/pdf", ".pdf"))[1]
                if not check_path.suffix or check_path.suffix != ext:
                    check_path = check_path.with_name(check_path.name + ext)
            if check_path.exists():
                skipped += 1
                pbar.update(1)
                continue

            result = download_with_retry(service, file, local_path, pbar)
            if result == 0 and not check_path.exists():
                failed += 1
            elif result == 0:
                skipped += 1
            else:
                downloaded += 1
                total_bytes += result
            pbar.update(1)

    print(
        f"Done. Downloaded: {downloaded}  |  Skipped (exists): {skipped}  |  Failed: {failed}  "
        f"|  Data: {total_bytes / 1_048_576:.1f} MB"
    )


def main():
    service = authenticate()
    root_id = get_root_id(service)

    print("Fetching top-level folder list...")
    folders = list_root_folders(service, root_id)

    choices = pick_folders(folders)
    if not choices:
        print("Nothing selected, exiting.")
        return

    for choice in choices:
        if choice == 0:
            print("\nScanning loose root files...")
            root_files = list_root_files(service, root_id)
            file_list = [(f, DOWNLOAD_ROOT / sanitize_name(f["name"])) for f in root_files]
            run_downloads(service, file_list, "Root loose files")
        else:
            folder = folders[choice - 1]
            folder_path = DOWNLOAD_ROOT / sanitize_name(folder["name"])
            print(f"\nScanning '{folder['name']}'...")
            file_list = scan_folder_tree(service, folder["id"], folder_path)
            run_downloads(service, file_list, folder["name"])


if __name__ == "__main__":
    main()
