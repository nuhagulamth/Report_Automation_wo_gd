"""
Google Drive helper for the GitHub Actions run.

Downloads every .xlsx in the Drive folder DRIVE_FOLDER_ID, and after each upload moves the
file into a 'processed' or 'failed' subfolder of that same folder (created if missing).
Because the GitHub runner is wiped after every run, this is what stops the next scheduled
run from uploading the same file twice.

Needs two environment variables:
    GDRIVE_SA_JSON    full contents of the service account's JSON key
    DRIVE_FOLDER_ID   ID of the 'Reports Excel' folder (the part after /folders/ in its URL)
The folder must be shared with the service account's email as Editor.
"""
import io
import json
import os

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
FOLDER = "application/vnd.google-apps.folder"
SCOPES = ["https://www.googleapis.com/auth/drive"]


class DriveSync:
    def __init__(self, creds_info, folder_id):
        creds = service_account.Credentials.from_service_account_info(creds_info, scopes=SCOPES)
        self.svc = build("drive", "v3", credentials=creds, cache_discovery=False)
        self.folder_id = folder_id
        self._subfolders = {}

    @classmethod
    def from_env(cls):
        raw, folder_id = os.environ.get("GDRIVE_SA_JSON"), os.environ.get("DRIVE_FOLDER_ID")
        if not raw or not folder_id:
            raise RuntimeError("Set GDRIVE_SA_JSON and DRIVE_FOLDER_ID (GitHub Secrets) to use --drive")
        return cls(json.loads(raw), folder_id)

    def _list(self, query):
        items, token = [], None
        while True:
            resp = self.svc.files().list(
                q=query,
                fields="nextPageToken, files(id, name)",
                pageToken=token,
                pageSize=100,
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
            ).execute()
            items += resp["files"]
            token = resp.get("nextPageToken")
            if not token:
                return items

    def download_all(self, dest):
        """Download all .xlsx files directly inside the folder. Returns {local_name: drive_file_id}."""
        dest.mkdir(parents=True, exist_ok=True)
        files = self._list(f"'{self.folder_id}' in parents and trashed = false and mimeType = '{XLSX}'")
        ids = {}
        for f in files:
            name = f["name"]
            if name in ids:  # two Drive files with the same name
                name = f"{f['id'][:6]}_{name}"
            request = self.svc.files().get_media(fileId=f["id"], supportsAllDrives=True)
            buf = io.BytesIO()
            downloader = MediaIoBaseDownload(buf, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            (dest / name).write_bytes(buf.getvalue())
            ids[name] = f["id"]
        return ids

    def _subfolder(self, name):
        if name in self._subfolders:
            return self._subfolders[name]
        found = self._list(
            f"'{self.folder_id}' in parents and trashed = false and mimeType = '{FOLDER}' and name = '{name}'"
        )
        if found:
            folder_id = found[0]["id"]
        else:
            folder_id = self.svc.files().create(
                body={"name": name, "mimeType": FOLDER, "parents": [self.folder_id]},
                fields="id",
                supportsAllDrives=True,
            ).execute()["id"]
        self._subfolders[name] = folder_id
        return folder_id

    def move(self, file_id, subfolder):
        """Move a file from the reports folder into its 'processed' or 'failed' subfolder."""
        target = self._subfolder(subfolder)
        self.svc.files().update(
            fileId=file_id,
            addParents=target,
            removeParents=self.folder_id,
            fields="id",
            supportsAllDrives=True,
        ).execute()
