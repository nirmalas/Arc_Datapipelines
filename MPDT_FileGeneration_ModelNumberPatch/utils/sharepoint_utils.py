from __future__ import annotations

import logging
import os
from pathlib import Path
from urllib.parse import unquote, urlparse

from office365.runtime.auth.client_credential import ClientCredential
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext


log = logging.getLogger("sharepoint")


def normalize_site_url(site_url: str) -> str:
    """Return the SharePoint site root from a site root, absolute URL, or sharing URL."""
    raw = str(site_url or "").strip()
    if not raw:
        return raw
    parsed = urlparse(raw)
    if not parsed.scheme or not parsed.netloc:
        return raw.rstrip("/")

    path = unquote(parsed.path or "")
    parts = [p for p in path.split("/") if p]
    if len(parts) >= 2 and parts[0].lower() in {"sites", "teams"}:
        return f"{parsed.scheme}://{parsed.netloc}/{parts[0]}/{parts[1]}"
    return f"{parsed.scheme}://{parsed.netloc}"


def to_server_relative_url(value: str) -> str:
    """Normalize SharePoint paths to server-relative URLs.

    Supports already server-relative paths, absolute SharePoint URLs, and common
    sharing links such as https://tenant/:f:/r/teams/site/Shared Documents/...
    """
    raw = str(value or "").strip()
    if not raw:
        return raw
    parsed = urlparse(raw)
    if parsed.scheme and parsed.netloc:
        path = unquote(parsed.path or "")
    else:
        path = unquote(raw)

    marker = "/:f:/r/"
    lower = path.lower()
    if marker in lower:
        idx = lower.index(marker)
        path = "/" + path[idx + len(marker):].lstrip("/")
    elif "/:x:/r/" in lower:
        idx = lower.index("/:x:/r/")
        path = "/" + path[idx + len("/:x:/r/"):].lstrip("/")
    elif "/:w:/r/" in lower:
        idx = lower.index("/:w:/r/")
        path = "/" + path[idx + len("/:w:/r/"):].lstrip("/")

    if not path.startswith("/"):
        path = "/" + path
    return path.rstrip("/")


def join_server_relative(*parts: str) -> str:
    cleaned = []
    for part in parts:
        text = str(part or "").strip()
        if not text:
            continue
        cleaned.append(text.strip("/\\"))
    return "/" + "/".join(cleaned) if cleaned else ""


class SharePointClient:
    def __init__(
        self,
        site_url: str,
        client_id: str = None,
        client_secret: str = None,
        username: str = None,
        password: str = None,
    ):
        """Initialize SharePoint connection using app or user credentials."""
        self.site_url = normalize_site_url(site_url)
        if client_id and client_secret:
            creds = ClientCredential(client_id, client_secret)
            self.ctx = ClientContext(self.site_url).with_credentials(creds)
        elif username and password:
            creds = UserCredential(username, password)
            self.ctx = ClientContext(self.site_url).with_credentials(creds)
        else:
            raise ValueError("Provide either app credentials or user credentials")

    def download_file(self, sharepoint_file_url: str, local_path: str):
        """Download a SharePoint file to a local path."""
        remote = to_server_relative_url(sharepoint_file_url)
        Path(local_path).parent.mkdir(parents=True, exist_ok=True)
        try:
            file = self.ctx.web.get_file_by_server_relative_url(remote)
            with open(local_path, "wb") as local_file:
                file.download(local_file).execute_query()
            log.info("Downloaded SharePoint file: %s -> %s", remote, local_path)
        except Exception:
            log.exception("Error downloading SharePoint file: %s", remote)
            raise

    def ensure_folder(self, sharepoint_folder_url: str):
        """Ensure a server-relative folder path exists, creating missing segments."""
        folder_url = to_server_relative_url(sharepoint_folder_url)
        try:
            folder = self.ctx.web.get_folder_by_server_relative_url(folder_url)
            folder.get().execute_query()
            return folder
        except Exception:
            pass

        parts = [p for p in folder_url.strip("/").split("/") if p]
        if len(parts) < 3:
            raise ValueError(f"Cannot create top-level SharePoint folder path: {folder_url}")

        current = "/" + "/".join(parts[:3])
        for part in parts[3:]:
            parent = self.ctx.web.get_folder_by_server_relative_url(current)
            try:
                current = f"{current}/{part}"
                folder = self.ctx.web.get_folder_by_server_relative_url(current)
                folder.get().execute_query()
            except Exception:
                parent.folders.add(part).execute_query()
                log.info("Created SharePoint folder: %s", current)
        return self.ctx.web.get_folder_by_server_relative_url(folder_url)

    def upload_file(self, local_file_path: str, sharepoint_folder_url: str):
        """Upload a local file to a SharePoint folder, creating folders as needed."""
        remote_folder = to_server_relative_url(sharepoint_folder_url)
        try:
            folder = self.ensure_folder(remote_folder)
            with open(local_file_path, "rb") as content_file:
                file_content = content_file.read()
            file_name = os.path.basename(local_file_path)
            uploaded = folder.upload_file(file_name, file_content).execute_query()
            uploaded_url = getattr(uploaded, "serverRelativeUrl", None) or uploaded.properties.get("ServerRelativeUrl", "")
            log.info("Uploaded SharePoint file: %s -> %s", local_file_path, uploaded_url or remote_folder)
            return uploaded_url or join_server_relative(remote_folder, file_name)
        except Exception:
            log.exception("Error uploading SharePoint file: %s -> %s", local_file_path, remote_folder)
            raise

    def list_files(self, sharepoint_folder_url: str):
        """List direct files in a SharePoint folder."""
        remote_folder = to_server_relative_url(sharepoint_folder_url)
        try:
            folder = self.ctx.web.get_folder_by_server_relative_url(remote_folder)
            files = folder.files.get().execute_query()
            return [f.properties["Name"] for f in files]
        except Exception:
            log.exception("Error listing SharePoint folder: %s", remote_folder)
            raise
