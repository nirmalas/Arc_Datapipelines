"""
Config-driven input/output file sync for local, SharePoint, and Azure Blob.

Pipeline code continues to read and write local Input/ and Output/ files. This
module moves files between those local folders and the configured remote storage
at the start/end of each run.
"""
from __future__ import annotations

import json
import logging
import os
from pathlib import Path
from typing import Any



def _as_bool(value: Any, default: bool = False) -> bool:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {"1", "true", "yes", "y", "on"}


def _join_remote(*parts: str) -> str:
    cleaned = [str(p).strip("/\\") for p in parts if str(p or "").strip("/\\")]
    return "/".join(cleaned)


def _local_path(workspace: Path, rel_path: str) -> Path:
    path = Path(rel_path)
    return path if path.is_absolute() else (workspace / path).resolve()


def _get_secret(cfg: dict[str, Any], key: str, env_key: str | None = None) -> str:
    value = cfg.get(key)
    if value:
        return str(value).strip()
    env_name = cfg.get(env_key or f"{key}_env")
    if env_name:
        return os.getenv(str(env_name).strip(), "").strip()
    return ""


def _default_input_items(cfg: dict[str, Any]) -> list[dict[str, Any]]:
    skip = {"db_cache_dir", "join_results"}
    items: list[dict[str, Any]] = []
    for path_key, rel_path in cfg.get("paths", {}).items():
        if path_key in skip or not str(rel_path or "").strip():
            continue
        item: dict[str, Any] = {"path_key": path_key}
        if path_key.endswith("_dir"):
            item["type"] = "prefix"
        items.append(item)
    return items


def _location_config(cfg: dict[str, Any], key: str) -> dict[str, Any]:
    raw = dict(cfg.get(key) or cfg.get(key[:1].upper() + key[1:]) or {})
    location_type = str(raw.get("type", "local")).strip().lower()
    provider_cfg = raw.get(location_type)
    if isinstance(provider_cfg, dict):
        merged = dict(raw)
        merged.update(provider_cfg)
        return merged
    return raw

def _build_blob_service_client(storage_cfg: dict[str, Any]):
    try:
        from azure.storage.blob import BlobServiceClient
    except ImportError as exc:
        raise RuntimeError("Azure Blob storage requires the azure-storage-blob package.") from exc

    connection_string = _get_secret(storage_cfg, "connection_string", "connection_string_env")
    if connection_string:
        return BlobServiceClient.from_connection_string(connection_string)

    account_url = _get_secret(storage_cfg, "account_url", "account_url_env")
    account_name = _get_secret(storage_cfg, "account_name", "account_name_env")
    if not account_url and account_name:
        account_url = f"https://{account_name}.blob.core.windows.net"
    if not account_url:
        raise ValueError("Blob storage requires account_url or account_name in config.")

    sas_token = _get_secret(storage_cfg, "sas_token", "sas_token_env")
    if sas_token:
        return BlobServiceClient(account_url=account_url, credential=sas_token.lstrip("?"))

    credential_mode = str(storage_cfg.get("credential", "managed_identity")).strip().lower()
    if credential_mode in {"managed_identity", "default_azure_credential", "aad"}:
        try:
            from azure.identity import DefaultAzureCredential, ManagedIdentityCredential
        except ImportError as exc:
            raise RuntimeError("Managed identity blob access requires the azure-identity package.") from exc
        client_id = str(storage_cfg.get("managed_identity_client_id", "")).strip()
        credential = ManagedIdentityCredential(client_id=client_id) if client_id else DefaultAzureCredential()
        return BlobServiceClient(account_url=account_url, credential=credential)

    raise ValueError(f"Unsupported blob credential mode: {credential_mode}")


def _build_sharepoint_client(storage_cfg: dict[str, Any]):
    try:
        from utils.sharepoint_utils import SharePointClient
    except ImportError as exc:
        raise RuntimeError("SharePoint storage requires the Office365-REST-Python-Client package.") from exc

    site_url = str(storage_cfg.get("site_url", "")).strip()
    if not site_url:
        raise ValueError("SharePoint storage requires site_url in config.")

    client_id = _get_secret(storage_cfg, "client_id", "client_id_env")
    client_secret = _get_secret(storage_cfg, "client_secret", "client_secret_env")
    username = _get_secret(storage_cfg, "username", "username_env")
    password = _get_secret(storage_cfg, "password", "password_env")
    return SharePointClient(
        site_url=site_url,
        client_id=client_id or None,
        client_secret=client_secret or None,
        username=username or None,
        password=password or None,
    )


def _download_blob_file(container_client, blob_name: str, local_path: Path) -> bool:
    blob_client = container_client.get_blob_client(blob_name)
    try:
        blob_client.get_blob_properties()
    except Exception:
        return False
    local_path.parent.mkdir(parents=True, exist_ok=True)
    with local_path.open("wb") as handle:
        blob_client.download_blob().readinto(handle)
    return True


def _upload_blob_file(container_client, blob_name: str, local_path: Path) -> None:
    with local_path.open("rb") as handle:
        container_client.upload_blob(name=blob_name, data=handle, overwrite=True)


def _resolve_input_item(workspace: Path, cfg: dict[str, Any], location_cfg: dict[str, Any], item: dict[str, Any]) -> tuple[str, Path, str]:
    paths = cfg.get("paths", {})
    path_key = str(item.get("path_key", "")).strip()
    rel_path = str(item.get("local_path") or paths.get(path_key, "")).strip()
    if not rel_path:
        raise ValueError(f"Input location item is missing local_path/path_key: {item}")

    local_path = _local_path(workspace, rel_path)
    default_remote = rel_path.replace("\\", "/")
    remote_file = str(item.get("remote_path") or item.get("blob") or item.get("sharepoint_file") or "").strip()
    remote_prefix = str(item.get("remote_prefix") or item.get("blob_prefix") or item.get("sharepoint_folder") or "").strip()
    base_path = str(location_cfg.get("base_path", "")).strip()
    item_type = str(item.get("type") or ("prefix" if remote_prefix else "file")).strip().lower()
    remote = _join_remote(base_path, remote_file or remote_prefix or default_remote)
    return item_type, local_path, remote


def sync_inputs(workspace: Path, cfg: dict[str, Any], logger: logging.Logger) -> dict[str, Any]:
    """Download configured input files from SharePoint or Blob to local Input/."""
    location_cfg = _location_config(cfg, "input_location")
    location_type = str(location_cfg.get("type", "local")).strip().lower()
    enabled = _as_bool(location_cfg.get("enabled", location_type != "local"))
    if not enabled or location_type in {"", "local"}:
        return {"enabled": False, "type": "local", "downloaded": 0, "skipped": 0, "missing": []}

    fail_on_missing = _as_bool(location_cfg.get("fail_on_missing", True), True)
    items = location_cfg.get("items") or _default_input_items(cfg)
    result: dict[str, Any] = {"enabled": True, "type": location_type, "downloaded": 0, "skipped": 0, "missing": []}

    logger.info("Input sync enabled: type=%s items=%d", location_type, len(items))

    if location_type == "blob":
        container = str(location_cfg.get("container", "")).strip()
        if not container:
            raise ValueError("input_location.type=blob requires input_location.container.")
        container_client = _build_blob_service_client(location_cfg).get_container_client(container)
        for item in items:
            item_type, local_path, remote = _resolve_input_item(workspace, cfg, location_cfg, item)
            if item_type in {"prefix", "directory", "dir"}:
                prefix = remote.rstrip("/") + "/"
                found = False
                for blob in container_client.list_blobs(name_starts_with=prefix):
                    blob_name = str(blob.name)
                    if blob_name.endswith("/"):
                        continue
                    found = True
                    relative = blob_name[len(prefix):].lstrip("/\\")
                    _download_blob_file(container_client, blob_name, local_path / relative)
                    result["downloaded"] += 1
                if not found:
                    result["missing"].append(remote)
                    if fail_on_missing:
                        raise FileNotFoundError(f"Blob prefix not found or empty: {remote}")
            elif _download_blob_file(container_client, remote, local_path):
                logger.info("Downloaded blob input: %s -> %s", remote, local_path)
                result["downloaded"] += 1
            else:
                result["missing"].append(remote)
                if fail_on_missing:
                    raise FileNotFoundError(f"Blob input not found: {remote}")
        return result

    if location_type == "sharepoint":
        sp = _build_sharepoint_client(location_cfg)
        for item in items:
            item_type, local_path, remote = _resolve_input_item(workspace, cfg, location_cfg, item)
            if item_type in {"prefix", "directory", "dir"}:
                names = sp.list_files(remote)
                for name in names:
                    target_path = local_path / name
                    target_path.parent.mkdir(parents=True, exist_ok=True)
                    sp.download_file(_join_remote(remote, name), str(target_path))
                    result["downloaded"] += 1
            else:
                local_path.parent.mkdir(parents=True, exist_ok=True)
                sp.download_file(remote, str(local_path))
                logger.info("Downloaded SharePoint input: %s -> %s", remote, local_path)
                result["downloaded"] += 1
        return result

    raise ValueError(f"Unsupported input_location.type: {location_type}")


def _iter_output_files(output_dir: Path, include_patterns: list[str] | None = None) -> list[Path]:
    if not output_dir.exists():
        return []
    patterns = include_patterns or ["*"]
    files: list[Path] = []
    for pattern in patterns:
        files.extend(p for p in output_dir.rglob(pattern) if p.is_file())
    return sorted(set(files))


def upload_outputs(workspace: Path, cfg: dict[str, Any], output_dir: Path, logger: logging.Logger) -> dict[str, Any]:
    """Upload pipeline outputs from local Output*/ folder to SharePoint or Blob."""
    location_cfg = _location_config(cfg, "output_location")
    location_type = str(location_cfg.get("type", "local")).strip().lower()
    enabled = _as_bool(location_cfg.get("enabled", location_type != "local"))
    if not enabled or location_type in {"", "local"}:
        return {"enabled": False, "type": "local", "uploaded": 0, "errors": []}

    include_patterns = location_cfg.get("include_patterns") or ["*.xlsx", "*.xlsm", "*.xls", "*.csv", "*.json", "*.ACBOS"]
    files = _iter_output_files(output_dir, include_patterns)
    result: dict[str, Any] = {"enabled": True, "type": location_type, "uploaded": 0, "files": [], "errors": []}

    logger.info("Output upload enabled: type=%s files=%d", location_type, len(files))

    if location_type == "blob":
        container = str(location_cfg.get("container", "")).strip()
        if not container:
            raise ValueError("output_location.type=blob requires output_location.container.")
        container_client = _build_blob_service_client(location_cfg).get_container_client(container)
        base_path = str(location_cfg.get("base_path", "")).strip()
        for local_path in files:
            try:
                rel = local_path.relative_to(output_dir).as_posix()
                blob_name = _join_remote(base_path, output_dir.name, rel)
                _upload_blob_file(container_client, blob_name, local_path)
                result["uploaded"] += 1
                result["files"].append({"local": str(local_path), "remote": blob_name})
            except Exception as exc:
                result["errors"].append({"local": str(local_path), "error": str(exc)})
                logger.warning("Blob output upload failed for %s: %s", local_path, exc)
        return result

    if location_type == "sharepoint":
        sp = _build_sharepoint_client(location_cfg)
        base_path = str(location_cfg.get("base_path", "")).rstrip("/")
        for local_path in files:
            try:
                rel_parent = local_path.relative_to(output_dir).parent.as_posix()
                remote_folder = _join_remote(base_path, output_dir.name, "" if rel_parent == "." else rel_parent)
                sp.upload_file(str(local_path), remote_folder)
                result["uploaded"] += 1
                result["files"].append({"local": str(local_path), "remote": _join_remote(remote_folder, local_path.name)})
            except Exception as exc:
                result["errors"].append({"local": str(local_path), "error": str(exc)})
                logger.warning("SharePoint output upload failed for %s: %s", local_path, exc)
        return result

    raise ValueError(f"Unsupported output_location.type: {location_type}")


def write_sync_result(workspace: Path, name: str, result: dict[str, Any]) -> None:
    path = workspace / "Output" / name
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(result, indent=2, default=str), encoding="utf-8")


def _workspace_from_here() -> Path:
    return Path(__file__).resolve().parents[1]


def _load_pipeline_config() -> dict[str, Any]:
    cfg_path = _workspace_from_here() / "config" / "pipeline_config.json"
    if not cfg_path.exists():
        return {}
    return json.loads(cfg_path.read_text(encoding="utf-8-sig"))


def _blob_cfg_from_pipeline(cfg: dict[str, Any] | None = None, section: str = "output_location") -> dict[str, Any]:
    cfg = cfg or _load_pipeline_config()
    shared = dict(cfg.get("blob_storage") or {})
    location = _location_config(cfg, section)
    provider = dict(location.get("blob") or {}) if isinstance(location.get("blob"), dict) else {}
    shared.update(provider)
    shared.update({k: v for k, v in location.items() if k not in {"sharepoint", "blob", "items", "include_patterns"}})
    return shared


def _get_blob_client_from_config(container: str, blob_name: str, cfg: dict[str, Any] | None = None, section: str = "output_location"):
    service = _build_blob_service_client(_blob_cfg_from_pipeline(cfg, section))
    return service.get_blob_client(container=container, blob=blob_name)


def delete_older_blobs(container: str, base_name: str, keep: int = 1, cfg: dict[str, Any] | None = None) -> int:
    """Delete older blobs with the same basename elsewhere in a container."""
    if keep < 1:
        keep = 1
    service = _build_blob_service_client(_blob_cfg_from_pipeline(cfg, "output_location"))
    container_client = service.get_container_client(container)
    matches = [
        blob for blob in container_client.list_blobs()
        if Path(str(blob.name)).name.lower() == Path(base_name).name.lower()
    ]
    matches.sort(key=lambda b: str(getattr(b, "last_modified", "") or ""), reverse=True)
    deleted = 0
    for blob in matches[keep:]:
        container_client.delete_blob(blob.name)
        deleted += 1
    return deleted


def upload_to_blob_storage(container: str, blob_name: str, local_file_name: str) -> None:
    """Upload a local file to Blob Storage using pipeline_config.json settings."""
    log = logging.getLogger("blob_storage")
    local_path = Path(local_file_name)
    cfg = _load_pipeline_config()
    blob_cfg = _blob_cfg_from_pipeline(cfg, "output_location")
    log.info("Uploading file %s -> %s/%s", local_path, container, blob_name)
    log.info("Blob account_url=%s credential_env=%s", blob_cfg.get("account_url", ""), blob_cfg.get("sas_token_env", ""))

    client = _get_blob_client_from_config(container, blob_name, cfg, "output_location")
    with local_path.open("rb") as handle:
        client.upload_blob(handle, overwrite=True)

    log.info("Upload complete")
    try:
        deleted = delete_older_blobs(container, local_path.name, keep=1, cfg=cfg)
        if deleted:
            log.info("Pruned %d older blob(s) for %s", deleted, local_path.name)
    except Exception:
        log.debug("Prune old blobs step failed; continuing", exc_info=True)


def upload_to_blob_storage_chunks(container: str, blob_name: str, local_file_name: str) -> None:
    """Upload a local file to Blob Storage using chunked block upload."""
    log = logging.getLogger("blob_storage")
    local_path = Path(local_file_name)
    log.info("Uploading file in chunks %s -> %s/%s", local_path, container, blob_name)

    client = _get_blob_client_from_config(container, blob_name, None, "output_location")
    with local_path.open("rb") as stream:
        client.upload_blob(stream, overwrite=True, length=local_path.stat().st_size, max_concurrency=4)

    log.info("Chunked upload complete")


def download_from_blob_storage(container: str, blob_name: str, local_file_name: str) -> None:
    """Download a blob to a local file using pipeline_config.json settings."""
    log = logging.getLogger("blob_storage")
    local_path = Path(local_file_name)
    local_path.parent.mkdir(parents=True, exist_ok=True)
    log.info("Downloading blob %s/%s -> %s", container, blob_name, local_path)
    client = _get_blob_client_from_config(container, blob_name, None, "input_location")
    with local_path.open("wb") as handle:
        client.download_blob().readinto(handle)
    log.info("Download complete")


def read_blob_bytes(container: str, blob_name: str) -> bytes:
    """Read a blob fully into memory using pipeline_config.json settings."""
    client = _get_blob_client_from_config(container, blob_name, None, "input_location")
    return client.download_blob().readall()