"""
utils/common.py — Shared utilities for the ACBOS/MPDT Pipeline V2.
"""
from __future__ import annotations

import json
import pickle
import logging
import os
import re
import shutil
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd


# ---------------------------------------------------------------------------
# Workspace & Config
# ---------------------------------------------------------------------------

def resolve_workspace(workspace: str | None = None) -> Path:
    if workspace:
        return Path(workspace).expanduser().resolve()
    return Path(__file__).resolve().parents[1]


def load_config(workspace: Path) -> dict[str, Any]:
    cfg_path = workspace / "config" / "pipeline_config.json"
    if not cfg_path.exists():
        raise FileNotFoundError(f"Config not found: {cfg_path}")
    with cfg_path.open("r", encoding="utf-8") as f:
        return json.load(f)


# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

def setup_logger(
    workspace: Path,
    name: str,
    level: str = "INFO",
    max_log_files: int = 10,
) -> logging.Logger:
    log_dir = workspace / "Logs"
    log_dir.mkdir(parents=True, exist_ok=True)

    # Trim old logs
    existing = sorted(log_dir.glob(f"{name}*.log"), key=lambda p: p.stat().st_mtime)
    while len(existing) >= max_log_files:
        existing.pop(0).unlink(missing_ok=True)

    log_file = log_dir / f"{name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

    logger = logging.getLogger(name)
    if logger.handlers:
        return logger

    logger.setLevel(getattr(logging, level.upper(), logging.INFO))

    class _ShortNameFormatter(logging.Formatter):
        """Show only the last segment of the logger name (e.g. 'step1' not 'pipeline.step1')."""
        def format(self, record: logging.LogRecord) -> str:
            record = logging.makeLogRecord(record.__dict__)
            record.name = record.name.rsplit(".", 1)[-1]
            return super().format(record)

    fmt = _ShortNameFormatter(
        "%(asctime)s | %(levelname)-8s | %(name)-8s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(fmt)
    logger.addHandler(sh)

    fh = logging.FileHandler(str(log_file), encoding="utf-8")
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    return logger


# ---------------------------------------------------------------------------
# Text normalisation
# ---------------------------------------------------------------------------

_STRIP_RE = re.compile(r"[\s_\-]+")


def normalize_text(value: Any) -> str:
    """Lower-case, collapse whitespace/underscores/hyphens to single space."""
    s = "" if value is None else str(value)
    return _STRIP_RE.sub(" ", s).strip().lower()


def strip_acbos_suffix(name: str) -> str:
    """Remove trailing '-ACBOS' (case-insensitive) from a document name."""
    return re.sub(r"-ACBOS$", "", name.strip(), flags=re.IGNORECASE)


def as_upper_key(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.strip().str.upper()


# ---------------------------------------------------------------------------
# Column helpers
# ---------------------------------------------------------------------------

def pick_first_existing_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    by_norm = {normalize_text(c): c for c in df.columns}
    for cand in candidates:
        col = by_norm.get(normalize_text(cand))
        if col:
            return col
    return None


# ---------------------------------------------------------------------------
# JSON helpers
# ---------------------------------------------------------------------------

def read_json(path: Path, default: dict | None = None) -> dict:
    if not path.exists():
        return default or {}
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, default=str), encoding="utf-8")


# ---------------------------------------------------------------------------
# Path / file helpers
# ---------------------------------------------------------------------------

def timestamped_dir(workspace: Path, prefix: str) -> Path:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    d = workspace / f"{prefix}_{ts}"
    d.mkdir(parents=True, exist_ok=True)
    return d


def get_available_path(path: Path) -> Path:
    """Return path, or path_new, path_new2 … if the original exists/is locked."""
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    counter = 1
    while True:
        candidate = parent / f"{stem}_v{counter}{suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


def sanitize_filename(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "_", name)


# ---------------------------------------------------------------------------
# Table read/write with caching
# ---------------------------------------------------------------------------

def read_table_any(path_no_ext: Path, logger: logging.Logger) -> pd.DataFrame:
    """Read a cached table using the fastest available format.

    Preference order is pickle -> parquet -> xlsx. Pickle is used as a local
    performance cache only; it avoids the very slow openpyxl path on repeated
    runs and does not require pyarrow/fastparquet.
    """
    pickle_path = path_no_ext.with_suffix(".pkl")
    parquet_path = path_no_ext.with_suffix(".parquet")
    xlsx_path = path_no_ext.with_suffix(".xlsx")

    if pickle_path.exists():
        try:
            return pd.read_pickle(pickle_path)
        except Exception as exc:
            logger.warning("Could not read pickle cache %s (%s). Trying parquet/Excel.", pickle_path.name, exc)
    if parquet_path.exists():
        try:
            return pd.read_parquet(parquet_path)
        except Exception as exc:
            logger.warning("Could not read parquet %s (%s). Trying Excel.", parquet_path.name, exc)
    if xlsx_path.exists():
        try:
            df = pd.read_excel(xlsx_path, dtype=str, engine="openpyxl")
            try:
                pickle_path.parent.mkdir(parents=True, exist_ok=True)
                df.to_pickle(pickle_path)
            except Exception as pexc:
                logger.debug("Could not create pickle cache for %s (%s).", xlsx_path.name, pexc)
            return df
        except Exception as exc:
            logger.warning("Could not read Excel %s (%s).", xlsx_path.name, exc)
    return pd.DataFrame()


def write_table_cache(df: pd.DataFrame, path_no_ext: Path, logger: logging.Logger) -> str:
    """Write local caches. Pickle is always written for fast repeat loads;
    Excel remains the human-readable interchange copy; parquet is optional.
    """
    xlsx_path = path_no_ext.with_suffix(".xlsx")
    pickle_path = path_no_ext.with_suffix(".pkl")
    xlsx_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        df.to_pickle(pickle_path)
    except Exception as exc:
        logger.debug("Could not write pickle cache for %s (%s).", path_no_ext.name, exc)

    # Excel cache writes are very slow for large SmartForms/Scope tables and are
    # not needed for normal pipeline execution once a pickle/parquet cache exists.
    # Keep small caches human-readable, but skip huge Excel duplicates by default.
    cell_count = int(getattr(df, "shape", (0, 0))[0]) * int(getattr(df, "shape", (0, 0))[1])
    write_large_excel = os.environ.get("MPDT_WRITE_LARGE_EXCEL_CACHE", "0").strip() in ("1", "true", "yes")
    wrote_xlsx = False
    if cell_count <= 1_000_000 or write_large_excel:
        df.to_excel(xlsx_path, index=False)
        wrote_xlsx = True
    else:
        logger.info("Skipped large Excel cache for %s (%d cells); pickle cache written.", path_no_ext.name, cell_count)
    # Only attempt parquet if an engine is available to avoid noisy pandas ImportErrors.
    def _detect_parquet_engine() -> str | None:
        try:
            import pyarrow  # type: ignore

            return "pyarrow"
        except Exception:
            try:
                import fastparquet  # type: ignore

                return "fastparquet"
            except Exception:
                return None

    engine = _detect_parquet_engine()
    if not engine:
        logger.debug("Parquet engine not found; skipping parquet cache for %s. Excel written.", path_no_ext.name)
        return "pkl+xlsx" if wrote_xlsx else "pkl"

    try:
        parquet_path = path_no_ext.with_suffix(".parquet")
        # pass explicit engine to avoid pandas attempting autodetect and raising long errors
        df.to_parquet(parquet_path, index=False, engine=engine)
        return "pkl+parquet+xlsx" if wrote_xlsx else "pkl+parquet"
    except Exception as exc:
        logger.warning("Failed to write parquet for %s; falling back to Excel. (%s)", path_no_ext.name, str(exc))
        return "pkl+xlsx" if wrote_xlsx else "pkl"


def cache_signature(path: Path) -> dict[str, object]:
    """Return a small signature for cache invalidation."""
    p = path.resolve()
    if not p.exists():
        return {"path": str(p), "missing": True}
    st = p.stat()
    return {"path": str(p), "mtime_ns": st.st_mtime_ns, "size": st.st_size}


def read_df_cache(cache_path: Path, source_sig: dict[str, object], logger: logging.Logger) -> pd.DataFrame | None:
    """Read a DataFrame cache if its source signature still matches."""
    meta_path = cache_path.with_suffix(cache_path.suffix + ".meta.json")
    if not cache_path.exists() or not meta_path.exists():
        return None
    try:
        meta = json.loads(meta_path.read_text(encoding="utf-8"))
        if meta.get("source") != source_sig:
            return None
        return pd.read_pickle(cache_path)
    except Exception as exc:
        logger.debug("Ignoring stale/unreadable cache %s (%s).", cache_path.name, exc)
        return None


def write_df_cache(df: pd.DataFrame, cache_path: Path, source_sig: dict[str, object], logger: logging.Logger) -> None:
    """Write a DataFrame pickle cache with source signature metadata."""
    try:
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        df.to_pickle(cache_path)
        cache_path.with_suffix(cache_path.suffix + ".meta.json").write_text(
            json.dumps({"source": source_sig}, indent=2), encoding="utf-8"
        )
    except Exception as exc:
        logger.debug("Could not write DataFrame cache %s (%s).", cache_path.name, exc)


# ---------------------------------------------------------------------------
# Subprocess helper (for PowerShell scripts)
# ---------------------------------------------------------------------------

def stream_subprocess(
    cmd: list[str],
    cwd: Path,
    logger: logging.Logger,
    prefix: str = "  ",
) -> int:
    env = os.environ.copy()
    env["PYTHONUNBUFFERED"] = "1"
    env["PYTHONUTF8"] = "1"

    proc = subprocess.Popen(
        cmd,
        cwd=str(cwd),
        env=env,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
    )

    def _drain_stderr() -> None:
        assert proc.stderr is not None
        for line in proc.stderr:
            line = line.rstrip()
            if line:
                logger.warning(f"{prefix}STDERR: {line}")

    t = threading.Thread(target=_drain_stderr, daemon=True)
    t.start()

    assert proc.stdout is not None
    for line in proc.stdout:
        line = line.rstrip()
        if line:
            logger.info(f"{prefix}{line}")

    proc.stdout.close()
    t.join(timeout=10)
    proc.wait()
    return proc.returncode
