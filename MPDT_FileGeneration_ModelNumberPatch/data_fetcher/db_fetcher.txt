"""
data_fetcher/db_fetcher.py — Fetch data from L3 Database (AssetsScope2/3 tables).

Can be run standalone:
  python -m data_fetcher.db_fetcher
"""
from __future__ import annotations

import argparse
import logging
from pathlib import Path

import pandas as pd

from utils.common import (
    load_config,
    resolve_workspace,
    setup_logger,
    write_table_cache,
    read_table_any,
)


def get_db_connection(cfg: dict):
    """Create pyodbc connection from config."""
    import pyodbc

    conn_str = cfg.get("db", {}).get("connection_string", "")
    if not conn_str or "<server>" in conn_str:
        raise ValueError(
            "Database connection string not configured. "
            "Edit config/pipeline_config.json → db.connection_string"
        )
    return pyodbc.connect(conn_str, timeout=60)


def fetch_table(
    cfg: dict,
    table_name: str,
    cache_dir: Path,
    use_cache: bool,
    logger: logging.Logger,
    extra_fallbacks: list[Path] | None = None,
) -> pd.DataFrame:
    """
    Fetch a single table from L3 database.
    Uses cache if available and use_cache=True.
    Falls back to cached Excel if DB connection fails.
    """
    cache_base = cache_dir / f"{table_name}_full"

    if use_cache:
        cached = read_table_any(cache_base, logger)
        if not cached.empty:
            logger.info("Using cached %s: %d rows", table_name, len(cached))
            return cached

    logger.info("Querying L3 database table: %s ...", table_name)
    try:
        conn = get_db_connection(cfg)
        try:
            df = pd.read_sql(f"SELECT * FROM dbo.[{table_name}]", conn)
        finally:
            conn.close()
    except Exception as exc:
        # Try extra fallbacks (e.g. l3_assets_scope_data) first, then generic DB cache.
        fallback_paths: list[Path] = list(extra_fallbacks or []) + [cache_base.with_suffix(".xlsx")]
        for fb in fallback_paths:
            if fb.exists():
                logger.warning(
                    "DB query failed for %s (%s). Loading from: %s",
                    table_name, exc, fb.name,
                )
                return pd.read_excel(str(fb), dtype=str, engine="openpyxl")
        raise RuntimeError(
            f"DB query failed for {table_name} and no local cache found: {exc}"
        ) from exc

    # Write cache
    cache_dir.mkdir(parents=True, exist_ok=True)
    fmt = write_table_cache(df, cache_base, logger)
    logger.info("  %s fetched: %d rows, cached as %s", table_name, len(df), fmt)
    return df


def fetch_all_db_tables(
    workspace: Path, cfg: dict, logger: logging.Logger
) -> dict[str, pd.DataFrame]:
    """Fetch all configured DB tables directly from the database.
    Always fetches fresh data — only called when fetch_external.db=True in config.
    Also writes a backup copy to DB_Cache/ for reference.
    """
    cache_dir = workspace / cfg.get("paths", {}).get("db_cache_dir", "Input/DB_Cache")
    # Always fetch fresh when this function is called (fetch_external.db controls the gate).
    use_cache = False

    tables = cfg.get("db", {}).get("tables", ["AssetsScope2", "AssetsScope3"])
    results = {}
    for table in tables:
        df = fetch_table(cfg, table, cache_dir, use_cache, logger)
        results[table] = df

    return results


# ---------------------------------------------------------------------------
# Standalone entry point
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description="Fetch L3 DB tables to local cache")
    parser.add_argument("--workspace", default=None)
    parser.add_argument("--no-cache", action="store_true", help="Force re-query from DB")
    args = parser.parse_args()

    workspace = resolve_workspace(args.workspace)
    cfg = load_config(workspace)
    if args.no_cache:
        cfg["use_db_cache"] = False
    logger = setup_logger(workspace, "db_fetcher", cfg.get("log_level", "INFO"))

    logger.info("=== DB Fetcher: Caching L3 tables ===")
    results = fetch_all_db_tables(workspace, cfg, logger)
    for name, df in results.items():
        logger.info("  %s: %d rows", name, len(df))
    logger.info("=== DB Fetcher complete ===")


if __name__ == "__main__":
    main()
