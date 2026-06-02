"""Load QRI IV snapshots from R2 (preferred) or local cache.

Mirrors the storage layout written by scripts/fetch_qri_iv.py:
  R2 key:    qri_iv/raw/YYYYMMDD/YYYYMMDD_HHMMSS.parquet
  Local:     cache/qri_iv/raw/YYYYMMDD/YYYYMMDD_HHMMSS.parquet

Each parquet is one fetch snapshot (all months × strikes × PUT/CALL).
"""
from __future__ import annotations

import io
import logging
from datetime import date
from pathlib import Path
from typing import Optional

import pandas as pd

import config
from data import r2_storage

logger = logging.getLogger(__name__)

_R2_PREFIX = "qri_iv/raw"
_LOCAL_ROOT = config.CACHE_DIR / "qri_iv" / "raw"


def list_snapshot_keys(day: date) -> list[str]:
    """List R2 keys for a given day, sorted ascending (oldest first)."""
    r2_storage._init_client()
    if r2_storage._client is None:
        return []
    prefix = f"{_R2_PREFIX}/{day.strftime('%Y%m%d')}/"
    try:
        resp = r2_storage._client.list_objects_v2(
            Bucket=r2_storage._bucket, Prefix=prefix
        )
        return sorted(o["Key"] for o in resp.get("Contents", []))
    except Exception as e:
        logger.warning("list_snapshot_keys failed: %s", e)
        return []


def load_snapshot(key: str) -> Optional[pd.DataFrame]:
    """Load one snapshot parquet by full R2 key (with local fallback)."""
    content = r2_storage.r2_get(key)
    if content is None:
        # local fallback: strip prefix → cache path
        rel = key[len(_R2_PREFIX) + 1:] if key.startswith(_R2_PREFIX) else key
        local = _LOCAL_ROOT / rel
        if local.exists():
            content = local.read_bytes()
    if content is None:
        return None
    try:
        return pd.read_parquet(io.BytesIO(content))
    except Exception as e:
        logger.warning("parquet read failed for %s: %s", key, e)
        return None


def load_day(day: date) -> pd.DataFrame:
    """Concatenate all snapshots for a day into one time-series DataFrame.

    Returns empty DataFrame if nothing found.
    """
    keys = list_snapshot_keys(day)
    if not keys:
        # local-only fallback
        local_dir = _LOCAL_ROOT / day.strftime("%Y%m%d")
        if local_dir.exists():
            frames = [pd.read_parquet(p) for p in sorted(local_dir.glob("*.parquet"))]
            return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
        return pd.DataFrame()
    frames = []
    for k in keys:
        df = load_snapshot(k)
        if df is not None and not df.empty:
            frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def load_latest_snapshot(day: date) -> pd.DataFrame:
    """Load only the most recent snapshot of a day (single point in time)."""
    keys = list_snapshot_keys(day)
    if keys:
        df = load_snapshot(keys[-1])
        return df if df is not None else pd.DataFrame()
    # local fallback
    local_dir = _LOCAL_ROOT / day.strftime("%Y%m%d")
    if local_dir.exists():
        files = sorted(local_dir.glob("*.parquet"))
        if files:
            return pd.read_parquet(files[-1])
    return pd.DataFrame()
