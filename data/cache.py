"""File caching layer to avoid redundant downloads.

Cache hierarchy:
    L1: Local filesystem (cache/ directory)
    L2: Cloudflare R2 (persistent across deployments)
    L3: JPX website (source of truth)

Read:  L1 -> L2 -> L3
Write: L1 + L2 simultaneously
"""

import json
import hashlib
import time
from pathlib import Path
from typing import Optional
import config
from data.r2_storage import r2_get, r2_put


def ensure_cache_dirs() -> None:
    """Create cache directory structure if it doesn't exist."""
    for d in [config.CACHE_VOLUME_DIR, config.CACHE_OI_DIR, config.CACHE_INDEX_DIR,
              config.CACHE_DAILY_OI_DIR]:
        d.mkdir(parents=True, exist_ok=True)


def _cache_path_for_url(url: str, subdir: Path) -> Path:
    """Generate a deterministic cache file path from a URL."""
    filename = url.split("/")[-1]
    if not filename:
        filename = hashlib.md5(url.encode()).hexdigest()
    return subdir / filename


def _r2_key(subdir: Path, filename: str) -> str:
    """Build an R2 object key from subdir and filename.

    e.g. 'volume/20260210_volume_by_participant_whole_day.xlsx'
    """
    return f"{subdir.name}/{filename}"


def _is_fresh(path: Path, max_age_hours: float) -> bool:
    """Check if a cached file is still fresh."""
    if not path.exists():
        return False
    age_seconds = time.time() - path.stat().st_mtime
    return age_seconds < max_age_hours * 3600


def get_cached_bytes(url: str, subdir: Path, max_age_hours: float = 24.0) -> Optional[bytes]:
    """Return cached file bytes if fresh enough, else None.

    Checks L1 (local) then L2 (R2).
    """
    path = _cache_path_for_url(url, subdir)

    # L1: local filesystem
    if _is_fresh(path, max_age_hours):
        return path.read_bytes()

    # L2: R2
    key = _r2_key(subdir, path.name)
    content = r2_get(key)
    if content is not None:
        # Populate L1 from L2
        path.write_bytes(content)
        return content

    return None


def save_to_cache(url: str, subdir: Path, content: bytes) -> Path:
    """Save downloaded content to L1 + L2."""
    path = _cache_path_for_url(url, subdir)

    # L1: local
    path.write_bytes(content)

    # L2: R2 (async-safe, fails silently)
    key = _r2_key(subdir, path.name)
    r2_put(key, content)

    return path


def get_cached_json(url: str, max_age_hours: float = 1.0) -> Optional[dict]:
    """Return cached JSON index if fresh enough, else None.

    Checks L1 (local) then L2 (R2).
    """
    path = _cache_path_for_url(url, config.CACHE_INDEX_DIR)

    # L1: local filesystem
    if _is_fresh(path, max_age_hours):
        return json.loads(path.read_text(encoding="utf-8"))

    # L2: R2
    key = _r2_key(config.CACHE_INDEX_DIR, path.name)
    content = r2_get(key)
    if content is not None:
        # Populate L1 from L2
        path.write_bytes(content)
        return json.loads(content.decode("utf-8"))

    return None


def save_json_to_cache(url: str, data: dict) -> None:
    """Save JSON index data to L1 + L2."""
    path = _cache_path_for_url(url, config.CACHE_INDEX_DIR)
    text = json.dumps(data, ensure_ascii=False)

    # L1: local
    path.write_text(text, encoding="utf-8")

    # L2: R2
    key = _r2_key(config.CACHE_INDEX_DIR, path.name)
    r2_put(key, text.encode("utf-8"))
