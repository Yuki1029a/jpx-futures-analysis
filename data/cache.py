"""File caching layer to avoid redundant downloads."""

import json
import hashlib
import time
from pathlib import Path
from typing import Optional
import config


def ensure_cache_dirs() -> None:
    """Create cache directory structure if it doesn't exist."""
    for d in [config.CACHE_VOLUME_DIR, config.CACHE_OI_DIR, config.CACHE_INDEX_DIR]:
        d.mkdir(parents=True, exist_ok=True)


def _cache_path_for_url(url: str, subdir: Path) -> Path:
    """Generate a deterministic cache file path from a URL."""
    filename = url.split("/")[-1]
    if not filename:
        filename = hashlib.md5(url.encode()).hexdigest()
    return subdir / filename


def _is_fresh(path: Path, max_age_hours: float) -> bool:
    """Check if a cached file is still fresh."""
    if not path.exists():
        return False
    age_seconds = time.time() - path.stat().st_mtime
    return age_seconds < max_age_hours * 3600


def get_cached_bytes(url: str, subdir: Path, max_age_hours: float = 24.0) -> Optional[bytes]:
    """Return cached file bytes if fresh enough, else None."""
    path = _cache_path_for_url(url, subdir)
    if _is_fresh(path, max_age_hours):
        return path.read_bytes()
    return None


def save_to_cache(url: str, subdir: Path, content: bytes) -> Path:
    """Save downloaded content to cache, return the file path."""
    path = _cache_path_for_url(url, subdir)
    path.write_bytes(content)
    return path


def get_cached_json(url: str, max_age_hours: float = 1.0) -> Optional[dict]:
    """Return cached JSON index if fresh enough, else None."""
    path = _cache_path_for_url(url, config.CACHE_INDEX_DIR)
    if _is_fresh(path, max_age_hours):
        return json.loads(path.read_text(encoding="utf-8"))
    return None


def save_json_to_cache(url: str, data: dict) -> None:
    """Save JSON index data to cache."""
    path = _cache_path_for_url(url, config.CACHE_INDEX_DIR)
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
