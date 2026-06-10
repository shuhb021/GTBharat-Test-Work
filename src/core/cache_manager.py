import hashlib
import pickle
import os
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

class CacheManager:
    """Intelligent Caching Layer using SHA256 hashes (Phase 6)."""
    
    def __init__(self, cache_dir=".far_cache"):
        self.cache_dir = Path(cache_dir)
        self.cache_dir.mkdir(exist_ok=True)
    
    def generate_hash(self, file_paths, params=None):
        """Generates a SHA256 hash from file contents and parameters."""
        hasher = hashlib.sha256()
        for fp in sorted(file_paths):
            if os.path.exists(fp):
                # Hash file stats for speed + first 1MB of content to guarantee uniqueness
                stat = os.stat(fp)
                hasher.update(f"{fp}_{stat.st_size}_{stat.st_mtime}".encode('utf-8'))
                try:
                    with open(fp, 'rb') as f:
                        hasher.update(f.read(1024 * 1024))
                except Exception:
                    pass
                    
        if params:
            # Sort params for consistency
            if isinstance(params, dict):
                hasher.update(str(sorted(params.items())).encode('utf-8'))
            else:
                hasher.update(str(params).encode('utf-8'))
            
        return hasher.hexdigest()
    
    def get(self, cache_key):
        cache_file = self.cache_dir / f"{cache_key}.pkl"
        if cache_file.exists():
            try:
                with open(cache_file, "rb") as f:
                    logger.info("Cache hit for %s — Returning instantly.", cache_key)
                    return pickle.load(f)
            except Exception as e:
                logger.warning("Failed to read cache: %s", e)
        return None

    def set(self, cache_key, data):
        cache_file = self.cache_dir / f"{cache_key}.pkl"
        try:
            with open(cache_file, "wb") as f:
                # Use highest protocol for memory efficiency
                pickle.dump(data, f, protocol=pickle.HIGHEST_PROTOCOL)
                logger.info("Data cached to %s", cache_file)
        except Exception as e:
            logger.warning("Failed to write cache: %s", e)
