"""Icon catalog lookup for the AWS architecture skill.

The catalog.json file is built once from the official AWS asset package. This
module loads it and resolves a user-supplied name (e.g. "ec2", "amazon-ec2",
"AWS Lambda") into an absolute filesystem path to a PNG icon.
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Iterable

SKILL_ROOT = Path(__file__).resolve().parent.parent
CATALOG_PATH = SKILL_ROOT / "catalog.json"


def _slug(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")


class Catalog:
    def __init__(self, path: Path = CATALOG_PATH):
        data = json.loads(path.read_text(encoding="utf-8"))
        self.icons: dict = data["icons"]
        self.aliases: dict = data["aliases"]
        self.root: Path = path.parent

    def resolve(self, name: str) -> Path:
        """Resolve a user name to an absolute icon path. Raises KeyError if not found."""
        key = self._lookup_key(name)
        rel = self.icons[key]["path"]
        return self.root / rel

    def info(self, name: str) -> dict:
        key = self._lookup_key(name)
        entry = dict(self.icons[key])
        entry["resolved_key"] = key
        return entry

    def has(self, name: str) -> bool:
        try:
            self._lookup_key(name)
            return True
        except KeyError:
            return False

    def suggest(self, name: str, limit: int = 8) -> list[str]:
        """Best-effort suggestions for a missing name."""
        s = _slug(name)
        tokens = [t for t in s.split("-") if t]
        scored: list[tuple[int, str]] = []
        for key in self.icons:
            score = sum(1 for t in tokens if t in key)
            if score:
                scored.append((score, key))
        scored.sort(key=lambda x: (-x[0], len(x[1])))
        return [k for _, k in scored[:limit]]

    def _lookup_key(self, name: str) -> str:
        s = _slug(name)
        if s in self.icons:
            return s
        if s in self.aliases:
            return self.aliases[s]
        # try with grp- / res- / cat- prefixes
        for prefix in ("grp-", "res-", "cat-"):
            if (prefix + s) in self.icons:
                return prefix + s
        # try without prefix if user passed one
        for prefix in ("grp-", "res-", "cat-"):
            if s.startswith(prefix) and s[len(prefix):] in self.icons:
                return s[len(prefix):]
        raise KeyError(name)


def list_categories(cat: Catalog) -> dict[str, list[str]]:
    """Return {category: [keys...]} for browsing."""
    out: dict[str, list[str]] = {}
    for key, entry in cat.icons.items():
        c = entry.get("category") or entry["kind"]
        out.setdefault(c, []).append(key)
    for v in out.values():
        v.sort()
    return out
