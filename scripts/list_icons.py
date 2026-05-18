"""Browse the icon catalog. Useful when you forget the exact icon name.

Usage:
    python list_icons.py                # categories overview
    python list_icons.py compute        # icons in a category (substring match)
    python list_icons.py --search ec2   # full-text search across keys + names
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from aws_catalog import Catalog, list_categories  # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("category", nargs="?", help="Substring of category name")
    ap.add_argument("--search", "-s", help="Search across icon keys and labels")
    ap.add_argument("--limit", type=int, default=200)
    args = ap.parse_args()

    cat = Catalog()

    if args.search:
        q = args.search.lower()
        hits = []
        for key, entry in cat.icons.items():
            if q in key or q in entry["name"].lower():
                hits.append((key, entry))
        for key, entry in hits[: args.limit]:
            print(f"  {key:50s} {entry['kind']:8s} {entry.get('category', '-')}: {entry['name']}")
        # also surface aliases
        alias_hits = [a for a in cat.aliases if q in a]
        for a in alias_hits:
            print(f"  alias: {a:42s} -> {cat.aliases[a]}")
        return 0

    if args.category:
        cats = list_categories(cat)
        q = args.category.lower()
        for c, keys in sorted(cats.items()):
            if q in c.lower():
                print(f"\n{c} ({len(keys)})")
                for k in keys[: args.limit]:
                    entry = cat.icons[k]
                    print(f"  {k:50s} {entry['name']}")
        return 0

    cats = list_categories(cat)
    print(f"Catalog ({sum(len(v) for v in cats.values())} icons across {len(cats)} categories):")
    for c, keys in sorted(cats.items()):
        print(f"  {c:40s} {len(keys):4d}")
    print(f"\nAliases: {len(cat.aliases)} (e.g. ec2, lambda, s3, vpc, ...)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
