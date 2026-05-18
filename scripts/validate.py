"""Validate an architecture spec without rendering. Reports unknown icons,
dangling edges, missing groups.

Usage: python validate.py <input.yaml|json>
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from aws_catalog import Catalog  # noqa: E402
from render import load_input  # noqa: E402


def main() -> int:
    if len(sys.argv) != 2:
        print("Usage: python validate.py <input.yaml|json>", file=sys.stderr)
        return 2

    spec = Path(sys.argv[1])
    if not spec.exists():
        print(f"Not found: {spec}", file=sys.stderr)
        return 1

    catalog = Catalog()
    errors: list[str] = []
    warnings: list[str] = []

    for i, diag in enumerate(load_input(spec)):
        prefix = f"[diagram {i}: {diag.title or 'untitled'}]"

        for nid, node in diag.nodes.items():
            if not catalog.has(node.icon):
                hints = catalog.suggest(node.icon, limit=5)
                msg = f"{prefix} node '{nid}' uses unknown icon '{node.icon}'."
                if hints:
                    msg += f" Did you mean: {', '.join(hints)}?"
                errors.append(msg)
            if node.group and node.group not in diag.groups:
                errors.append(f"{prefix} node '{nid}' refers to missing group '{node.group}'.")

        for g in diag.groups.values():
            if g.parent and g.parent not in diag.groups:
                errors.append(f"{prefix} group '{g.id}' refers to missing parent '{g.parent}'.")

        for e in diag.edges:
            for endpoint in (e.source, e.target):
                if endpoint not in diag.nodes:
                    errors.append(f"{prefix} edge {e.source}->{e.target} references missing node '{endpoint}'.")

        if not diag.nodes:
            warnings.append(f"{prefix} no nodes defined.")

    for w in warnings:
        print(f"warning: {w}")
    for e in errors:
        print(f"error: {e}", file=sys.stderr)

    if errors:
        return 1
    print(f"OK ({spec})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
