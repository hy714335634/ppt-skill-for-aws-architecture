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

    SLIDE_W = 13.333
    SLIDE_H = 7.5

    for i, diag in enumerate(load_input(spec)):
        prefix = f"[diagram {i}: {diag.title or 'untitled'}]"
        title_h = (diag.title_height if diag.title_height is not None
                   else (0.85 if (diag.title and diag.subtitle)
                         else 0.55 if diag.title else 0.0))
        body_top = title_h + 0.20 if diag.title else 0.0

        for nid, node in diag.nodes.items():
            if not catalog.has(node.icon):
                hints = catalog.suggest(node.icon, limit=5)
                msg = f"{prefix} node '{nid}' uses unknown icon '{node.icon}'."
                if hints:
                    msg += f" Did you mean: {', '.join(hints)}?"
                errors.append(msg)
            if node.group and node.group not in diag.groups:
                errors.append(f"{prefix} node '{nid}' refers to missing group '{node.group}'.")
            # Manual layout overflow warnings
            if diag.layout == "manual" and node.x is not None and node.y is not None:
                if node.x < 0 or node.y < 0:
                    warnings.append(f"{prefix} node '{nid}' x/y={node.x:.2f}/{node.y:.2f} starts off-slide.")
                if node.x + node.size > SLIDE_W:
                    warnings.append(f"{prefix} node '{nid}' icon extends past right edge "
                                    f"(x={node.x:.2f} + size={node.size:.2f} > {SLIDE_W}).")
                if node.y + node.size > SLIDE_H:
                    warnings.append(f"{prefix} node '{nid}' icon extends past bottom edge.")
                if diag.title and node.y < body_top - 0.1:
                    warnings.append(f"{prefix} node '{nid}' y={node.y:.2f} is inside the title bar "
                                    f"(reserve y >= {body_top:.2f} or set title_height).")

        for g in diag.groups.values():
            if g.parent and g.parent not in diag.groups:
                errors.append(f"{prefix} group '{g.id}' refers to missing parent '{g.parent}'.")
            # Manual group overflow warnings
            if all(v is not None for v in (g.x, g.y, g.w, g.h)):
                if g.x < 0 or g.y < 0:  # type: ignore[operator]
                    warnings.append(f"{prefix} group '{g.id}' starts off-slide "
                                    f"(x={g.x:.2f} y={g.y:.2f}).")
                if g.x + g.w > SLIDE_W + 0.01:  # type: ignore[operator]
                    warnings.append(f"{prefix} group '{g.id}' extends past right edge "
                                    f"(x+w={g.x + g.w:.2f} > {SLIDE_W}).")
                if g.y + g.h > SLIDE_H + 0.01:  # type: ignore[operator]
                    warnings.append(f"{prefix} group '{g.id}' extends past bottom edge.")
                if (diag.title and g.y is not None and g.y < body_top - 0.1
                        and g.label_position == "inside-top"):
                    warnings.append(f"{prefix} group '{g.id}' y={g.y:.2f} is inside the title bar "
                                    f"(reserve y >= {body_top:.2f}, set title_height, "
                                    f"or use label_position: outside-top).")
            for sid in g.spans:
                if sid not in diag.groups:
                    errors.append(f"{prefix} group '{g.id}' spans missing group '{sid}'.")

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
