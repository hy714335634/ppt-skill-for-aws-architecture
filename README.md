# ppt-skill-for-aws-architecture

A [Claude Code](https://claude.com/claude-code) skill that generates `.pptx` AWS architecture diagrams from a YAML or JSON spec. It uses the **official AWS Architecture Icons** (2025-07-31 release), supports the official group containers (VPC, AZ, subnet, AWS Cloud, account, ASG…), and renders connectors with arrowheads, orthogonal routing, waypoints, and italic annotation labels — the same visual vocabulary AWS uses in its reference diagrams.

The skill itself lives under `skills/ppt-skill-for-aws-architecture/` so this repo can hold one or more skills side-by-side in the future.

> **What this is:** a deck generator. Hand it a spec, get a `.pptx`. The shapes are real PowerPoint shapes — fully editable in PowerPoint or Keynote afterwards.
>
> **What this is not:** an automatic-from-CloudFormation tool. You describe the architecture; the skill draws it.

## Why this exists

Most PowerPoint AWS diagrams are hand-drawn pixel-by-pixel. That's slow, inconsistent, and the moment the architecture changes the diagram drifts. With this skill:

- Describe the architecture in YAML — Claude (or you) can iterate on the spec rapidly.
- Re-render in seconds. Add a service, change a tier, swap a region — one edit, one re-render.
- The slide is a real `.pptx` — share, comment on, and edit it like any other deck.
- Faithful enough to **replicate complex community reference diagrams** with manual layout (see `examples/`).

## Capabilities at a glance

- **822 official AWS icons**: 307 services, 477 resources, 13 group containers, 25 category overviews — copied straight from the AWS Asset Package, no redraws.
- **238 short aliases**: `ec2`, `lambda`, `s3`, `vpc`, `cloudfront`, `dynamodb`, `workspaces`, `cloud9`, `cloudtrail`, `efs`, `glacier`, `opensearch`, `internet-gateway`, `nat-gateway`, `vpc-endpoints`, `transit-gateway`, `client`, `users`, `office`, `internet`, `router`, … plus full slugs (`amazon-ec2`, `aws-lambda`).
- **Group containers** with the official AWS border styles: `vpc`, `aws-cloud`, `region`, `availability-zone`, `public-subnet`, `private-subnet`, `account`, `auto-scaling-group`, `corporate-data-center`, `ec2-instance-contents`, `server-contents`, `spot-fleet`.
- **Custom containers** — paint a "Next Generation Firewall" or "Compute ENV" box with any colour, dash, fill, label style. Pre-styled palettes shipped: `next-gen-firewall`, `operating-env`, `compute-env`, `data-source`, `business-intelligence`, `ai-bigdata`, `report-generation`, `workflow`, `workspaces`, `office`, `highlight`.
- **ASG-style spanning overlays** — an ASG that wraps multiple subnets without containing all their nodes (`spans: [subnet_a, subnet_b]`).
- **Connectors** with `solid` / `dash` / `dot` / `dashdot` / `long-dash` styles, `->` / `<->` / `-` directions, custom colour & width.
- **Routing**: `straight` (default), `orthogonal` (auto L/Z bend), or explicit `waypoints` for edges that wrap around unrelated icons.
- **Edge labels** with full control: position along the polyline, offset, italic / bold / colour / font size / background (incl. transparent).
- **Inline tables** (route tables, IP plans) with header row, custom colours, and free positioning.
- **Themes**: `dark` (AWS Squid Ink navy, default) and `light`. Both use AWS Smile Orange (`#FF9900`) for accents.
- **Manual or auto layout** — auto for quick first drafts, manual for pixel-faithful replicas.
- **Multi-slide decks** via `diagrams: [...]`.

## Install / setup

```bash
# 1. Clone the repo, then symlink the skill directory into your Claude Code
#    skills folder so the loader picks it up automatically.
git clone https://github.com/hy714335634/ppt-skill-for-aws-architecture.git
cd ppt-skill-for-aws-architecture
ln -s "$(pwd)/skills/ppt-skill-for-aws-architecture" \
      ~/.claude/skills/ppt-skill-for-aws-architecture

# 2. Python deps:
pip install python-pptx pyyaml lxml

# 3. (Optional, for visual verification) LibreOffice + poppler:
brew install --cask libreoffice
brew install poppler
```

The skill is self-contained: under `skills/ppt-skill-for-aws-architecture/`, `SKILL.md` is the entry point, `catalog.json` indexes the icons, `assets/` ships the PNGs, and `scripts/` holds the renderer.

## Quick start

All commands run from the repo root. The skill scripts use paths relative to themselves so this works regardless of CWD.

```bash
SKILL=skills/ppt-skill-for-aws-architecture

# Validate (cheap, catches typos before render)
python "$SKILL/scripts/validate.py" "$SKILL/examples/three-tier-web.yaml"

# Render a YAML spec to .pptx
python "$SKILL/scripts/render.py" "$SKILL/examples/three-tier-web.yaml" -o /tmp/three-tier.pptx

# Browse the icon catalog
python "$SKILL/scripts/list_icons.py"                      # categories overview
python "$SKILL/scripts/list_icons.py" compute              # icons in a category
python "$SKILL/scripts/list_icons.py" --search dynamodb    # find by name

# Visually verify (recommended)
soffice --headless --convert-to pdf /tmp/three-tier.pptx
pdftoppm -jpeg -r 110 /tmp/three-tier.pdf /tmp/slide
open /tmp/slide-1.jpg
```

## Spec at a glance

```yaml
title: "3-Tier Web Application"
subtitle: "Multi-AZ, with managed RDS and CloudFront"
theme: dark | light                # default: dark
layout: auto | manual              # default: auto

groups:
  - id: cloud
    kind: aws-cloud
    label: "AWS Cloud"
  - id: vpc
    kind: vpc
    label: "VPC 10.0.0.0/16"
    parent: cloud

nodes:
  - id: cdn
    icon: cloudfront
    label: "CloudFront"
  - id: alb
    icon: alb
    label: "Application LB"
    group: vpc

edges:
  - { from: cdn, to: alb, label: "https" }

tables:
  - id: rt_public
    columns: ["Destination", "Gateway"]
    rows:
      - ["10.0.0.0/16", "local"]
      - ["0.0.0.0/0",   "Internet Gateway"]
    x: 5.2
    y: 0.6
    w: 2.85
    h: 0.85
```

See **[SKILL.md](SKILL.md)** for the complete schema (every field on every type, with defaults).

## Examples

Six end-to-end specs in `examples/`:

| Spec | Demonstrates |
|---|---|
| `three-tier-web.yaml`    | **Auto-layout**, VPC + 2 AZs + tiered subnets, multi-AZ RDS, CloudFront → ALB. |
| `serverless-api.yaml`    | Flat serverless: API Gateway → Lambda → DynamoDB + async fan-out. |
| `vdi-workspaces.yaml`    | **Manual layout**, custom containers ("Next Generation Firewall", "Operating ENV", "Workspaces", "Office"), `orthogonal` + `waypoints` routing, italic orange labels, side stack of management services. |
| `multi-az-ha.yaml`       | 2 AZs × 4 private subnets + 1 public subnet, **ASG `spans` overlays**, URL-path routing labels, `<->` HA replication arrows. |
| `two-az-ha-simple.yaml`  | Compact 2-AZ web app + bastion VPN + **inline route tables**. |
| `genomics-pipeline.yaml` | Pink workflow box, AWS Batch fan-in, **Compute ENV**, four side dashed containers (Data Source / BI / AI-Big Data / Report Generation), bidirectional read/write with S3 + Glacier lifecycle. |

Each is a few seconds to render:

```bash
SKILL=skills/ppt-skill-for-aws-architecture
for f in "$SKILL"/examples/*.yaml; do
  python "$SKILL/scripts/render.py" "$f" -o "/tmp/$(basename "${f%.yaml}").pptx"
done
```

## Replicating reference diagrams

The skill is designed to faithfully replicate community AWS reference diagrams. The trick:

1. **Use `layout: manual`** and pin every group with `x/y/w/h`. For replication, pixel-similar coordinates beat fighting auto-layout.
2. **Approximate the canvas as a 13.33" × 7.5" grid.** Read coordinates off the reference; if it's 2000px wide, divide by 150 to land in slide units.
3. For each rectangular container, decide: official AWS group (use the matching `kind`) or descriptive box (use a pre-styled palette or `kind: custom` with `border_color`).
4. **Italic orange labels** on edges (`label_color: "FF9900"`, `label_italic: true`, `label_bg: transparent`) match the AWS reference style.
5. **`routing: orthogonal`** for any 90° edge; **`routing: waypoints`** when wrapping around unrelated icons.

The four reference-replica specs in `examples/` are the canonical templates. Pattern-match: which one is closest to your target, copy it, and adapt.

## Updating the icon catalog

Official AWS Architecture Icons are republished quarterly:

```bash
# 1. Download the latest Asset Package from https://aws.amazon.com/architecture/icons/
# 2. Edit SRC in skills/ppt-skill-for-aws-architecture/scripts/build_catalog.py
#    to point at the unzipped directory
# 3. Rebuild
python skills/ppt-skill-for-aws-architecture/scripts/build_catalog.py
```

This rewrites `skills/ppt-skill-for-aws-architecture/assets/` and `catalog.json` in place. Bump the `version` in `catalog.json` to match the AWS release date.

## Repository layout

```
ppt-skill-for-aws-architecture/
├── README.md                                     # this file
└── skills/
    └── ppt-skill-for-aws-architecture/
        ├── SKILL.md              # Claude Code skill entry: full schema + workflow
        ├── catalog.json          # icon name -> path index (822 icons, 238 aliases)
        ├── assets/
        │   ├── services/         # 307 PNG @64px (Compute, Database, Networking, ...)
        │   ├── resources/        # 477 PNG @48px (sub-resources: EC2 Instance, Lambda Function, ...)
        │   ├── groups/           # 13 PNG @32px (VPC, Region, Subnet, Account, ...)
        │   └── categories/       # 25 PNG @64px (category overview icons)
        ├── scripts/
        │   ├── render.py         # main entry: spec -> .pptx
        │   ├── validate.py       # quick spec linter (no rendering)
        │   ├── list_icons.py     # browse the catalog
        │   ├── aws_catalog.py    # icon name resolver (importable module)
        │   └── build_catalog.py  # one-shot: rebuild catalog.json from an AWS Asset Package
        └── examples/
            ├── three-tier-web.yaml
            ├── serverless-api.yaml
            ├── vdi-workspaces.yaml
            ├── multi-az-ha.yaml
            ├── two-az-ha-simple.yaml
            └── genomics-pipeline.yaml
```

## Licensing

- **AWS Architecture Icons** (under `assets/`) are © Amazon Web Services, Inc. and licensed under the [AWS Architecture Icons License](https://aws.amazon.com/architecture/icons/).
- **Skill code** (everything else) is unrestricted internal use.

## Dependencies

- Python ≥ 3.9
- `python-pptx ≥ 1.0`
- `pyyaml`
- `lxml` (transitive via python-pptx)

For visual verification: `soffice` (LibreOffice) + `pdftoppm` (poppler).
