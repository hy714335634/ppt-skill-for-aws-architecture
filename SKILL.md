---
name: aws-architecture
description: "Use this skill any time the user wants to draw, generate, or update an AWS architecture diagram in PowerPoint (.pptx). Triggers: 'AWS architecture diagram', 'draw an AWS architecture', 'AWS architecture deck/slides/PPT', 'design an AWS solution and diagram it', 'put this AWS architecture in slides', or any request that involves visualizing an AWS solution with services like EC2 / Lambda / S3 / RDS / VPC and AZs / subnets / regions / accounts. The skill takes a natural-language description (or YAML/JSON spec) of the architecture and produces a slide deck with official AWS Architecture icons, group containers (VPC, AZ, subnet, account, region, AWS Cloud), and connectors."
license: AWS Architecture Icons assets are © Amazon Web Services, Inc. and licensed under the AWS Architecture Icons License (https://aws.amazon.com/architecture/icons/). Skill code itself is unrestricted internal use.
---

# AWS Architecture Diagram Skill

Generates `.pptx` decks with official AWS Architecture icons (2025-07-31 release), group containers (VPC, Region, AZ, subnet, account, AWS Cloud), connectors with orthogonal routing & waypoints, custom-styled containers, and inline tables. Powered by `python-pptx`. Invoked from a YAML or JSON spec.

## When to use this skill

Use this skill — not the generic `pptx` skill — whenever the user wants an **AWS architecture diagram in PowerPoint**. Hand off to the `pptx` skill for general slide editing/reading after the diagram is generated.

## Quick start

```bash
# 1. Author or generate a spec
python scripts/render.py examples/serverless-api.yaml -o /tmp/aws.pptx

# 2. Browse the icon catalog when you need a service name
python scripts/list_icons.py --search dynamodb
python scripts/list_icons.py compute       # browse a category

# 3. Validate before rendering (fast feedback on typos)
python scripts/validate.py my-spec.yaml
```

The skill ships **822 PNG icons** (307 services, 477 resources, 13 group containers, 25 categories) plus **238 short aliases** covering everyday services (`ec2`, `lambda`, `s3`, `vpc`, `cloudfront`, `dynamodb`), edge resources (`internet-gateway`, `nat-gateway`, `vpc-endpoints`, `vpn-gateway`, `transit-gateway`), end-user computing (`workspaces`, `cloud9`), security/observability (`cloudtrail`, `guardduty`, `waf`, `shield`), data (`opensearch`, `athena`, `emr`, `glacier`, `efs`, `elasticache-redis`), and more.

## Spec schema

A spec is a YAML or JSON document. Top level (single diagram):

```yaml
title: "..."             # large heading
subtitle: "..."          # small subheading
description: "..."       # footer caption
theme: dark | light      # default: dark (AWS Squid Ink navy)
layout: auto | manual    # default: auto
groups: [...]            # group containers (VPC, AZ, subnet, etc.)
nodes:  [...]            # services/resources/users
edges:  [...]            # connectors between nodes
```

Or multiple diagrams in one deck — wrap with `diagrams: [ {...}, {...} ]`.

### Groups (containers)

```yaml
groups:
  - id: vpc                    # required, referenced by nodes/groups
    kind: vpc                  # vpc | aws-cloud | aws-cloud-alt | region
                               #   az | availability-zone
                               #   public-subnet | private-subnet | account
                               #   auto-scaling-group | corporate-data-center
                               #   ec2-instance-contents | server-contents | spot-fleet
                               #   custom | default
                               # Pre-styled palettes (all caller-overridable):
                               #   next-gen-firewall | operating-env | compute-env
                               #   data-source | business-intelligence | ai-bigdata
                               #   report-generation | workflow | workspaces
                               #   office | highlight
    label: "VPC 10.0.0.0/16"   # shown in border
    parent: cloud              # optional: nest inside another group
    direction: row | column | auto   # how children are arranged inside; default auto
    pad: 0.45                  # inches of inner padding (default 0.45)

    # --- Optional manual placement ---
    # Set ALL FOUR to lock the group's bounding box. The group skips auto-layout;
    # its children still auto-flow inside (or you can pin them too).
    x: 0.6
    y: 1.2
    w: 5.5
    h: 5.0

    # --- Optional custom styling ---
    border_color: "FF9900"      # hex (no '#'); overrides the kind default
    border_style: dash          # solid | dash | dot | dashdot | long-dash
    fill: "232F3E"              # hex; optional inner fill
    fill_alpha: 0.08            # 0..1 opacity if fill is set
    label_color: "C5CCD8"       # label text color
    label_size: 11              # point size
    label_bold: true            # default true
    label_italic: false
    show_icon: true             # AWS group icon in the corner; default true

    # --- Optional ASG-style overlay ---
    # If `spans` is set, the group is positioned to cover the listed groups
    # (after they've been laid out) — used for Auto Scaling Group bands that
    # straddle multiple subnets without containing all of their nodes.
    spans: [pub_a, pub_b]
```

Each `kind` ships an official border style (color + dash) and a small icon in the top-left corner. The `custom`/`default` kinds show no preset border colour — provide `border_color` yourself.

**Manual group placement** is what you reach for when auto-layout produces an unbalanced canvas — typically when one region/VPC has many children and another has few. Pin the larger groups with `x/y/w/h`, leave smaller ones to auto-flow, and connectors anchor cleanly between the two.

**Custom-named containers** (e.g. "Next Generation Firewall", "Data Source", "Business Intelligence", "Compute ENV") are achievable in two ways:
1. Pick the matching pre-styled `kind` and override `label`.
2. Use `kind: custom` and set `border_color`, `border_style`, `label`, `label_color` yourself.

**Spanning overlays** (image 2's "Auto Scaling Group" wrapping multiple subnets) are achievable with `spans: [subnet_id_a, subnet_id_b]` — after the underlying subnets get laid out, the spanning group is sized to cover them with a small padding.

### Nodes (services)

```yaml
nodes:
  - id: handler                # required
    icon: lambda               # alias OR full key from catalog.json
    label: "API Handler"       # shown under the icon
    sublabel: "concurrency 10" # optional second line
    group: api_tier            # optional: place inside a group
    size: 0.85                 # icon side in inches (default 0.85)
    x: 4.2                     # optional manual position (inches from slide left)
    y: 3.0                     # optional manual position (inches from slide top)
```

`icon` accepts:
- A short alias: `ec2`, `lambda`, `s3`, `rds`, `dynamodb`, `sqs`, `sns`, `apigateway`, `cloudfront`, `route53`, `iam`, `cognito`, `bedrock`, `sagemaker`, `cloudwatch`, `user`, `users`, `client` ...
- A full slug: `amazon-ec2`, `aws-lambda`, `amazon-rds`, `amazon-bedrock`, ...
- A resource key prefixed with `res-`: `res-user`, `res-mobile-client`, `res-amazon-ec2-instance`.

If the icon name is unknown, the renderer prints suggestions and exits non-zero. Use `python scripts/list_icons.py --search <term>` to discover names.

### Edges (connectors)

```yaml
edges:
  - from: handler
    to:   table
    label: "read/write"        # optional inline label
    style: solid | dash | dot | dashdot | long-dash   # default solid
    direction: "->" | "<->" | "-"   # default "->"
    color: "FF9900"            # line color, optional hex (no '#')
    width: 1.5                 # line thickness in points (default 1.5)

    # --- Anchoring ---
    from_anchor: auto | top | bottom | left | right   # default auto
    to_anchor:   auto | top | bottom | left | right   # default auto

    # --- Routing ---
    routing: straight | orthogonal | waypoints   # default straight
    waypoints:                                   # used when routing=waypoints
      - [4.5, 2.3]
      - [4.5, 4.7]
      - [7.2, 4.7]

    # --- Label appearance ---
    label_color: "FF9900"      # default theme foreground
    label_italic: true         # default false (AWS reference docs use italic)
    label_bold: false
    label_size: 9              # default 9pt
    label_position: 0.5        # 0..1 along the polyline
    label_offset: [0, -0.05]   # nudge from natural position (inches)
    label_bg: transparent      # transparent | none | hex; default theme bg
    label_pad: 0.05            # text-frame padding
```

**Routing modes:**
- `straight` (default) — direct line between anchored faces.
- `orthogonal` — automatic L/Z routing with one or two 90° bends. Pair with `from_anchor` to lock the first segment to a face.
- `waypoints` — explicit bend points; the segment from source to the first waypoint and from the last waypoint to target are anchored to the closest face automatically (or to `from_anchor` / `to_anchor` if specified).

Anchors auto-pick the side facing the other node, so simple flows render cleanly without coordinates. Set `from_anchor` / `to_anchor` to force a specific face when auto-routing crosses an unrelated icon or VPC border.

### Tables (inline data tables)

Useful for embedding route tables, IP plans, parameter lists alongside an architecture.

```yaml
tables:
  - id: rt_public
    title: "Public Route Table"      # optional merged header row
    columns: ["Destination", "Gateway"]
    rows:
      - ["10.0.0.0/16", "local"]
      - ["0.0.0.0/0",    "Internet Gateway"]
    x: 5.2                          # inches from slide left
    y: 0.55                         # inches from slide top
    w: 2.85                         # width in inches
    h: 0.85                         # height in inches
    font_size: 9
    header_bg: "424F60"             # hex; default dark slate
    header_fg: "FFFFFF"
    body_bg:   "2A3645"
    body_fg:   "FFFFFF"
    border_color: "EB6F6F"          # red highlight, mimics image 3
```

### Themes

- `dark` (default) — AWS Squid Ink (`#232F3E`) background, white text, light-gray edges. Use for executive or pitch decks.
- `light` — white background, navy text, dark-gray edges. Use for documentation/whiteboard style.

Both themes use the AWS Smile Orange (`#FF9900`) for accents and the official border colors per group kind.

## Replicating complex reference diagrams

When a user hands you a screenshot of an AWS reference architecture and asks for a faithful PPT, treat it as a layout-replication task. The four `examples/` files (vdi-workspaces, multi-az-ha, two-az-ha-simple, genomics-pipeline) are the canonical templates — each one demonstrates a different combination of features. Pattern-match: which one is closest to the user's diagram, then copy and adapt.

The trick that unlocks most reference diagrams:

1. **Use `layout: manual` and pin every group** with `x/y/w/h`. Auto-layout is great for new designs; for replication, pixel-similar coordinates are faster than fighting auto-layout.
2. **Approximate the canvas as a 13.33" × 7.5" grid.** Read coordinates off the reference image and convert to inches (e.g., reference image is 2000px wide → divide by 150 to land in slide units).
3. **For each rectangular container in the reference, decide:** is it an official AWS group (`vpc`, `private-subnet`, `auto-scaling-group`) — then use that `kind`. Or is it a labelled rectangle ("Next Generation Firewall", "Compute ENV") — then use one of the pre-styled palettes or `kind: custom` with a `border_color`.
4. **Use orange italic labels** (`label_color: "FF9900"`, `label_italic: true`, `label_bg: transparent`) for the descriptive edge annotations characteristic of AWS reference diagrams.
5. **Use `routing: orthogonal`** for any edge that has a 90° bend in the reference. Pair with `from_anchor` to dictate the first leg direction.
6. **Use `routing: waypoints`** when an edge wraps around an unrelated icon or container — common in the "VPC Flowlog" arrows in image 1 or NAT routes in image 3.
7. **Tables (route tables, IP plans)** — use the top-level `tables:` section, not nodes.

## Authoring workflow

When the user asks for an AWS diagram:

1. **Clarify, don't hallucinate.** If the user gives vague requirements ("design a video processing pipeline"), ask **one round** of questions: scale, regions, sync vs. async, data store choices. Then produce a draft spec.
2. **Pick a layout strategy:**
   - **Few nodes (≤ 8) without VPC:** drop them in `groups` of `aws-cloud` only or no groups at all and let auto-layout flow.
   - **VPC architectures:** model `cloud > vpc > az > subnet`. Set `direction: column` on AZs to stack public/app/db subnets.
   - **Outside-the-cloud actors** (users, on-prem, mobile clients): leave them ungrouped — they sit beside the cloud and the renderer auto-positions them.
3. **Use `direction` aggressively.** Auto-mode picks a square grid; horizontal pipelines (`row`) or vertical tiers (`column`) almost always look better when stated explicitly.
4. **Validate first:** `python scripts/validate.py spec.yaml`. Cheap, catches typos before the slow PDF render in QA.
5. **Render:** `python scripts/render.py spec.yaml -o out.pptx`.
6. **Visually verify (REQUIRED).** Convert and inspect — diagram correctness is hard to confirm from code:
    ```bash
    soffice --headless --convert-to pdf out.pptx
    pdftoppm -jpeg -r 110 out.pdf slide
    ```
    Then `Read` each `slide-NN.jpg` with fresh eyes (or delegate to a subagent). Look specifically for:
    - Edge crossings through unrelated icons (consider re-grouping)
    - Group icons overlapping member nodes (group probably too small — shorten labels or `pad: 0.3`)
    - Edge labels colliding (avoid labeling every edge — keep labels for the 2–3 most important flows)
    - Off-slide overflow on the right/bottom (auto-layout failed; switch to `layout: manual` and set `x`/`y` per node)

Iterate until a clean pass.

## Conventions to follow

- **Stick to the official AWS group containers.** Don't invent a "Microservices" or "Data layer" container with a custom border — use `kind: default` if you must.
- **Region/AZ before VPC.** AWS conventional nesting is `Region > VPC > AZ > Subnet`. The skill doesn't enforce this, but matching it improves recognition.
- **One AWS Cloud per slide.** If you have multiple regions, nest them inside the same `cloud` group rather than drawing separate clouds.
- **Pair primary icon with sublabel for instances.** `icon: amazon-rds` + `sublabel: "Primary"` is clearer than two unlabeled RDS icons.
- **Don't ladder more than 4 levels deep on auto-layout.** Region/AZ/subnet plus a node is 4. Deeper nesting compresses icons unreadable; switch to `layout: manual`.

## Files

```
aws-architecture/
├── SKILL.md              # this file
├── catalog.json          # icon name -> path index (built once from AWS Asset Package)
├── assets/
│   ├── services/         # 307 PNG @64px (Compute, Database, Networking, ...)
│   ├── resources/        # 477 PNG @48px (sub-resources like EC2 Instance, Lambda Function)
│   ├── groups/           # 13 PNG @32px (VPC, Region, Subnet, Account, ...)
│   └── categories/       # 25 PNG @64px (category overview icons)
├── scripts/
│   ├── render.py         # main entry point: spec -> .pptx
│   ├── validate.py       # quick spec linter (no rendering)
│   ├── list_icons.py     # browse the catalog
│   ├── aws_catalog.py    # icon name resolver (importable module)
│   └── build_catalog.py  # one-shot: rebuild catalog.json from an AWS Asset Package
└── examples/
    ├── three-tier-web.yaml
    └── serverless-api.yaml
```

## Examples

Run them as smoke tests:

```bash
python scripts/render.py examples/three-tier-web.yaml      -o /tmp/three-tier.pptx
python scripts/render.py examples/serverless-api.yaml      -o /tmp/serverless.pptx
python scripts/render.py examples/vdi-workspaces.yaml      -o /tmp/vdi.pptx
python scripts/render.py examples/multi-az-ha.yaml         -o /tmp/multi-az.pptx
python scripts/render.py examples/two-az-ha-simple.yaml    -o /tmp/two-az.pptx
python scripts/render.py examples/genomics-pipeline.yaml   -o /tmp/genomics.pptx
```

| File | Demonstrates |
|---|---|
| `three-tier-web.yaml`     | Auto-layout, VPC + 2 AZs + tiered subnets, multi-AZ RDS, CloudFront → ALB. |
| `serverless-api.yaml`     | Flat serverless pattern: API Gateway → Lambda → DynamoDB plus async fan-out. |
| `vdi-workspaces.yaml`     | Manual layout, custom containers ("Next Generation Firewall", "Operating ENV", "Workspaces", "Office"), routing with `orthogonal` + `waypoints`, italic orange annotation labels, side-stack of management services. |
| `multi-az-ha.yaml`        | 2 AZs × 4 private subnets + 1 public subnet, ASG `spans` overlay, URL-path routing labels, `<->` HA replication arrows, side data services. |
| `two-az-ha-simple.yaml`   | Compact 2-AZ web app + bastion VPN + inline route tables (top-level `tables:` block). |
| `genomics-pipeline.yaml`  | Pink workflow box, AWS Batch fan-in, Compute ENV, four side dashed containers (Data Source / BI / AI-Big Data / Report Generation), bidirectional read/write with S3+Glacier lifecycle. |

## Common pitfalls

- **`Unknown icon 'foo'`** — alias not registered. Run `python scripts/list_icons.py --search foo` to find the canonical key, or add a short alias in `/tmp/build_aws_catalog.py` and rebuild `catalog.json`.
- **Subnet group label looks clipped** — the group is too narrow for the label. Shorten (`Public` instead of `Public subnet`) or widen the parent group (fewer siblings, or set `direction: column`).
- **Edges crossing the diagram** — re-group nodes so closely-connected ones share a parent, or split the architecture across multiple slides (`diagrams: [...]`).
- **Edge labels stack on top of each other** — only label edges that carry information the diagram doesn't already show. Skip "calls", "request", obvious arrows.

## Dependencies

Already on the system if you have the project venv activated (`source .venv/bin/activate`):

- `python-pptx >= 1.0`
- `pyyaml`
- `lxml` (transitive via python-pptx)

For visual QA: `soffice` (LibreOffice) and `pdftoppm` (poppler) — both present in the standard sandboxed environment.

## Updating the icon catalog

Official AWS Architecture Icons are republished quarterly. To refresh:

1. Download the latest **Asset Package** from <https://aws.amazon.com/architecture/icons/>.
2. Update `SRC` in `scripts/build_catalog.py` (the path to the unzipped Asset Package).
3. Run it: `python scripts/build_catalog.py`. This rewrites `assets/` and `catalog.json` in-place.
4. Bump the `version` field at the top of `catalog.json` to match the AWS release date.
