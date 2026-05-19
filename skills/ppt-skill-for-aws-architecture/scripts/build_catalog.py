"""One-shot: copy curated AWS icon PNGs into the skill, emit catalog.json.

Source layout (Asset-Package):
  Architecture-Service-Icons_07312025/Arch_<Category>/{16,32,48,64}/Arch_<Name>_64.png
  Resource-Icons_07312025/Res_<Category>/{32,48}/Res_<Name>_48.png
  Architecture-Group-Icons_07312025/<Name>_32.png
  Category-Icons_07312025/Arch-Category_{16,32,48,64}/Arch-Category_<Name>_64.png

Skipped: SVG, @5x.png, retina variants.
"""
import json
import re
import shutil
from pathlib import Path

SRC = Path("/Users/qangz/Downloads/Asset-Package_07312025.49d3aab7f9e6131e51ade8f7c6c8b961ee7d3bb1")
DST = Path(__file__).resolve().parent.parent / "assets"

CATEGORY_RE = re.compile(r"^(?:Arch|Res)_(.+)$")


def norm_name(s: str) -> str:
    s = s.replace("Arch_", "").replace("Res_", "").replace("Arch-Category_", "")
    s = s.removesuffix("_64").removesuffix("_48").removesuffix("_32")
    return s


def slug(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")


def copy_service_icons() -> dict:
    catalog = {}
    src_root = SRC / "Architecture-Service-Icons_07312025"
    for cat_dir in sorted(src_root.iterdir()):
        if not cat_dir.is_dir():
            continue
        category = CATEGORY_RE.match(cat_dir.name).group(1)
        cat_slug = slug(category)
        png_dir = cat_dir / "64"
        if not png_dir.exists():
            continue
        for png in sorted(png_dir.glob("*_64.png")):
            if "@5x" in png.name:
                continue
            name = norm_name(png.stem)
            sslug = slug(name)
            dst_dir = DST / "services" / cat_slug
            dst_dir.mkdir(parents=True, exist_ok=True)
            dst = dst_dir / f"{sslug}.png"
            shutil.copy2(png, dst)
            catalog[sslug] = {
                "kind": "service",
                "category": category,
                "name": name.replace("_", " ").replace("-", " "),
                "path": str(dst.relative_to(DST.parent)),
            }
    return catalog


def copy_resource_icons() -> dict:
    catalog = {}
    src_root = SRC / "Resource-Icons_07312025"
    for cat_dir in sorted(src_root.iterdir()):
        if not cat_dir.is_dir():
            continue
        category = CATEGORY_RE.match(cat_dir.name).group(1)
        cat_slug = slug(category)

        # Possible structures: flat PNGs, "48/", or "Res_48_Light/" subfolders.
        candidate_dirs = [cat_dir / "48", cat_dir / "Res_48_Light", cat_dir]
        png_dir = next((d for d in candidate_dirs if d.exists() and any(d.glob("*_48*.png"))), None)
        if png_dir is None:
            continue

        for png in sorted(png_dir.glob("*_48*.png")):
            if "@5x" in png.name:
                continue
            stem = png.stem
            # strip _Light suffix introduced by Res_48_Light variant
            for suffix in ("_48_Light", "_48"):
                if stem.endswith(suffix):
                    stem = stem[: -len(suffix)]
                    break
            name = stem.replace("Res_", "")
            sslug = "res-" + slug(name)
            dst_dir = DST / "resources" / cat_slug
            dst_dir.mkdir(parents=True, exist_ok=True)
            dst = dst_dir / f"{slug(name)}.png"
            shutil.copy2(png, dst)
            catalog[sslug] = {
                "kind": "resource",
                "category": category,
                "name": name.replace("_", " ").replace("-", " "),
                "path": str(dst.relative_to(DST.parent)),
            }
    return catalog


def copy_group_icons() -> dict:
    catalog = {}
    src_root = SRC / "Architecture-Group-Icons_07312025"
    DST_DIR = DST / "groups"
    DST_DIR.mkdir(parents=True, exist_ok=True)
    for png in sorted(src_root.glob("*_32.png")):
        name = norm_name(png.stem)
        sslug = "grp-" + slug(name)
        # Skip Dark variants — we'll provide light by default
        if "_Dark" in name:
            continue
        dst = DST_DIR / f"{slug(name)}.png"
        shutil.copy2(png, dst)
        catalog[sslug] = {
            "kind": "group",
            "name": name.replace("_", " ").replace("-", " "),
            "path": str(dst.relative_to(DST.parent)),
        }
    return catalog


def copy_category_icons() -> dict:
    catalog = {}
    src_root = SRC / "Category-Icons_07312025/Arch-Category_64"
    DST_DIR = DST / "categories"
    DST_DIR.mkdir(parents=True, exist_ok=True)
    for png in sorted(src_root.glob("*_64.png")):
        if "@5x" in png.name:
            continue
        name = norm_name(png.stem)
        sslug = "cat-" + slug(name)
        dst = DST_DIR / f"{slug(name)}.png"
        shutil.copy2(png, dst)
        catalog[sslug] = {
            "kind": "category",
            "name": name.replace("_", " ").replace("-", " "),
            "path": str(dst.relative_to(DST.parent)),
        }
    return catalog


def build_aliases(catalog: dict) -> dict:
    """Add common short aliases pointing into the catalog."""
    aliases = {}

    def find_first(pattern: str) -> str | None:
        rx = re.compile(pattern, re.I)
        for k in catalog:
            if rx.search(k):
                return k
        return None

    common = {
        # Generic resources / external actors
        "user": "res-user",
        "users": "res-users",
        "client": "res-client",
        "clients": "res-client",
        "mobile-client": "res-mobile-client",
        "mobile": "res-mobile-client",
        "authenticated-user": "res-authenticated-user",
        "internet": "res-internet",
        "internet-alt1": "res-internet-alt1",
        "internet-alt2": "res-internet-alt2",
        "office": "res-office-building",
        "office-building": "res-office-building",
        "server": "res-server",
        "servers": "res-servers",
        "firewall": "res-firewall",
        "shield-resource": "res-shield",
        "globe": "res-globe",
        "email": "res-email",
        "chat": "res-chat",
        "forums": "res-forums",
        "credentials": "res-credentials",
        "alert": "res-alert",
        "logs": "res-logs",
        "metrics": "res-metrics",
        "gear": "res-gear",
        "magnifying-glass": "res-magnifying-glass",
        "generic-application": "res-generic-application",
        "git-repository": "res-git-repository",
        "json-script": "res-json-script",
        "saml-token": "res-saml-token",
        "sdk": "res-sdk",
        "data-stream": "res-data-stream",
        "data-table": "res-data-table",
        "database": "res-database",
        "documents": "res-documents",
        "document": "res-document",
        "folder": "res-folder",
        "folders": "res-folders",
        "disk": "res-disk",
        "cold-storage": "res-cold-storage",
        "tape-storage": "res-tape-storage",
        "toolkit": "res-toolkit",
        "source-code": "res-source-code",
        "traditional-server": "res-traditional-server",
        "shared-file-system": "res-shared-file-system",
        "users-shared": "res-users",

        # Compute
        "ec2": "amazon-ec2",
        "ec2-instance": "res-amazon-ec2-instance",
        "ec2-instances": "res-amazon-ec2-instances",
        "spot-instance": "res-amazon-ec2-spot-instance",
        "lambda": "aws-lambda",
        "lambda-function": "res-aws-lambda-lambda-function",
        "batch": "aws-batch",
        "ecs": "amazon-elastic-container-service",
        "eks": "amazon-elastic-kubernetes-service",
        "fargate": "amazon-ecs-anywhere",
        "lightsail": "amazon-lightsail",
        "elastic-beanstalk": "aws-elastic-beanstalk",
        "ecr": "amazon-elastic-container-registry",
        "auto-scaling": "amazon-ec2-auto-scaling",
        "autoscaling": "amazon-ec2-auto-scaling",
        "asg": "grp-auto-scaling-group",

        # Storage
        "s3": "amazon-simple-storage-service",
        "s3-glacier": "amazon-simple-storage-service-glacier",
        "glacier": "amazon-simple-storage-service-glacier",
        "efs": "amazon-efs",
        "elastic-file-system": "amazon-efs",
        "ebs": "amazon-elastic-block-store",
        "fsx": "amazon-fsx",
        "backup": "aws-backup",
        "storage-gateway": "aws-storage-gateway",
        "transfer": "aws-transfer-family",

        # Database
        "rds": "amazon-rds",
        "dynamodb": "amazon-dynamodb",
        "aurora": "amazon-aurora",
        "redshift": "amazon-redshift",
        "elasticache": "amazon-elasticache",
        "elasticache-redis": "res-amazon-elasticache-elasticache-for-redis",
        "elasticache-memcached": "res-amazon-elasticache-elasticache-for-memcached",
        "neptune": "amazon-neptune",
        "documentdb": "amazon-documentdb",
        "memorydb": "amazon-memorydb",
        "timestream": "amazon-timestream",
        "keyspaces": "amazon-keyspaces",
        "qldb": "amazon-quantum-ledger-database",
        "rds-master": "res-amazon-rds-multi-az-db-cluster",
        "rds-replica": "res-amazon-rds-amazon-rds-instance-alternate",

        # Networking
        "vpc": "grp-virtual-private-cloud-vpc",
        "region": "grp-region",
        "aws-cloud": "grp-aws-cloud",
        "aws-cloud-logo": "grp-aws-cloud-logo",
        "account": "grp-aws-account",
        "aws-account": "grp-aws-account",
        "public-subnet": "grp-public-subnet",
        "private-subnet": "grp-private-subnet",
        "az": "grp-availability-zone",
        "availability-zone": "grp-availability-zone",
        "auto-scaling-group": "grp-auto-scaling-group",
        "corporate-data-center": "grp-corporate-data-center",
        "data-center": "grp-corporate-data-center",
        "ec2-instance-contents": "grp-ec2-instance-contents",
        "server-contents": "grp-server-contents",
        "spot-fleet": "grp-spot-fleet",

        "internet-gateway": "res-amazon-vpc-internet-gateway",
        "igw": "res-amazon-vpc-internet-gateway",
        "nat-gateway": "res-amazon-vpc-nat-gateway",
        "vpc-endpoint": "res-amazon-vpc-endpoints",
        "vpc-endpoints": "res-amazon-vpc-endpoints",
        "endpoints": "res-amazon-vpc-endpoints",
        "vpn-gateway": "res-amazon-vpc-vpn-gateway",
        "vpn-connection": "res-amazon-vpc-vpn-connection",
        "client-vpn": "aws-client-vpn",
        "site-to-site-vpn": "aws-site-to-site-vpn",
        "transit-gateway": "aws-transit-gateway",
        "router": "res-amazon-vpc-router",
        "route-table": "res-amazon-route-53-route-table",
        "direct-connect": "aws-direct-connect",
        "directconnect": "aws-direct-connect",
        "global-accelerator": "aws-global-accelerator",
        "cloud-map": "aws-cloud-map",
        "app-mesh": "aws-app-mesh",
        "private-link": "aws-privatelink",
        "privatelink": "aws-privatelink",
        "network-firewall": "aws-network-firewall",

        # Edge / DNS / CDN / Load balancing
        "cloudfront": "amazon-cloudfront",
        "route53": "amazon-route-53",
        "route-53": "amazon-route-53",
        "alb": "elastic-load-balancing",
        "elb": "elastic-load-balancing",
        "elb-application": "elastic-load-balancing",
        "load-balancer": "elastic-load-balancing",
        "application-load-balancer": "res-elastic-load-balancing-application-load-balancer",
        "network-load-balancer": "res-elastic-load-balancing-network-load-balancer",
        "gateway-load-balancer": "res-elastic-load-balancing-gateway-load-balancer",
        "classic-load-balancer": "res-elastic-load-balancing-classic-load-balancer",

        # Application integration
        "sqs": "amazon-simple-queue-service",
        "sns": "amazon-simple-notification-service",
        "step-functions": "aws-step-functions",
        "stepfunctions": "aws-step-functions",
        "eventbridge": "amazon-eventbridge",
        "mq": "amazon-mq",

        # Streaming / Big data
        "kinesis": "amazon-kinesis",
        "kinesis-data-streams": "amazon-kinesis-data-streams",
        "kinesis-firehose": "amazon-data-firehose",
        "firehose": "amazon-data-firehose",
        "kinesis-analytics": "amazon-managed-service-for-apache-flink",
        "msk": "amazon-managed-streaming-for-apache-kafka",
        "kafka": "amazon-managed-streaming-for-apache-kafka",
        "athena": "amazon-athena",
        "glue": "aws-glue",
        "emr": "amazon-emr",
        "quicksight": "amazon-quicksight",
        "lake-formation": "aws-lake-formation",
        "datazone": "amazon-datazone",
        "opensearch": "amazon-opensearch-service",
        "elasticsearch": "amazon-opensearch-service",

        # Security / Identity
        "iam": "aws-identity-and-access-management",
        "kms": "aws-key-management-service",
        "secrets-manager": "aws-secrets-manager",
        "cognito": "amazon-cognito",
        "waf": "aws-waf",
        "shield": "aws-shield",
        "guardduty": "amazon-guardduty",
        "inspector": "amazon-inspector",
        "macie": "amazon-macie",
        "security-hub": "aws-security-hub",
        "certificate-manager": "aws-certificate-manager",
        "acm": "aws-certificate-manager",
        "directory-service": "aws-directory-service",
        "verified-access": "aws-verified-access",
        "verified-permissions": "amazon-verified-permissions",
        "private-ca": "aws-private-certificate-authority",
        "iam-identity-center": "aws-iam-identity-center",
        "sso": "aws-iam-identity-center",
        "ram": "aws-resource-access-manager",

        # AI / ML
        "bedrock": "amazon-bedrock",
        "sagemaker": "amazon-sagemaker",
        "sagemaker-ai": "amazon-sagemaker-ai",
        "comprehend": "amazon-comprehend",
        "rekognition": "amazon-rekognition",
        "textract": "amazon-textract",
        "translate": "amazon-translate",
        "polly": "amazon-polly",
        "transcribe": "amazon-transcribe",
        "lex": "amazon-lex",
        "kendra": "amazon-kendra",
        "personalize": "amazon-personalize",
        "forecast": "amazon-forecast",
        "fraud-detector": "amazon-fraud-detector",
        "q": "amazon-q",
        "q-developer": "amazon-q-developer",
        "q-business": "amazon-q-business",

        # Front-end / Mobile / API
        "apigateway": "amazon-api-gateway",
        "api-gateway": "amazon-api-gateway",
        "amplify": "aws-amplify",
        "appsync": "aws-appsync",
        "device-farm": "aws-device-farm",
        "location-service": "amazon-location-service",
        "pinpoint": "aws-end-user-messaging",

        # Compute / Serverless variants
        "step": "aws-step-functions",

        # Email / Comm
        "ses": "amazon-simple-email-service",
        "ses-icon": "amazon-simple-email-service",

        # Management / Governance / Observability
        "cloudwatch": "amazon-cloudwatch",
        "cloudtrail": "aws-cloudtrail",
        "config": "aws-config",
        "systems-manager": "aws-systems-manager",
        "ssm": "aws-systems-manager",
        "control-tower": "aws-control-tower",
        "organizations": "aws-organizations",
        "trusted-advisor": "aws-trusted-advisor",
        "service-catalog": "aws-service-catalog",
        "license-manager": "aws-license-manager",
        "managed-services": "aws-managed-services",
        "well-architected-tool": "aws-well-architected-tool",
        "compute-optimizer": "aws-compute-optimizer",
        "health-dashboard": "aws-health-dashboard",
        "auto-scaling-icon": "aws-auto-scaling",
        "appconfig": "aws-appconfig",
        "xray": "aws-x-ray",
        "x-ray": "aws-x-ray",
        "managed-grafana": "amazon-managed-grafana",
        "managed-prometheus": "amazon-managed-service-for-prometheus",

        # Developer tools / DevOps
        "codebuild": "aws-codebuild",
        "codepipeline": "aws-codepipeline",
        "codedeploy": "aws-codedeploy",
        "codecommit": "aws-codecommit",
        "codeartifact": "aws-codeartifact",
        "codecatalyst": "amazon-codecatalyst",
        "cloudformation": "aws-cloudformation",
        "cdk": "aws-cloud-development-kit",
        "cloud-development-kit": "aws-cloud-development-kit",
        "cloud9": "aws-cloud9",
        "cloudshell": "aws-cloudshell",
        "x-ray-analytics": "aws-x-ray",

        # End-user computing
        "workspaces": "amazon-workspaces-family",
        "workspaces-family": "amazon-workspaces-family",
        "workspaces-core": "res-amazon-workspaces-family-amazon-workspaces-core",
        "workspaces-secure-browser": "res-amazon-workspaces-family-amazon-workspaces-secure-browser",
        "appstream": "amazon-appstream",

        # Migration
        "dms": "aws-database-migration-service",
        "snowball": "aws-snow-family",
        "snowmobile": "aws-snow-family",
        "datasync": "aws-datasync",
        "migration-hub": "aws-migration-hub",
        "application-migration-service": "aws-application-migration-service",

        # Containers / Service Mesh / Microservices
        "ecs-anywhere": "amazon-ecs-anywhere",
        "eks-anywhere": "amazon-eks-anywhere",
        "service-mesh": "aws-app-mesh",
    }
    for alias, target in common.items():
        if target and target in catalog:
            aliases[alias] = target
        elif target is None:
            # try to find dynamically
            hit = find_first(alias.replace("-", ".*"))
            if hit:
                aliases[alias] = hit
    return aliases


def main():
    if DST.exists():
        for sub in ("services", "resources", "groups", "categories"):
            d = DST / sub
            if d.exists():
                shutil.rmtree(d)
    DST.mkdir(parents=True, exist_ok=True)

    catalog = {}
    catalog.update(copy_service_icons())
    catalog.update(copy_resource_icons())
    catalog.update(copy_group_icons())
    catalog.update(copy_category_icons())

    aliases = build_aliases(catalog)

    out = {
        "version": "2025-07-31",
        "icons": catalog,
        "aliases": aliases,
    }

    catalog_path = DST.parent / "catalog.json"
    catalog_path.write_text(json.dumps(out, indent=2, ensure_ascii=False), encoding="utf-8")

    # Stats
    by_kind = {}
    for v in catalog.values():
        by_kind[v["kind"]] = by_kind.get(v["kind"], 0) + 1

    print(f"Wrote {catalog_path}")
    print(f"Total icons: {len(catalog)}")
    for k, n in sorted(by_kind.items()):
        print(f"  {k}: {n}")
    print(f"Aliases: {len(aliases)}")


if __name__ == "__main__":
    main()
