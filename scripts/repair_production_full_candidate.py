from __future__ import annotations

import csv
import hashlib
import html
import json
import math
import re
import shutil
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from urllib.parse import quote, unquote
from xml.sax.saxutils import escape as xml_escape

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from build_school_region_preview import (  # noqa: E402
    SETTINGS_PATH,
    build_school_records,
    generation_phase,
    make_blocker_resistant_body,
    plan_pages,
    read_excel,
    site_join,
)
from school_region_quality_review import INPUT_DEFAULT  # noqa: E402


SOURCE = ROOT / "preview" / "production-full-candidate_20260724_151919"
REPORT_DIR = ROOT / "reports" / "production_full_repair"
HIGH_REPORT = ROOT / "reports" / "production_full_body_similarity_high.csv"
DUPLICATE_REPORT = ROOT / "reports" / "production_full_duplicate_check.csv"
TEMPLATE_REPORT = ROOT / "reports" / "production_full_template_like_pages.csv"
AUDIT_REPORT = ROOT / "reports" / "production_full_candidate_final_audit.json"
CONTENT_FIELDS = (
    "url", "slug", "failure_type", "paired_url", "similarity",
    "normalized_duplicate_group", "substitution_type",
    "current_variant_id", "proposed_variant_id",
)
NAV_FIELDS = (
    "url", "slug", "page_type", "region", "parent_hub",
    "current_depth", "proposed_parent_link",
)


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as source:
        return list(csv.DictReader(source))


def write_csv(path: Path, fields: tuple[str, ...], rows: list[dict[str, object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as target:
        writer = csv.DictWriter(target, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def public_path(category: str, slug: str) -> str:
    return "/" + quote(category, safe="") + "/" + quote(slug, safe="") + "/"


def choose_output() -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output = ROOT / "preview" / f"production-full-candidate-repaired_{stamp}"
    suffix = 2
    while output.exists():
        output = output.with_name(
            f"production-full-candidate-repaired_{stamp}_{suffix}"
        )
        suffix += 1
    return output


def variant(slug: str, failure_type: str, original: int) -> int:
    value = (
        "production-full-repair:" + slug + failure_type + str(original)
    ).encode("utf-8")
    return int(hashlib.sha256(value).hexdigest(), 16) % 100 + 1


def collect_content_targets(
    plans: dict[str, dict[str, object]],
    current_variants: dict[str, int],
) -> tuple[list[dict[str, object]], dict[str, dict[str, object]]]:
    targets: dict[str, dict[str, object]] = {}

    def add(
        slug: str, failure_type: str, paired: str = "",
        similarity: str = "", group: str = "", substitution: str = "",
    ) -> None:
        if slug not in plans:
            return
        row = targets.setdefault(slug, {
            "url": public_path("과외", slug),
            "slug": slug,
            "failure_types": [],
            "pairs": [],
            "similarities": [],
            "groups": [],
            "substitutions": [],
        })
        row["failure_types"].append(failure_type)
        if paired:
            row["pairs"].append(paired)
        if similarity:
            row["similarities"].append(similarity)
        if group:
            row["groups"].append(group)
        if substitution:
            row["substitutions"].append(substitution)

    for row in read_csv(DUPLICATE_REPORT):
        if row["duplicate_type"] != "normalized_body":
            continue
        group = row["value_signature"]
        add(row["page_a"], "normalized_body_exact", row["page_b"], group=group)
        add(row["page_b"], "normalized_body_exact", row["page_a"], group=group)
    for row in read_csv(HIGH_REPORT):
        try:
            score = float(row["similarity"])
        except ValueError:
            continue
        if score < .95:
            continue
        add(row["page_a"], "body_ge_95", row["page_b"], row["similarity"])
        add(row["page_b"], "body_ge_95", row["page_a"], row["similarity"])
    for row in read_csv(TEMPLATE_REPORT):
        kind = "simple_substitution"
        add(
            row["page_a"], kind, row["page_b"], row["similarity"],
            substitution=row.get("page_type", ""),
        )
        add(
            row["page_b"], kind, row["page_a"], row["similarity"],
            substitution=row.get("page_type", ""),
        )

    priority = {
        "normalized_body_exact": 0,
        "body_ge_95": 1,
        "simple_substitution": 2,
    }
    output_rows: list[dict[str, object]] = []
    for slug, row in sorted(targets.items()):
        failure_types = sorted(set(row["failure_types"]), key=priority.get)
        failure = "+".join(failure_types)
        original = current_variants.get(slug, 0)
        proposed = variant(slug, failure, original)
        row["failure_type"] = failure
        row["current_variant_id"] = original
        row["proposed_variant_id"] = proposed
        output_rows.append({
            "url": row["url"],
            "slug": slug,
            "failure_type": failure,
            "paired_url": "|".join(
                public_path("과외", value) for value in sorted(set(row["pairs"]))
            ),
            "similarity": max(row["similarities"], default=""),
            "normalized_duplicate_group": "|".join(sorted(set(row["groups"]))),
            "substitution_type": "|".join(sorted(set(row["substitutions"]))),
            "current_variant_id": original,
            "proposed_variant_id": proposed,
        })
    return output_rows, targets


def replace_body(path: Path, new_body: str) -> None:
    text = path.read_text(encoding="utf-8")
    pattern = re.compile(
        r'(<div class="content">.*?<br>\s*)(.*?)(</div>)',
        re.I | re.S,
    )
    replacement = lambda match: (
        match.group(1) + html.escape(new_body) + match.group(3)
    )
    updated, count = pattern.subn(replacement, text, count=1)
    if count != 1:
        raise RuntimeError(f"본문 영역을 찾을 수 없습니다: {path}")
    path.write_text(updated, encoding="utf-8")


def all_html_paths(output: Path) -> list[tuple[str, str]]:
    rows = []
    for file in output.rglob("*.html"):
        relative = file.relative_to(output)
        if relative == Path("index.html"):
            url = "/"
            slug = ""
        else:
            url = "/" + "/".join(
                quote(part, safe="") for part in relative.parent.parts
            ) + "/"
            slug = relative.parent.name
        rows.append((url, slug))
    return sorted(rows)


def hub_html(
    site_url: str, path: str, title: str,
    description: str, links: list[tuple[str, str]],
) -> str:
    return f"""<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{html.escape(title)}</title>
  <meta name="description" content="{html.escape(description, quote=True)}">
  <link rel="canonical" href="{html.escape(site_join(site_url, path), quote=True)}">
  <meta property="og:type" content="website">
  <meta property="og:title" content="{html.escape(title, quote=True)}">
  <meta property="og:description" content="{html.escape(description, quote=True)}">
  <meta property="og:url" content="{html.escape(site_join(site_url, path), quote=True)}">
  <link rel="stylesheet" href="/assets/site.css">
</head>
<body>
  <header class="site-header"><a href="/">공부업</a></header>
  <main>
    <h1>{html.escape(title)}</h1>
    <p>{html.escape(description)}</p>
    <nav aria-label="전체 운영 탐색">
{chr(10).join(f'      <a href="{href}">{html.escape(label)}</a>' for href, label in links)}
    </nav>
  </main>
  <footer><a href="/">홈으로</a></footer>
</body>
</html>
"""


def create_navigation(
    output: Path,
    existing_pages: list[tuple[str, str]],
    site_url: str,
) -> tuple[list[dict[str, object]], list[str], dict[str, str]]:
    category = "과외"
    batch_size = 30
    group_size = 30
    super_size = 30
    targets = [(url, slug or "홈") for url, slug in existing_pages]
    parent_by_url: dict[str, str] = {}
    new_urls: list[str] = []
    batch_links: list[tuple[str, str]] = []
    for index in range(0, len(targets), batch_size):
        number = index // batch_size + 1
        slug = f"전체운영묶음{number:04d}"
        path = public_path(category, slug)
        links = targets[index:index + batch_size]
        for target, _ in links:
            parent_by_url[target] = path
        file = output / category / slug / "index.html"
        file.parent.mkdir(parents=True, exist_ok=False)
        file.write_text(
            hub_html(
                site_url, path, f"전체 운영 페이지 묶음 {number}",
                "전체 운영 페이지를 30개 이하 단위로 연결합니다.", links,
            ),
            encoding="utf-8",
        )
        batch_links.append((path, f"전체 운영 페이지 묶음 {number}"))
        new_urls.append(site_join(site_url, path))

    group_links: list[tuple[str, str]] = []
    for index in range(0, len(batch_links), group_size):
        number = index // group_size + 1
        slug = f"전체운영그룹{number:03d}"
        path = public_path(category, slug)
        links = batch_links[index:index + group_size]
        file = output / category / slug / "index.html"
        file.parent.mkdir(parents=True, exist_ok=False)
        file.write_text(
            hub_html(
                site_url, path, f"전체 운영 탐색 그룹 {number}",
                "운영 페이지 묶음을 단계별로 탐색합니다.", links,
            ),
            encoding="utf-8",
        )
        group_links.append((path, f"전체 운영 탐색 그룹 {number}"))
        new_urls.append(site_join(site_url, path))

    super_links: list[tuple[str, str]] = []
    for index in range(0, len(group_links), super_size):
        number = index // super_size + 1
        slug = f"전체운영상위그룹{number:02d}"
        path = public_path(category, slug)
        links = group_links[index:index + super_size]
        file = output / category / slug / "index.html"
        file.parent.mkdir(parents=True, exist_ok=False)
        file.write_text(
            hub_html(
                site_url, path, f"전체 운영 상위 그룹 {number}",
                "전체 운영 탐색 그룹을 연결합니다.", links,
            ),
            encoding="utf-8",
        )
        super_links.append((path, f"전체 운영 상위 그룹 {number}"))
        new_urls.append(site_join(site_url, path))

    root_slug = "전체운영탐색"
    root_path = public_path(category, root_slug)
    root_file = output / category / root_slug / "index.html"
    root_file.parent.mkdir(parents=True, exist_ok=False)
    root_file.write_text(
        hub_html(
            site_url, root_path, "전체 운영 페이지 탐색",
            "모든 운영 페이지를 최대 다섯 단계 안에서 탐색합니다.",
            super_links,
        ),
        encoding="utf-8",
    )
    new_urls.append(site_join(site_url, root_path))

    home = output / "index.html"
    home_text = home.read_text(encoding="utf-8")
    insertion = (
        f'\n  <aside class="full-navigation">'
        f'<a href="{root_path}">전체 운영 페이지 탐색</a></aside>\n'
    )
    if root_path not in home_text:
        home_text = home_text.replace("</main>", insertion + "</main>", 1)
        home.write_text(home_text, encoding="utf-8")

    sitemap_name = "sitemap-navigation-repair-001.xml"
    sitemap = output / sitemap_name
    sitemap.write_text(
        "\n".join([
            '<?xml version="1.0" encoding="UTF-8"?>',
            '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">',
            *[
                f"  <url><loc>{xml_escape(url)}</loc></url>"
                for url in new_urls
            ],
            "</urlset>",
            "",
        ]),
        encoding="utf-8",
    )
    index = output / "sitemap_index.xml"
    index_text = index.read_text(encoding="utf-8")
    entry = (
        f"  <sitemap><loc>{xml_escape(site_join(site_url, '/' + sitemap_name))}"
        "</loc></sitemap>\n"
    )
    if sitemap_name not in index_text:
        index_text = index_text.replace("</sitemapindex>", entry + "</sitemapindex>")
        index.write_text(index_text, encoding="utf-8")
    return (
        [
            {
                "url": url,
                "slug": slug,
                "page_type": "existing",
                "region": "",
                "parent_hub": parent_by_url.get(url, ""),
                "current_depth": 99,
                "proposed_parent_link": parent_by_url.get(url, ""),
            }
            for url, slug in existing_pages
        ],
        new_urls,
        parent_by_url,
    )


def main() -> int:
    if not SOURCE.is_dir():
        raise FileNotFoundError(SOURCE)
    settings = json.loads(SETTINGS_PATH.read_text(encoding="utf-8-sig"))
    config = dict(settings["school_region_generation"])
    links, regions, _, _ = read_excel(INPUT_DEFAULT)
    schools = build_school_records(links)
    plans, _, _ = plan_pages(regions, links, schools, config)
    enabled = set(str(value) for value in config["enabled_generation_phases"])
    plans = [plan for plan in plans if generation_phase(plan) in enabled]
    by_slug = {str(plan["slug"]): plan for plan in plans}
    current_variants = {
        str(plan["slug"]): int(plan.get("similarity_variant", 0))
        for plan in plans
    }
    content_rows, target_info = collect_content_targets(
        by_slug, current_variants
    )
    write_csv(REPORT_DIR / "repair_content_targets.csv", CONTENT_FIELDS, content_rows)

    output = choose_output()
    shutil.copytree(SOURCE, output, copy_function=shutil.copy2)
    modified_content = 0
    for row in content_rows:
        slug = str(row["slug"])
        plan = dict(by_slug[slug])
        plan["similarity_variant"] = int(row["proposed_variant_id"])
        body = make_blocker_resistant_body(plan)
        file = output / "과외" / slug / "index.html"
        replace_body(file, body)
        modified_content += 1

    existing_pages = all_html_paths(output)
    nav_rows, new_hub_urls, _ = create_navigation(
        output, existing_pages, str(settings["site_url"]).rstrip("/")
    )
    audit = json.loads(AUDIT_REPORT.read_text(encoding="utf-8"))
    nav_rows = [
        row for row in nav_rows
        if row["url"] != "/" and (
            int(audit["unreachable_from_home_count"]) > 0
            or int(audit["orphan_page_count"]) > 0
        )
    ]
    write_csv(REPORT_DIR / "repair_navigation_targets.csv", NAV_FIELDS, nav_rows)
    summary = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "source_candidate": str(SOURCE),
        "repaired_candidate": str(output),
        "content_target_count": len(content_rows),
        "navigation_target_count": len(nav_rows),
        "new_hub_count": len(new_hub_urls),
        "modified_content_html_count": modified_content,
        "modified_navigation_html_count": len(new_hub_urls) + 1,
        "source_html_count": len(existing_pages),
        "expected_repaired_html_count": len(existing_pages) + len(new_hub_urls),
        "source_sitemap_url_count": audit["sitemap_url_count"],
        "expected_repaired_sitemap_url_count": (
            int(audit["sitemap_url_count"]) + len(new_hub_urls)
        ),
    }
    (REPORT_DIR / "repair_execution_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
