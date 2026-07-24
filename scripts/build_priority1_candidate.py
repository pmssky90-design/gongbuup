from __future__ import annotations

import csv
import hashlib
import json
import re
import socket
import sys
import time
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from urllib.parse import unquote, urlparse

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from build_school_region_preview import (  # noqa: E402
    REPORTS,
    SETTINGS_PATH,
    build_school_records,
    generate_preview,
    generation_phase,
    plan_pages,
    read_excel,
    validate_preview,
    write_csv,
)
from fix_priority1_blockers import (  # noqa: E402
    entity,
    normalized_signature,
    score,
)
from school_region_quality_review import (  # noqa: E402
    INPUT_DEFAULT,
    normalized,
    priority_of,
)

BASE_OUTPUT = ROOT / "preview" / "production-priority-1-candidate"
ADDITIONAL_VARIANTS = ROOT / "config" / "priority1_candidate_additional_variants.json"


def choose_output() -> Path:
    if not BASE_OUTPUT.exists():
        return BASE_OUTPUT
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    candidate = BASE_OUTPUT.with_name(f"{BASE_OUTPUT.name}_{stamp}")
    suffix = 2
    while candidate.exists():
        candidate = BASE_OUTPUT.with_name(f"{BASE_OUTPUT.name}_{stamp}_{suffix}")
        suffix += 1
    return candidate


def html_target(output: Path, url: str) -> Path:
    path = unquote(urlparse(url).path).strip("/")
    return output / Path(path.replace("/", "\\")) / "index.html" if path else output / "index.html"


def write_json(name: str, value: object) -> None:
    (REPORTS / name).write_text(
        json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )


def main() -> int:
    started = time.perf_counter()
    settings = json.loads(SETTINGS_PATH.read_text(encoding="utf-8-sig"))
    config = dict(settings["school_region_generation"])
    variant_path = ROOT / str(config["priority1_blocker_variant_file"])
    variants = json.loads(variant_path.read_text(encoding="utf-8"))
    if len(variants) != 530:
        raise RuntimeError(f"priority1 blocker mapping count must be 530, got {len(variants)}")

    links, regions, _, _ = read_excel(INPUT_DEFAULT)
    schools = build_school_records(links)
    plans, duplicate_rows, school_validation = plan_pages(
        regions, links, schools, config
    )
    additional_variants: dict[str, int] = {}
    if ADDITIONAL_VARIANTS.is_file():
        additional_variants = json.loads(
            ADDITIONAL_VARIANTS.read_text(encoding="utf-8")
        )
        for plan in plans:
            slug = str(plan["slug"])
            if slug in additional_variants and slug not in variants:
                plan["blocker_variant"] = True
                plan["similarity_variant"] = int(additional_variants[slug])
    enabled = set(str(value) for value in config["enabled_generation_phases"])
    priority_plans = [
        plan for plan in plans
        if generation_phase(plan) in enabled and priority_of(plan) == 1
    ]
    priority_plans.sort(key=lambda row: str(row["slug"]))
    if len(priority_plans) != 2412:
        raise RuntimeError(f"priority1 plan count must be 2412, got {len(priority_plans)}")

    output = choose_output()
    generated = generate_preview(
        output,
        priority_plans,
        sorted({str(plan.get("지역명", "")) for plan in priority_plans}),
        schools,
        links,
        settings,
        config,
    )
    structural = validate_preview(output, generated, schools, settings)
    records = list(generated["records"])

    # Sitemap URLs are checked against actual candidate files.
    sitemap_rows: list[dict[str, object]] = []
    sitemap_urls: list[str] = []
    for sitemap in sorted(output.glob("sitemap-*.xml")):
        urls = re.findall(
            r"<loc>(.*?)</loc>", sitemap.read_text(encoding="utf-8"), re.DOTALL
        )
        for url in urls:
            target = html_target(output, url)
            sitemap_urls.append(url)
            sitemap_rows.append({
                "sitemap": sitemap.name,
                "url": url,
                "html_path": str(target),
                "html_exists": target.is_file(),
            })

    content_paths = [str(record["path"]) for record in records]
    content_slugs = [str(record["slug"]) for record in records]
    content_canonicals = [str(record["canonical"]) for record in records]
    actual_canonical_errors = sum(
        str(record["canonical"])
        != settings["site_url"].rstrip("/") + str(record["path"])
        for record in records
    )

    exact_groups: dict[str, list[str]] = defaultdict(list)
    normalized_groups: dict[str, list[str]] = defaultdict(list)
    by_type: dict[str, list[dict[str, object]]] = defaultdict(list)
    for record in records:
        slug = str(record["slug"])
        body = str(record["body"])
        exact_groups[hashlib.sha256(body.encode("utf-8")).hexdigest()].append(slug)
        normalized_groups[normalized_signature(body, record)].append(slug)
        by_type[str(record["page_type"])].append(record)

    exact_duplicates = [group for group in exact_groups.values() if len(group) > 1]
    normalized_duplicates = [
        group for group in normalized_groups.values() if len(group) > 1
    ]
    high_rows: list[dict[str, object]] = []
    manual_rows: list[dict[str, object]] = []
    statistical_rows: list[dict[str, object]] = []
    template_rows: list[dict[str, object]] = []
    for page_type, group in sorted(by_type.items()):
        for index, left in enumerate(group):
            for right in group[index + 1:]:
                similarity = score(
                    str(left["body"]), str(right["body"]), left, right
                )
                row = {
                    "page_a": left["slug"],
                    "page_b": right["slug"],
                    "page_type": page_type,
                    "similarity": round(similarity, 6),
                }
                if similarity >= .95:
                    high_rows.append(row)
                elif similarity >= .90:
                    manual_rows.append({**row, "classification": "수동 검토"})
                elif similarity >= .80:
                    statistical_rows.append({**row, "classification": "통계"})
                if normalized(
                    str(left["body"]), (entity(left), entity(right))
                ) == normalized(
                    str(right["body"]), (entity(left), entity(right))
                ):
                    template_rows.append({**row, "reason": "단순 엔터티 치환형"})

    duplicate_slug_count = len(content_slugs) - len(set(content_slugs))
    duplicate_url_count = len(content_paths) - len(set(content_paths))
    duplicate_canonical_count = len(content_canonicals) - len(set(content_canonicals))
    sitemap_missing_html = sum(not bool(row["html_exists"]) for row in sitemap_rows)
    forbidden_count = int(structural["elementary_internal_exam_count"]) + int(
        structural["school_invalid_keyword_count"]
    ) + int(structural["school_unnatural_grade_title_count"])
    keyword_error_count = (
        int(structural["title_keyword_missing_count"])
        + int(structural["description_keyword_missing_count"])
        + int(structural["body_keyword_missing_count"])
    )
    blocker_counts = {
        "body_ge_95_count": len(high_rows),
        "exact_body_duplicate_group_count": len(exact_duplicates),
        "exact_normalized_body_duplicate_group_count": len(normalized_duplicates),
        "simple_replacement_pair_count": len(template_rows),
        "duplicate_slug_count": duplicate_slug_count,
        "duplicate_url_count": duplicate_url_count,
        "duplicate_canonical_count": duplicate_canonical_count,
        "canonical_actual_url_mismatch_count": actual_canonical_errors,
        "sitemap_missing_html_count": sitemap_missing_html,
        "orphan_page_count": int(structural["orphan_page_count"]),
        "broken_internal_link_count": int(structural["broken_internal_link_count"]),
        "broken_image_count": int(structural["broken_image_count"]),
        "forbidden_combination_count": forbidden_count,
        "required_keyword_error_count": keyword_error_count,
    }
    determinism = json.loads(
        (REPORTS / "priority1_blocker_invariant_validation.json").read_text(
            encoding="utf-8"
        )
    )
    determinism_errors = (
        int(determinism["body_determinism_mismatch_count"])
        + int(determinism["fresh_meta_mismatch_count"])
        + int(determinism["title_change_count"])
        + int(determinism["description_change_count"])
        + int(determinism["slug_change_count"])
    )
    blocker_counts["determinism_mismatch_count"] = determinism_errors
    # With a 30-link ceiling, 2,412 important pages cannot all be at depth <= 2
    # (the mathematical capacity is at most 30 * 30). Keep that legacy metric
    # visible as a warning while enforcing the requested <= 3 total depth.
    structural_blocking_errors = [
        error for error in structural["errors"]
        if not str(error).startswith("important_pages_over_two_clicks=")
    ]
    passed = not any(blocker_counts.values()) and not structural_blocking_errors

    html_files = list(output.rglob("*.html"))
    output_bytes = sum(path.stat().st_size for path in output.rglob("*") if path.is_file())
    summary = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "candidate_output": str(output),
        "priority1_content_page_count": len(records),
        "navigation_page_count": len(generated["navigation_pages"]) + 1,
        "html_count": len(html_files),
        "sitemap_url_count": len(sitemap_urls),
        "sitemap_url_unique_count": len(set(sitemap_urls)),
        "output_size_bytes": output_bytes,
        "blocker_variant_mapping_count": len(variants),
        "candidate_additional_variant_mapping_count": len(additional_variants),
        "body_90_to_95_manual_review_pair_count": len(manual_rows),
        "body_80_to_90_statistical_pair_count": len(statistical_rows),
        "blocker_counts": blocker_counts,
        "structural_validation_passed": not structural_blocking_errors,
        "structural_warning_count": (
            int(structural["important_pages_over_two_clicks"])
        ),
        "structural_blocking_errors": structural_blocking_errors,
        "passed": passed,
        "generation_seconds": generated["duration_seconds"],
        "total_validation_seconds": round(time.perf_counter() - started, 3),
        "site_modified": False,
        "production_executed": False,
        "git_used": False,
        "deployed": False,
    }

    write_csv(
        REPORTS / "production_priority_1_candidate_sitemap_urls.csv",
        ("sitemap", "url", "html_path", "html_exists"),
        sitemap_rows,
    )
    write_csv(
        REPORTS / "production_priority_1_candidate_similarity_95.csv",
        ("page_a", "page_b", "page_type", "similarity"),
        high_rows,
    )
    write_csv(
        REPORTS / "production_priority_1_candidate_similarity_90_95.csv",
        ("page_a", "page_b", "page_type", "similarity", "classification"),
        manual_rows,
    )
    write_csv(
        REPORTS / "production_priority_1_candidate_similarity_80_90.csv",
        ("page_a", "page_b", "page_type", "similarity", "classification"),
        statistical_rows,
    )
    write_csv(
        REPORTS / "production_priority_1_candidate_template_like.csv",
        ("page_a", "page_b", "page_type", "similarity", "reason"),
        template_rows,
    )
    write_csv(
        REPORTS / "production_priority_1_candidate_html_sizes.csv",
        ("url_path", "click_depth", "internal_link_count", "html_size_bytes", "is_orphan"),
        list(structural["page_metrics"]),
    )
    write_json("production_priority_1_candidate_summary.json", summary)
    write_json("production_priority_1_candidate_validation.json", {
        "passed": passed,
        "blocker_counts": blocker_counts,
        "structural_validation": structural,
        "similarity_policy": {
            "body_ge_95": "실패",
            "normalized_body_exact_duplicate": "실패",
            "simple_replacement": "실패",
            "body_90_to_95": "수동 검토, 자동 실패 제외",
            "body_80_to_90": "통계, 자동 실패 제외",
        },
    })
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if passed else 1


if __name__ == "__main__":
    raise SystemExit(main())
