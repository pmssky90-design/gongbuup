from __future__ import annotations

import json
import re
import sys
from collections import Counter
from pathlib import Path
from urllib.parse import quote

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from build_school_region_preview import (  # noqa: E402
    SETTINGS_PATH,
    build_school_records,
    generation_phase,
    plan_pages,
    read_excel,
)
from full_plan_validation import (  # noqa: E402
    compact_normalized_body,
    duplicate_reports,
    generate_inventory,
)
from school_region_quality_review import INPUT_DEFAULT  # noqa: E402


REPORT = ROOT / "reports" / "production_full_preflight.json"
INVALID = re.compile(r'[<>:"/\\|?*]')
RESERVED = re.compile(
    r"^(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\..*)?$",
    re.I,
)


def main() -> int:
    settings = json.loads(SETTINGS_PATH.read_text(encoding="utf-8-sig"))
    config = settings["school_region_generation"]
    links, regions, _, _ = read_excel(INPUT_DEFAULT)
    schools = build_school_records(links)
    plans, removed_duplicates, _ = plan_pages(regions, links, schools, config)
    enabled = set(str(value) for value in config["enabled_generation_phases"])
    plans = [row for row in plans if generation_phase(row) in enabled]

    slugs = [str(row["slug"]) for row in plans]
    urls = [
        "/" + quote(str(settings["page_category"]), safe="") + "/"
        + quote(slug, safe="") + "/"
        for slug in slugs
    ]
    canonicals = [str(settings["site_url"]).rstrip("/") + url for url in urls]
    invalid_character = sum(bool(INVALID.search(slug)) for slug in slugs)
    reserved_name = sum(bool(RESERVED.match(slug)) for slug in slugs)
    trailing_dot_space = sum(slug.endswith((".", " ")) for slug in slugs)
    # Windows 장기 경로 여유를 남긴 240자 기준으로 생성 전 차단한다.
    output_root = ROOT / str(config.get("production_preflight_output", "preview"))
    path_length = sum(
        len(str(output_root / str(settings["page_category"]) / slug / "index.html")) > 240
        for slug in slugs
    )

    inventory, _ = generate_inventory(plans, settings)
    duplicate_summary, _ = duplicate_reports(inventory)
    compact_counts = Counter(compact_normalized_body(row) for row in inventory)
    simple_replacement = sum(
        count - 1 for count in compact_counts.values() if count > 1
    )

    # 30진 계층: 홈→루트→상위→그룹→묶음→콘텐츠, 최대 5단계.
    total_planned_html = len(plans) + 2729
    batch_count = (total_planned_html + 29) // 30
    group_count = (batch_count + 29) // 30
    super_count = (group_count + 29) // 30
    expected_max_depth = 5
    expected_orphan = 0 if super_count <= 30 else total_planned_html
    result = {
        "plan_count": len(plans),
        "windows_invalid_character_count": invalid_character,
        "windows_reserved_name_count": reserved_name,
        "windows_trailing_dot_space_count": trailing_dot_space,
        "path_over_240_count": path_length,
        "slug_duplicate_count": len(slugs) - len(set(slugs)),
        "url_duplicate_count": len(urls) - len(set(urls)),
        "canonical_duplicate_count": len(canonicals) - len(set(canonicals)),
        "removed_duplicate_plan_count": len(removed_duplicates),
        "expected_navigation_batch_count": batch_count,
        "expected_navigation_group_count": group_count,
        "expected_navigation_super_group_count": super_count,
        "expected_max_click_depth": expected_max_depth,
        "expected_orphan_page_count": expected_orphan,
        "planned_exact_body_duplicate_count": duplicate_summary[
            "duplicate_body_count"
        ],
        "planned_exact_normalized_body_duplicate_count": duplicate_summary[
            "duplicate_normalized_body_count"
        ],
        "planned_simple_replacement_count": simple_replacement,
    }
    blockers = (
        "windows_invalid_character_count", "windows_reserved_name_count",
        "windows_trailing_dot_space_count", "path_over_240_count",
        "slug_duplicate_count", "url_duplicate_count",
        "canonical_duplicate_count", "removed_duplicate_plan_count",
        "expected_orphan_page_count", "planned_exact_body_duplicate_count",
        "planned_exact_normalized_body_duplicate_count",
        "planned_simple_replacement_count",
    )
    result["passed"] = all(result[key] == 0 for key in blockers)
    result["html_generation_allowed"] = result["passed"]
    result["blocking_behavior"] = (
        "PASS일 때만 HTML 생성을 시작합니다."
        if result["passed"]
        else "FAIL: HTML 생성을 시작하지 않습니다."
    )
    REPORT.write_text(
        json.dumps(result, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
