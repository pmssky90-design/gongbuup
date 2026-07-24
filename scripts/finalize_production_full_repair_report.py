from __future__ import annotations

import json
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
REPORT_DIR = ROOT / "reports" / "production_full_repair"


def load(path: Path) -> dict[str, object]:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def main() -> int:
    execution = load(REPORT_DIR / "repair_execution_summary.json")
    similarity = load(REPORT_DIR / "repair_similarity_validation.json")
    audit = load(REPORT_DIR / "repair_final_audit.json")
    invariant = load(REPORT_DIR / "repair_invariant_validation.json")
    navigation = load(REPORT_DIR / "repair_navigation_target_summary.json")
    preflight = load(ROOT / "reports" / "production_full_preflight.json")
    preview = load(ROOT / "reports" / "production_full_repair_preview.json")
    baseline = load(ROOT / "reports" / "production_full_candidate_final_audit.json")

    content_modified = int(execution["modified_content_html_count"])
    existing_html = int(execution["source_html_count"])
    navigation_modified_existing = 1  # 기존 홈에 전체 탐색 루트 링크 추가
    summary = {
        "source_candidate": execution["source_candidate"],
        "repaired_candidate": execution["repaired_candidate"],
        "content_repair_target_page_count": execution["content_target_count"],
        "navigation_repair_target_page_count": navigation[
            "navigation_repair_target_count"
        ],
        "new_hub_count": execution["new_hub_count"],
        "modified_existing_html_count": (
            content_modified + navigation_modified_existing
        ),
        "unmodified_existing_html_count": (
            existing_html - content_modified - navigation_modified_existing
        ),
        "new_html_count": execution["new_hub_count"],
        "final_html_count": audit["html_count"],
        "body_ge_95_before": baseline["body_ge_95"],
        "body_ge_95_after": audit["body_ge_95"],
        "exact_body_before": baseline["exact_body"],
        "exact_body_after": audit["exact_body"],
        "normalized_duplicate_before": baseline["exact_normalized_body"],
        "normalized_duplicate_after": audit["exact_normalized_body"],
        "simple_replacement_before": baseline["simple_replacement"],
        "simple_replacement_after": audit["simple_replacement"],
        "orphan_page_before": baseline["orphan_page_count"],
        "orphan_page_after": audit["orphan_page_count"],
        "unreachable_from_home_before": baseline["unreachable_from_home_count"],
        "unreachable_from_home_after": audit["unreachable_from_home_count"],
        "maximum_click_depth_after": audit["home_max_click_depth"],
        "sitemap_url_count": audit["sitemap_url_count"],
        "body_90_to_95_after": audit["body_90_to_95"],
        "body_80_to_90_after": audit["body_80_to_90"],
        "broken_internal_link_count": audit["broken_internal_link_count"],
        "broken_image_count": audit["broken_image_count"],
        "sitemap_html_mismatch_count": audit["sitemap_html_mismatch_count"],
        "canonical_error_count": audit["canonical_error_count"],
        "jsonld_error_count": audit["jsonld_error_count"],
        "h1_error_count": audit["h1_error_count"],
        "school_region_error_count": invariant["school_region_error_count"],
        "forbidden_combination_count": invariant[
            "forbidden_combination_count"
        ],
        "required_keyword_error_count": invariant[
            "required_keyword_error_count"
        ],
        "determinism_mismatch_count": invariant[
            "body_determinism_mismatch_count"
        ],
        "immutable_metadata_change_error_count": invariant[
            "title_description_canonical_h1_change_error_count"
        ],
        "preflight_passed": preflight["passed"],
        "preflight_html_generation_allowed": preflight[
            "html_generation_allowed"
        ],
        "final_passed": bool(
            similarity["passed"]
            and audit["passed"]
            and invariant["passed"]
            and preflight["passed"]
        ),
        "preview_url": preview["url"],
        "preview_pid": preview["pid"],
    }
    (REPORT_DIR / "repair_final_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    lines = [
        "Production Full 증분 보정 최종 보고",
        "=" * 48,
        *[f"{key}: {value}" for key, value in summary.items()],
    ]
    (REPORT_DIR / "repair_final_summary.txt").write_text(
        "\n".join(lines) + "\n", encoding="utf-8"
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if summary["final_passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
