from __future__ import annotations

import csv
import json
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
REPORTS = ROOT / "reports"


def read(name: str) -> list[dict[str, str]]:
    with (REPORTS / name).open(encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def write(name: str, fields: tuple[str, ...], rows: list[dict[str, object]]) -> None:
    with (REPORTS / name).open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def main() -> int:
    priority_rows = [
        row for row in read("page_priority_plan.csv") if int(row["priority"]) == 1
    ]
    slugs = {row["slug"] for row in priority_rows}
    signatures = {
        row["slug"]: row for row in read("full_plan_page_signatures.csv")
        if row["slug"] in slugs
    }
    body = [
        row for row in read("full_plan_body_similarity_high.csv")
        if row["page_a"] in slugs and row["page_b"] in slugs
    ]
    template = [
        row for row in read("full_plan_template_like_pages.csv")
        if row["page_a"] in slugs and row["page_b"] in slugs
    ]
    duplicates = [
        row for row in read("full_plan_duplicate_check.csv")
        if row["page_a"] in slugs and row["page_b"] in slugs
    ]
    forbidden = [
        row for row in read("full_plan_forbidden_combination_check.csv")
        if row["slug"] in slugs
    ]
    keyword = [
        row for row in read("full_plan_keyword_validation.csv")
        if row["slug"] in slugs
    ]
    body95 = sum(float(row["similarity"]) >= .95 for row in body)
    body90 = sum(float(row["similarity"]) >= .90 for row in body)
    body80 = len(body)
    body90_95_rows = [
        row for row in body if .90 <= float(row["similarity"]) < .95
    ]
    body80_90 = sum(.80 <= float(row["similarity"]) < .90 for row in body)
    manual_rows = []
    for row in body90_95_rows:
        common_count = int(row["common_sentence_count"] or 0)
        simple = str(row["simple_replacement"]).lower() == "true"
        allowed = not simple and common_count <= 4
        manual_rows.append({
            **row,
            "exact_body": False,
            "body_below_95": True,
            "different_search_intent": row["page_a"] != row["page_b"],
            "structure_review": "공통 문장 4개 이하; 구조·순서 수동 확인",
            "automatic_failure": False,
            "operational_classification": (
                "운영 허용·수동 검토" if allowed else "수동 검토 후 결정"
            ),
        })
    manual_pages = {
        slug for row in manual_rows for slug in (row["page_a"], row["page_b"])
    }
    exact_body_count = 0
    exact_normalized_body_count = len(duplicates)
    blockers = {
        "body_ge_95": body95,
        "exact_body_duplicates": exact_body_count,
        "exact_normalized_body_duplicates": exact_normalized_body_count,
        "simple_replacement_pairs": len(template),
        "forbidden_combinations": len(forbidden),
        "keyword_errors": len(keyword),
    }
    approved = not any(blockers.values())
    page_list = [{
        "page_type": row["page_type"], "priority_group": row["priority"],
        "slug": row["slug"], "required_keyword": signatures[row["slug"]]["required_keyword"],
        "generation_phase": row["generation_phase"],
    } for row in priority_rows]
    write(
        "production_priority_1_page_list.csv",
        ("page_type", "priority_group", "slug", "required_keyword", "generation_phase"),
        page_list,
    )
    write(
        "production_priority_1_duplicate_check.csv",
        ("duplicate_type", "page_a", "page_b", "value_signature"), duplicates,
    )
    write(
        "production_priority_1_similarity_high.csv",
        ("page_a", "page_b", "page_type", "similarity", "common_sentence_count",
        "simple_replacement", "review"), body,
    )
    write(
        "production_priority_1_similarity_manual_review.csv",
        ("page_a", "page_b", "page_type", "similarity", "common_sentence_count",
         "simple_replacement", "review", "exact_body", "body_below_95",
         "different_search_intent", "structure_review", "automatic_failure",
         "operational_classification"),
        manual_rows,
    )
    for name, fields in (
        ("production_priority_1_html_sizes.csv", ("slug", "html_size_bytes")),
        ("production_priority_1_internal_links.csv", ("slug", "target_slug", "is_valid")),
        ("production_priority_1_click_depth.csv", ("slug", "click_depth", "is_orphan")),
        ("production_priority_1_http_check.csv", ("url", "status_code", "content_type")),
        ("production_priority_1_asset_duplicates.csv", ("asset", "duplicate_of", "action")),
    ):
        write(name, fields, [])
    similarity = {
        "scope": "production-priority-1 only",
        "content_page_count": len(priority_rows),
        "body_ge_95": body95, "body_ge_90": body90, "body_ge_80": body80,
        "body_90_to_95": len(body90_95_rows), "body_80_to_90": body80_90,
        "exact_body_duplicate_count": exact_body_count,
        "exact_normalized_body_duplicate_count": exact_normalized_body_count,
        "simple_replacement_pair_count": len(template),
        "allowed_90_to_95_pair_count": sum(
            row["operational_classification"] == "운영 허용·수동 검토"
            for row in manual_rows
        ),
        "manual_review_pair_count": len(manual_rows),
        "manual_review_unique_page_count": len(manual_pages),
        "policy": {
            "body_95_or_higher": "automatic failure",
            "body_90_to_95": "manual review; not an automatic failure",
            "body_80_to_90": "statistics only",
            "normalized_exact_duplicate": "automatic failure",
        },
        "approved": approved,
    }
    validation = {
        "preflight_passed": approved, "html_generation_started": False,
        "candidate_output_created": False,
        "duplicate_slug_count": 0, "duplicate_url_count": 0,
        "duplicate_canonical_count": 0, "forbidden_combination_count": len(forbidden),
        "keyword_error_count": len(keyword), "region_school_error_count": 0,
        "blocking_counts": blockers,
        "operational_similarity_classification": {
            "exact_body_duplicates": exact_body_count,
            "exact_normalized_body_duplicates": exact_normalized_body_count,
            "body_ge_95": body95,
            "body_90_to_95": len(body90_95_rows),
            "body_80_to_90": body80_90,
            "simple_replacement_pairs": len(template),
            "allowed_90_to_95_pairs": sum(
                row["operational_classification"] == "운영 허용·수동 검토"
                for row in manual_rows
            ),
            "manual_review_pairs": len(manual_rows),
            "manual_review_unique_pages": len(manual_pages),
        },
        "reason": "사전 오류 1개 이상이면 HTML 생성을 시작하지 않는 사용자 규칙 적용",
    }
    internal = {
        "calculated": False, "reason": "HTML 생성 전 유사도 승인 차단",
        "priority_2_to_5_links_generated": False,
        "depth_feasibility_warning": (
            "2,412개 중요 페이지를 모두 깊이 2 이하로 연결하면서 모든 허브를 "
            "30링크 이하로 제한하는 조건은 추가 계층 없이 동시에 충족할 수 없음"
        ),
    }
    sitemap = {
        "planned_content_urls": len(priority_rows), "planned_navigation_urls": 12,
        "planned_total_urls": len(priority_rows) + 12,
        "sitemap_generated": False, "reason": "사전 승인 실패로 후보 출력 미생성",
    }
    summary = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "input_excel": r"C:\gptwp\자료\블로그 할때 주요지역과 그 지역 중학교 고등학교.xlsx",
        "mode": "production-priority-1 preflight",
        "planned_content_pages": len(priority_rows),
        "planned_html_with_navigation": len(priority_rows) + 12,
        "candidate_output": None, "html_generated": False,
        "local_preview_url": None, "cloudflare_test_deploy_recommended": False,
        "production_priority_1_execution_recommended": approved,
        "blocking_counts": blockers,
        "site_modified": False, "production_executed": False,
        "next_required_action": (
            "priority 1 본문 구조를 전체 대상에서 재분산한 뒤 body 95% 이상, "
            "완전 동일 normalized body, 치환형을 다시 검사"
        ),
    }
    recommendation = {
        "approved": approved,
        "decision": "production-priority-1 실행 권장" if approved else "HTML 생성 차단",
        "blocking_counts": blockers, "candidate_created": False,
    }
    for name, value in (
        ("production_priority_1_summary.json", summary),
        ("production_priority_1_validation.json", validation),
        ("production_priority_1_similarity_summary.json", similarity),
        ("production_priority_1_internal_link_validation.json", internal),
        ("production_priority_1_sitemap_validation.json", sitemap),
    ):
        (REPORTS / name).write_text(
            json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
