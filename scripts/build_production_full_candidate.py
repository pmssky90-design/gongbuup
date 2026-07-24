from __future__ import annotations

import json
import re
import subprocess
import sys
import time
from collections import Counter
from datetime import datetime
from pathlib import Path

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
from full_plan_validation import (  # noqa: E402
    compact_normalized_body,
    duplicate_reports,
    sha,
    similarity_scan,
)
from school_region_quality_review import INPUT_DEFAULT, normalized  # noqa: E402


OLD_NAME = "이화여자대학교사범대학부속이화?금란중학교"
OFFICIAL_NAME = "이화여자대학교사범대학부속이화·금란중학교"
BASE_NAME = "production-full-candidate"


def choose_output() -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output = ROOT / "preview" / f"{BASE_NAME}_{stamp}"
    suffix = 2
    while output.exists():
        output = ROOT / "preview" / f"{BASE_NAME}_{stamp}_{suffix}"
        suffix += 1
    return output


def preflight(plans: list[dict[str, object]], normalization_info: dict[str, int]) -> dict[str, object]:
    targets = [row for row in plans if row.get("학교명") == OFFICIAL_NAME]
    slugs = [str(row["slug"]) for row in plans]
    urls = ["/과외/" + slug + "/" for slug in slugs]
    canonicals = ["https://example.local" + url for url in urls]
    result = {
        "plan_count": len(plans),
        "corrected_source_row_count": normalization_info.get("school_name_corrected", 0),
        "old_name_residual_count": sum(
            OLD_NAME in str(value)
            for row in plans
            for value in row.values()
        ),
        "question_mark_slug_count": sum("?" in slug for slug in slugs),
        "corrected_school_page_count": len(targets),
        "corrected_school_slugs": [str(row["slug"]) for row in targets],
        "corrected_school_slug_duplicate_count": len(targets) - len(set(
            str(row["slug"]) for row in targets
        )),
        "slug_duplicate_count": len(slugs) - len(set(slugs)),
        "url_duplicate_count": len(urls) - len(set(urls)),
        "canonical_duplicate_count": len(canonicals) - len(set(canonicals)),
        "school_url_mapping_error_count": sum(
            OFFICIAL_NAME not in str(row["slug"]) for row in targets
        ),
    }
    required_zero = (
        "old_name_residual_count", "question_mark_slug_count",
        "corrected_school_slug_duplicate_count", "slug_duplicate_count",
        "url_duplicate_count", "canonical_duplicate_count",
        "school_url_mapping_error_count",
    )
    result["passed"] = (
        len(plans) == 64104
        and len(targets) == 6
        and all(result[key] == 0 for key in required_zero)
    )
    return result


def inventory_from_records(records: list[dict[str, object]]) -> list[dict[str, object]]:
    rows = []
    for index, record in enumerate(records, start=1):
        body = str(record["body"])
        normalized_body = normalized(
            body,
            (
                str(record["지역명"]), str(record["학교명"]),
                str(record["과목표현"]), str(record["학년표현"]),
            ),
        )
        rows.append({
            **record,
            "page_id": index,
            "URL": record["canonical"],
            "required_keyword": record["keyword"],
            "body_signature": sha(body),
            "normalized_body": normalized_body,
            "normalized_body_signature": sha(normalized_body),
            "title_signature": sha(str(record["title"])),
            "description_signature": sha(str(record["description"])),
        })
    return rows


def json_write(path: Path, value: object) -> None:
    path.write_text(
        json.dumps(value, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def main() -> int:
    started = time.perf_counter()
    preflight = subprocess.run(
        [sys.executable, str(ROOT / "scripts" / "production_full_preflight.py")],
        cwd=ROOT,
        check=False,
    )
    if preflight.returncode != 0:
        raise RuntimeError(
            "Production preflight FAIL: HTML 생성을 시작하지 않습니다."
        )
    settings = json.loads(SETTINGS_PATH.read_text(encoding="utf-8-sig"))
    config = dict(settings["school_region_generation"])
    links, regions, _, normalization_info = read_excel(INPUT_DEFAULT)
    schools = build_school_records(links)
    plans, duplicate_rows, _ = plan_pages(regions, links, schools, config)
    enabled = set(str(value) for value in config["enabled_generation_phases"])
    plans = [
        row for row in plans if generation_phase(row) in enabled
    ]
    before = preflight(plans, normalization_info)
    REPORTS.mkdir(parents=True, exist_ok=True)
    json_write(REPORTS / "production_full_school_correction_preflight.json", before)
    if not before["passed"]:
        raise RuntimeError("학교명 정정 사전 검증 실패")

    output = choose_output()
    generated = generate_preview(
        output,
        plans,
        sorted({str(row.get("지역명", "")) for row in plans if row.get("지역명")}),
        schools,
        links,
        settings,
        config,
    )
    structural = validate_preview(output, generated, schools, settings)
    records = list(generated["records"])
    inventory = inventory_from_records(records)
    duplicate_summary, duplicate_rows_report = duplicate_reports(inventory)
    similarity_summary, similarity_rows = similarity_scan(inventory)

    similarity_fields = (
        "page_a", "page_b", "page_type", "similarity",
        "common_sentence_count", "simple_replacement", "review",
    )
    write_csv(
        REPORTS / "production_full_body_similarity_high.csv",
        similarity_fields,
        similarity_rows["body"],
    )
    write_csv(
        REPORTS / "production_full_template_like_pages.csv",
        similarity_fields,
        similarity_rows["template"],
    )
    write_csv(
        REPORTS / "production_full_duplicate_check.csv",
        ("duplicate_type", "page_a", "page_b", "value_signature"),
        duplicate_rows_report,
    )

    html_files = list(output.rglob("*.html"))
    sitemap_urls: list[str] = []
    for sitemap in output.glob("sitemap-*.xml"):
        sitemap_urls.extend(re.findall(
            r"<loc>(.*?)</loc>",
            sitemap.read_text(encoding="utf-8"),
            re.DOTALL,
        ))
    corrected_records = [
        row for row in records if row.get("학교명") == OFFICIAL_NAME
    ]
    corrected_consistency_errors = sum(
        any(
            OFFICIAL_NAME not in str(row[field])
            for field in (
                "keyword", "slug", "canonical", "title",
                "description", "path",
            )
        )
        for row in corrected_records
    )
    blocker_counts = {
        "body_ge_95": similarity_summary.get("body_ge_95", 0),
        "exact_body": duplicate_summary.get("duplicate_body_count", 0),
        "exact_normalized_body": duplicate_summary.get(
            "duplicate_normalized_body_count", 0
        ),
        "simple_replacement": similarity_summary.get(
            "simple_replacement_pair_count", 0
        ),
        "duplicate_slug": duplicate_summary.get("duplicate_slug_count", 0),
        "duplicate_url": duplicate_summary.get("duplicate_URL_count", 0),
        "duplicate_canonical": duplicate_summary.get(
            "duplicate_canonical_count", 0
        ),
        "structural_errors": len(structural.get("errors", [])),
        "corrected_school_consistency": corrected_consistency_errors,
    }
    passed = not any(blocker_counts.values())
    report = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "candidate_output": str(output),
        "corrected_school_name": OFFICIAL_NAME,
        "correction_preflight": before,
        "content_page_count": len(records),
        "html_count": len(html_files),
        "sitemap_url_count": len(sitemap_urls),
        "robots_exists": (output / "robots.txt").is_file(),
        "generation": {
            key: value for key, value in generated.items() if key != "records"
        },
        "structural_validation": structural,
        "duplicate_summary": duplicate_summary,
        "similarity_summary": similarity_summary,
        "body_90_to_95": (
            similarity_summary.get("body_ge_90", 0)
            - similarity_summary.get("body_ge_95", 0)
        ),
        "body_80_to_90": (
            similarity_summary.get("body_ge_80", 0)
            - similarity_summary.get("body_ge_90", 0)
        ),
        "blocker_counts": blocker_counts,
        "passed": passed,
        "preview_started": False,
        "elapsed_seconds": round(time.perf_counter() - started, 3),
    }
    json_write(REPORTS / "production_full_candidate_validation.json", report)
    print(json.dumps({
        "candidate_output": str(output),
        "content_page_count": len(records),
        "html_count": len(html_files),
        "sitemap_url_count": len(sitemap_urls),
        "passed": passed,
        "blocker_counts": blocker_counts,
        "elapsed_seconds": report["elapsed_seconds"],
    }, ensure_ascii=False, indent=2))
    return 0 if passed else 1


if __name__ == "__main__":
    raise SystemExit(main())
