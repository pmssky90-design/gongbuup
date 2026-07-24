from __future__ import annotations

import csv
import html
import json
import re
import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from build_school_region_preview import (  # noqa: E402
    SCHOOL_ALLOWED_SUFFIXES,
    SCHOOL_FORBIDDEN_AFTER_NAME,
    SETTINGS_PATH,
    build_school_records,
    generation_phase,
    make_blocker_resistant_body,
    plan_pages,
    read_excel,
)
from school_region_quality_review import INPUT_DEFAULT  # noqa: E402


SOURCE = ROOT / "preview" / "production-full-candidate_20260724_151919"
REPORT_DIR = ROOT / "reports" / "production_full_repair"
EXECUTION = REPORT_DIR / "repair_execution_summary.json"
TARGETS = REPORT_DIR / "repair_content_targets.csv"
PATTERNS = {
    "title": re.compile(r"<title>(.*?)</title>", re.I | re.S),
    "description": re.compile(
        r'<meta\s+name=["\']description["\']\s+content=["\'](.*?)["\']',
        re.I | re.S,
    ),
    "canonical": re.compile(
        r'<link\s+rel=["\']canonical["\']\s+href=["\'](.*?)["\']',
        re.I | re.S,
    ),
    "h1": re.compile(r"<h1\b[^>]*>(.*?)</h1>", re.I | re.S),
    "body": re.compile(
        r'<div class="content">.*?<br>\s*(.*?)</div>', re.I | re.S
    ),
}


def extract(path: Path) -> dict[str, str]:
    text = path.read_text(encoding="utf-8")
    result = {"html": text}
    for name, pattern in PATTERNS.items():
        match = pattern.search(text)
        result[name] = html.unescape(match.group(1)).strip() if match else ""
    result["body"] = " ".join(re.sub(r"<[^>]+>", " ", result["body"]).split())
    return result


def main() -> int:
    execution = json.loads(EXECUTION.read_text(encoding="utf-8"))
    repaired = Path(execution["repaired_candidate"])
    with TARGETS.open("r", encoding="utf-8-sig", newline="") as source:
        target_rows = list(csv.DictReader(source))
    settings = json.loads(SETTINGS_PATH.read_text(encoding="utf-8-sig"))
    config = settings["school_region_generation"]
    links, regions, _, _ = read_excel(INPUT_DEFAULT)
    schools = build_school_records(links)
    plans, duplicate_rows, school_validation = plan_pages(
        regions, links, schools, config
    )
    enabled = set(str(value) for value in config["enabled_generation_phases"])
    plans = [row for row in plans if generation_phase(row) in enabled]
    by_slug = {str(row["slug"]): row for row in plans}

    def inspect(row: dict[str, str]) -> dict[str, int]:
        slug = row["slug"]
        original = extract(SOURCE / "과외" / slug / "index.html")
        updated = extract(repaired / "과외" / slug / "index.html")
        plan = dict(by_slug[slug])
        plan["similarity_variant"] = int(row["proposed_variant_id"])
        expected_body = make_blocker_resistant_body(plan)
        return {
            "immutable_error": int(any(
                original[field] != updated[field]
                for field in ("title", "description", "canonical", "h1")
            )),
            "body_determinism_error": int(updated["body"] != expected_body),
            "keyword_error": int(str(plan["keyword"]) not in updated["html"]),
        }

    with ThreadPoolExecutor(max_workers=8) as pool:
        target_checks = list(pool.map(inspect, target_rows, chunksize=64))

    school_region_errors = 0
    forbidden_errors = 0
    for plan in plans:
        slug = str(plan["slug"])
        if plan["scope"] == "school":
            school = str(plan["학교명"])
            keyword = str(plan["keyword"])
            suffix = keyword.removeprefix(school)
            school_region_errors += int(
                school not in schools or suffix not in SCHOOL_ALLOWED_SUFFIXES
            )
            forbidden_errors += int(any(
                school + token in keyword
                for token in SCHOOL_FORBIDDEN_AFTER_NAME
            ))
        forbidden_errors += int(
            str(plan["학년표현"]).startswith("초등")
            and bool(plan["내신사용"])
        )
    result = {
        "candidate": str(repaired),
        "target_page_count": len(target_rows),
        "title_description_canonical_h1_change_error_count": sum(
            row["immutable_error"] for row in target_checks
        ),
        "body_determinism_mismatch_count": sum(
            row["body_determinism_error"] for row in target_checks
        ),
        "required_keyword_error_count": sum(
            row["keyword_error"] for row in target_checks
        ),
        "school_region_error_count": school_region_errors,
        "forbidden_combination_count": forbidden_errors,
        "duplicate_plan_rows_removed": len(duplicate_rows),
        "school_source_validation_error_count": sum(
            not bool(row.get("허용조합", True))
            or bool(row.get("학교급중복표현", False))
            for row in school_validation
        ),
    }
    blocker_keys = (
        "title_description_canonical_h1_change_error_count",
        "body_determinism_mismatch_count",
        "required_keyword_error_count",
        "school_region_error_count",
        "forbidden_combination_count",
        "duplicate_plan_rows_removed",
        "school_source_validation_error_count",
    )
    result["passed"] = all(result[key] == 0 for key in blocker_keys)
    (REPORT_DIR / "repair_invariant_validation.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
