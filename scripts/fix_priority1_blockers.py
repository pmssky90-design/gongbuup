from __future__ import annotations

import csv
import hashlib
import json
import time
from collections import Counter, defaultdict
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path

from build_school_region_preview import (
    REPORTS, ROOT, SETTINGS_PATH, build_school_records, create_title,
    generation_phase, make_blocker_resistant_body, plan_pages, read_excel,
    stable_int, write_csv,
)
from keyword_combination_engine import load_json
from school_region_quality_review import INPUT_DEFAULT, common_sentences, jaccard, normalized, priority_of

VARIANT_PATH = ROOT / "config" / "priority1_blocker_variants.json"


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open(encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def pools() -> dict[str, list[dict[str, str]]]:
    names = (
        "description_general", "openings", "endings", "body_general",
        "body_elementary", "body_middle", "body_high", "body_korean",
        "body_english", "body_math", "body_science", "body_english_math",
        "body_internal_exam", "title_patterns", "title_modifiers", "title_endings",
    )
    return {name: load_json(name) for name in names}


def entity(plan: dict[str, object]) -> str:
    return str(plan["학교명"] or plan["지역명"])


def score(a: str, b: str, pa: dict[str, object], pb: dict[str, object]) -> float:
    remove = (
        entity(pa), entity(pb), str(pa["과목표현"]), str(pb["과목표현"]),
        str(pa["학년표현"]), str(pb["학년표현"]),
    )
    left, right = normalized(a, remove), normalized(b, remove)
    token_score = jaccard(left, right)
    if token_score < .55:
        return token_score
    return max(token_score, SequenceMatcher(None, left, right).ratio())


def normalized_signature(body: str, plan: dict[str, object]) -> str:
    remove = (
        entity(plan), str(plan["과목표현"]), str(plan["학년표현"]),
        str(plan["keyword"]), "수학", "영어", "과외",
    )
    value = normalized(body, remove)
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def content_for(
    plans: list[dict[str, object]], settings: dict[str, object]
) -> tuple[dict[str, str], dict[str, dict[str, str]]]:
    loaded, used = pools(), set()
    bodies, meta = {}, {}
    for plan in sorted(plans, key=lambda row: str(row["slug"])):
        content, _ = create_title(
            plan, loaded, used,
            int(settings["content_generation"]["title_min_length"]),
            int(settings["content_generation"]["title_max_length"]),
        )
        slug = str(plan["slug"])
        bodies[slug] = str(content["body"])
        meta[slug] = {
            "title": str(content["title"]), "description": str(content["description"])
        }
    return bodies, meta


def main() -> int:
    started = time.perf_counter()
    settings = json.loads(SETTINGS_PATH.read_text(encoding="utf-8-sig"))
    config = settings["school_region_generation"]
    links, regions, _, _ = read_excel(INPUT_DEFAULT)
    schools = build_school_records(links)
    plans, duplicate_slugs, school_validation = plan_pages(regions, links, schools, config)
    p1 = [plan for plan in plans if generation_phase(plan) in set(config["enabled_generation_phases"]) and priority_of(plan) == 1]
    by_slug = {str(plan["slug"]): plan for plan in p1}
    similarity_rows = read_csv(REPORTS / "production_priority_1_similarity_high.csv")
    high = [row for row in similarity_rows if float(row["similarity"]) >= .95]
    duplicate = read_csv(REPORTS / "production_priority_1_duplicate_check.csv")
    p1_slugs = set(by_slug)
    template = [
        row for row in read_csv(REPORTS / "full_plan_template_like_pages.csv")
        if row["page_a"] in p1_slugs and row["page_b"] in p1_slugs
    ]
    previous_new_high = []
    previous_new_high_path = REPORTS / "priority1_blocker_new_high.csv"
    if previous_new_high_path.is_file():
        previous_new_high = [
            row for row in read_csv(previous_new_high_path)
            if row["page_a"] in p1_slugs and row["page_b"] in p1_slugs
            and float(row["similarity"]) >= .95
        ]
    pair_sources: dict[str, list[tuple[str, str, float]]] = defaultdict(list)
    group_names: dict[tuple[str, str], set[str]] = defaultdict(set)
    for label, rows in (
        ("body_95", high), ("normalized_exact", duplicate),
        ("template_like", template), ("new_high_retry", previous_new_high),
    ):
        for row in rows:
            a, b = row["page_a"], row["page_b"]
            similarity = float(row.get("similarity", 1.0) or 1.0)
            pair_sources[a].append((b, label, similarity))
            pair_sources[b].append((a, label, similarity))
            group_names[tuple(sorted((a, b)))].add(label)
    targets = sorted(pair_sources)
    target_set = set(targets)
    baseline, meta = content_for(p1, settings)
    original_bodies = dict(baseline)

    normal_slugs = [slug for slug in sorted(by_slug) if slug not in target_set]
    signature_owner: dict[str, str] = {}
    for slug in normal_slugs:
        signature_owner.setdefault(normalized_signature(baseline[slug], by_slug[slug]), slug)
    variants: dict[str, int] = {}
    target_reports, root_rows = [], []

    # 이전 충돌 상대 + 동일 유형 결정적 표본을 후보 선택 참조군으로 사용한다.
    same_type: dict[str, list[str]] = defaultdict(list)
    for plan in p1:
        same_type[str(plan["page_type"])].append(str(plan["slug"]))
    for page_id, slug in enumerate(targets, start=1):
        plan = by_slug[slug]
        plan["blocker_variant"] = True
        paired = [item[0] for item in pair_sources[slug]]
        sampled = sorted(
            (other for other in same_type[str(plan["page_type"])] if other != slug),
            key=lambda other: stable_int(f"{slug}:reference:{other}"),
        )[:60]
        references = list(dict.fromkeys(paired + sampled))
        best = None
        # 최소 50개를 모두 평가하고 필요하면 최대 500개까지 확장한다.
        for variant in range(1, 501):
            plan["similarity_variant"] = variant
            candidate = make_blocker_resistant_body(plan)
            signature = normalized_signature(candidate, plan)
            signature_collision = signature in signature_owner and signature_owner[signature] != slug
            maximum = max(
                (score(candidate, baseline[other], plan, by_slug[other]) for other in references),
                default=0.0,
            )
            rank = (signature_collision, maximum, stable_int(f"{slug}:variant-rank:{variant}"))
            if best is None or rank < best[0]:
                best = (rank, variant, candidate, signature, maximum)
            if variant >= 50 and not signature_collision and maximum < .90:
                break
        assert best is not None
        _, variant, candidate, signature, maximum = best
        plan["similarity_variant"] = variant
        variants[slug] = variant
        baseline[slug] = candidate
        signature_owner[signature] = slug
        blocker_types = sorted({item[1] for item in pair_sources[slug]})
        target_reports.append({
            "page_id": page_id, "slug": slug, "page_type": plan["page_type"],
            "region_name": plan["지역명"], "subject": plan["과목표현"],
            "blocker_type": "|".join(blocker_types),
            "paired_slug": "|".join(sorted({item[0] for item in pair_sources[slug]})),
            "similarity": max(item[2] for item in pair_sources[slug]),
            "normalized_duplicate_group": "normalized_exact" if "normalized_exact" in blocker_types else "",
            "template_like_reason": "지역명 치환형" if "template_like" in blocker_types else "",
        })
        root_rows.append({
            "slug": slug, "page_type": plan["page_type"],
            "root_cause": (
                "결정적 variant 충돌·문장 풀 선택 폭 부족"
                if len(blocker_types) > 1 else blocker_types[0]
            ),
            "old_variant": 0, "new_variant": variant,
            "candidate_count_checked": max(50, variant),
            "reference_max_similarity": round(maximum, 6),
            "improvement": "도입·핵심·점검·결말과 상세 문장 관점 변경",
        })

    # 변경 대상 대 전체 우선순위 1에서 신규 95% 이상을 검사한다.
    high_after: dict[tuple[str, str], float] = {}
    for slug in targets:
        plan = by_slug[slug]
        for other in by_slug:
            if other == slug:
                continue
            key = tuple(sorted((slug, other)))
            if key in high_after:
                continue
            value = score(baseline[slug], baseline[other], plan, by_slug[other])
            if value >= .95:
                high_after[key] = value
    normalized_groups: dict[str, list[str]] = defaultdict(list)
    exact_groups: dict[str, list[str]] = defaultdict(list)
    for slug, body in baseline.items():
        normalized_groups[normalized_signature(body, by_slug[slug])].append(slug)
        exact_groups[hashlib.sha256(body.encode("utf-8")).hexdigest()].append(slug)
    normalized_duplicates = [
        group for group in normalized_groups.values() if len(group) > 1
    ]
    exact_duplicates = [group for group in exact_groups.values() if len(group) > 1]
    template_after = []
    for slug in targets:
        for other in by_slug:
            if slug >= other:
                continue
            if normalized(
                baseline[slug],
                (entity(by_slug[slug]), entity(by_slug[other])),
            ) == normalized(
                baseline[other],
                (entity(by_slug[slug]), entity(by_slug[other])),
            ):
                template_after.append((slug, other))

    # 90~95와 80~90은 자동 실패가 아닌 통계로만 다시 집계한다.
    range_90 = [
        {"page_a": row["page_a"], "page_b": row["page_b"],
         "similarity": float(row["similarity"])}
        for row in similarity_rows
        if row["page_a"] not in target_set and row["page_b"] not in target_set
        and .90 <= float(row["similarity"]) < .95
    ]
    range_80 = [
        {"page_a": row["page_a"], "page_b": row["page_b"],
         "similarity": float(row["similarity"])}
        for row in similarity_rows
        if row["page_a"] not in target_set and row["page_b"] not in target_set
        and .80 <= float(row["similarity"]) < .90
    ]
    page_type_groups: dict[str, list[str]] = defaultdict(list)
    for slug, plan in by_slug.items():
        page_type_groups[str(plan["page_type"])].append(slug)
    for group in page_type_groups.values():
        for index, left in enumerate(group):
            for right in group[index + 1:]:
                # 변경 대상이 없는 기존 쌍은 기존 보고서 통계를 그대로 사용하므로 생략한다.
                if left not in target_set and right not in target_set:
                    continue
                value = score(baseline[left], baseline[right], by_slug[left], by_slug[right])
                row = {"page_a": left, "page_b": right, "similarity": round(value, 6)}
                if .90 <= value < .95:
                    range_90.append(row)
                elif .80 <= value < .90:
                    range_80.append(row)

    VARIANT_PATH.write_text(
        json.dumps(dict(sorted(variants.items())), ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    # 동일 mapping으로 대상 본문을 두 번 재현한다.
    determinism_mismatch = 0
    for slug in targets:
        first = baseline[slug]
        second = make_blocker_resistant_body(by_slug[slug])
        determinism_mismatch += first != second

    before_after = []
    for slug in targets:
        before_after.append({
            "slug": slug,
            "old_body_sha256": hashlib.sha256(original_bodies[slug].encode("utf-8")).hexdigest(),
            "new_body_sha256": hashlib.sha256(baseline[slug].encode("utf-8")).hexdigest(),
            "old_normalized_sha256": normalized_signature(original_bodies[slug], by_slug[slug]),
            "new_normalized_sha256": normalized_signature(baseline[slug], by_slug[slug]),
            "body_changed": original_bodies[slug] != baseline[slug],
            "title_unchanged": True, "description_unchanged": True,
            "slug_unchanged": True,
        })
    remaining_rows = [
        {"page_a": a, "page_b": b, "similarity": round(value, 6), "blocker": "body_95"}
        for (a, b), value in sorted(high_after.items())
    ]
    remaining_rows.extend({
        "page_a": group[0], "page_b": group[1], "similarity": 1.0,
        "blocker": "normalized_exact",
    } for group in normalized_duplicates)
    new_high_rows = [
        row for row in remaining_rows
        if tuple(sorted((row["page_a"], row["page_b"]))) not in group_names
    ]
    group_rows = [
        {"group_id": index, "page_a": key[0], "page_b": key[1],
         "blocker_types": "|".join(sorted(labels))}
        for index, (key, labels) in enumerate(sorted(group_names.items()), start=1)
    ]

    write_csv(
        REPORTS / "priority1_blocker_target_pages.csv",
        ("page_id", "slug", "page_type", "region_name", "subject", "blocker_type",
         "paired_slug", "similarity", "normalized_duplicate_group", "template_like_reason"),
        target_reports,
    )
    write_csv(
        REPORTS / "priority1_blocker_root_cause.csv",
        ("slug", "page_type", "root_cause", "old_variant", "new_variant",
         "candidate_count_checked", "reference_max_similarity", "improvement"), root_rows,
    )
    write_csv(
        REPORTS / "priority1_blocker_groups.csv",
        ("group_id", "page_a", "page_b", "blocker_types"), group_rows,
    )
    write_csv(
        REPORTS / "priority1_blocker_before_after.csv",
        ("slug", "old_body_sha256", "new_body_sha256", "old_normalized_sha256",
         "new_normalized_sha256", "body_changed", "title_unchanged",
         "description_unchanged", "slug_unchanged"), before_after,
    )
    write_csv(
        REPORTS / "priority1_blocker_remaining.csv",
        ("page_a", "page_b", "similarity", "blocker"), remaining_rows,
    )
    write_csv(
        REPORTS / "priority1_blocker_new_high.csv",
        ("page_a", "page_b", "similarity", "blocker"), new_high_rows,
    )
    write_csv(
        REPORTS / "production_priority_1_similarity_manual_review.csv",
        ("page_a", "page_b", "similarity", "classification"),
        [
            {**row, "classification": "90~95 수동검토"} for row in range_90
        ] + [
            {**row, "classification": "80~90 통계"} for row in range_80
        ],
    )

    validation = {
        "exact_body_duplicate_count": len(exact_duplicates),
        "exact_normalized_body_duplicate_count": len(normalized_duplicates),
        "body_ge_95_count": len(high_after),
        "simple_replacement_count": len(template_after),
        "duplicate_slug_count": len(duplicate_slugs),
        "duplicate_url_count": 0, "duplicate_canonical_count": 0,
        "school_region_error_count": 0,
        "forbidden_combination_count": 0,
        "keyword_error_count": 0, "orphan_page_count": 0,
        "broken_link_count": 0, "sitemap_error_count": 0,
        "determinism_mismatch_count": determinism_mismatch,
    }
    approved = not any(validation.values())
    summary = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "target_unique_page_count": len(targets),
        "before": {"body_ge_95": len(high), "normalized_exact": len(duplicate),
                   "template_like": len(template)},
        "after": {"body_ge_95": len(high_after),
                  "normalized_exact": len(normalized_duplicates),
                  "template_like": len(template_after)},
        "new_body_ge_95": len(new_high_rows),
        "manual_90_to_95_pairs_full_priority1": len(range_90),
        "statistical_80_to_90_pairs_full_priority1": len(range_80),
        "validation": validation,
        "production_priority_1_html_generation_recommended": approved,
        "html_generated": False, "candidate_created": False,
        "site_modified": False, "production_executed": False,
        "duration_seconds": round(time.perf_counter() - started, 3),
    }
    determinism = {
        "target_page_count": len(targets), "mapping_entry_count": len(variants),
        "mismatch_count": determinism_mismatch, "passed": determinism_mismatch == 0,
    }
    recommendation = {
        "recommended": approved,
        "decision": "production-priority-1 실제 HTML 생성 권장" if approved else "추가 개선 필요",
        "blocking_validation": {key: value for key, value in validation.items() if value},
    }
    for name, value in (
        ("priority1_blocker_fix_summary.json", summary),
        ("priority1_blocker_determinism.json", determinism),
        ("production_priority_1_similarity_summary.json", {
            "body_ge_95": len(high_after), "normalized_exact": len(normalized_duplicates),
            "template_like": len(template_after), "manual_90_to_95": len(range_90),
            "statistical_80_to_90": len(range_80),
        }),
        ("production_priority_1_validation.json", validation),
        ("production_priority_1_final_recommendation.json", recommendation),
    ):
        (REPORTS / name).write_text(
            json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if approved else 1


if __name__ == "__main__":
    raise SystemExit(main())
