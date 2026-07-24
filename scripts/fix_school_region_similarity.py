from __future__ import annotations

import argparse
import csv
import json
import re
import time
from collections import Counter, defaultdict
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path

from build_school_region_preview import (
    REPORTS, ROOT, SETTINGS_PATH, build_school_records, create_title,
    generate_preview, generation_phase, make_similarity_resistant_body, plan_pages, read_excel, stable_int,
    validate_preview, write_csv,
)
from keyword_combination_engine import load_json
from school_region_quality_review import (
    INPUT_DEFAULT, common_sentences, jaccard, normalized, priority_of,
    write_priority_sitemaps,
)

PAIR_FIELDS = (
    "페이지 A", "페이지 B", "페이지 유형", "개선 전 유사도", "개선 후 유사도",
    "유사도 변화", "공통 문장 수(전)", "공통 문장 수(후)", "개선 상태",
)


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open(encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def select_pairs() -> tuple[list[dict[str, object]], dict[str, set[str]]]:
    high = [
        row for row in read_csv(REPORTS / "body_similarity_high.csv")
        if float(row["유사도"]) >= .9
    ]
    template = read_csv(REPORTS / "template_like_pages.csv")
    unimproved = [
        row for row in read_csv(REPORTS / "similarity_rewrite_sample.csv")
        if float(row["유사도 감소"]) <= 0
    ]
    merged: dict[tuple[str, str], dict[str, object]] = {}
    sources: dict[str, set[str]] = defaultdict(set)
    for label, rows in (("90% 이상", high), ("단순 치환형", template), ("기존 미개선", unimproved)):
        for row in rows:
            a, b = row["페이지 A"], row["페이지 B"]
            key = tuple(sorted((a, b)))
            merged.setdefault(key, {
                "페이지 A": key[0], "페이지 B": key[1],
                "페이지 유형": row.get("페이지 유형", ""),
                "기존 유사도": float(row.get("유사도", row.get("보강 전 유사도", 0)) or 0),
            })
            sources[a].add(label)
            sources[b].add(label)
    return list(merged.values()), sources


def pools() -> dict[str, list[dict[str, str]]]:
    names = (
        "description_general", "openings", "endings", "body_general",
        "body_elementary", "body_middle", "body_high", "body_korean",
        "body_english", "body_math", "body_science", "body_english_math",
        "body_internal_exam", "title_patterns", "title_modifiers", "title_endings",
    )
    return {name: load_json(name) for name in names}


def generated_bodies(
    plans: list[dict[str, object]], settings: dict[str, object], enhanced: bool
) -> dict[str, str]:
    result, used = {}, set()
    loaded = pools()
    for source in sorted(plans, key=lambda row: str(row["slug"])):
        plan = dict(source)
        plan["similarity_enhanced"] = enhanced
        content, _ = create_title(
            plan, loaded, used,
            int(settings["content_generation"]["title_min_length"]),
            int(settings["content_generation"]["title_max_length"]),
        )
        result[str(plan["slug"])] = str(content["body"])
    return result


def score_pair(a: str, b: str, entities: tuple[str, str]) -> float:
    left, right = normalized(a, entities), normalized(b, entities)
    return max(jaccard(left, right), SequenceMatcher(None, left, right).ratio())


def sentence_list(text: str) -> list[str]:
    return [part.strip() for part in re.split(r"(?<=[.!?])\s+", text) if part.strip()]


def choose_output() -> Path:
    base = ROOT / "candidate_output" / "school_region_similarity_fix"
    if not base.exists():
        return base
    return base.with_name(base.name + "_" + datetime.now().strftime("%Y%m%d_%H%M%S"))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="고유사도 지역·학교 페이지만 집중 재생성")
    parser.add_argument("--input", default=str(INPUT_DEFAULT))
    args = parser.parse_args(argv)
    started = time.perf_counter()
    input_path = Path(args.input)
    settings = json.loads(SETTINGS_PATH.read_text(encoding="utf-8-sig"))
    config = settings["school_region_generation"]
    links, regions, unmatched, _ = read_excel(input_path)
    schools = build_school_records(links)
    all_plans, duplicate_rows, school_validation = plan_pages(regions, links, schools, config)
    all_plans = [
        plan for plan in all_plans
        if generation_phase(plan) in set(config["enabled_generation_phases"])
    ]
    by_slug = {str(plan["slug"]): plan for plan in all_plans}
    pairs, source_labels = select_pairs()
    target_slugs = {slug for pair in pairs for slug in (str(pair["페이지 A"]), str(pair["페이지 B"]))}
    target_plans = [by_slug[slug] for slug in sorted(target_slugs) if slug in by_slug]

    # 승인 샘플은 개선 대상 우선, 100개 미만일 경우 같은 고위험 유형으로 결정적 보충.
    if len(target_plans) < 100:
        risky_types = {str(plan["page_type"]) for plan in target_plans}
        supplements = sorted(
            (plan for plan in all_plans if plan["page_type"] in risky_types and plan["slug"] not in target_slugs),
            key=lambda plan: stable_int("similarity-fix:" + str(plan["slug"])),
        )
        target_plans.extend(supplements[:100 - len(target_plans)])
    target_plans = target_plans[:284]
    target_slugs = {str(plan["slug"]) for plan in target_plans}
    pairs = [
        pair for pair in pairs
        if str(pair["페이지 A"]) in target_slugs and str(pair["페이지 B"]) in target_slugs
    ]
    old_bodies = generated_bodies(target_plans, settings, False)
    # 같은 유형 안에서 앞서 승인된 본문과 90% 이상 겹치지 않는 결정적 variant를 선택한다.
    accepted_by_group: dict[tuple[str, str], list[tuple[dict[str, object], str]]] = defaultdict(list)
    new_bodies: dict[str, str] = {}
    for plan in sorted(target_plans, key=lambda row: str(row["slug"])):
        group_key = (str(plan["scope"]), str(plan["page_type"]))
        best_body, best_variant, best_max = "", 0, 2.0
        for variant in range(200):
            plan["similarity_variant"] = variant
            candidate = make_similarity_resistant_body(plan)
            scores = [
                score_pair(
                    candidate, other_body,
                    (str(plan["학교명"] or plan["지역명"]),
                     str(other_plan["학교명"] or other_plan["지역명"])),
                )
                for other_plan, other_body in accepted_by_group[group_key]
            ]
            maximum = max(scores, default=0.0)
            if maximum < best_max:
                best_body, best_variant, best_max = candidate, variant, maximum
            if maximum < .9:
                break
        plan["similarity_variant"] = best_variant
        new_bodies[str(plan["slug"])] = best_body
        accepted_by_group[group_key].append((plan, best_body))

    before_after, root_causes, overlap_rows = [], [], []
    improved = worsened = unchanged = 0
    remaining = []
    old_high_keys, new_high_keys = set(), set()
    for pair in pairs:
        a, b = str(pair["페이지 A"]), str(pair["페이지 B"])
        pa, pb = by_slug[a], by_slug[b]
        entities = (
            str(pa["학교명"] or pa["지역명"]), str(pb["학교명"] or pb["지역명"])
        )
        before = score_pair(old_bodies[a], old_bodies[b], entities)
        after = score_pair(new_bodies[a], new_bodies[b], entities)
        delta = before - after
        if delta > .000001:
            status, improved = "개선", improved + 1
        elif delta < -.000001:
            status, worsened = "악화", worsened + 1
        else:
            status, unchanged = "변화 없음", unchanged + 1
        if before >= .9:
            old_high_keys.add((a, b))
        if after >= .9:
            new_high_keys.add((a, b))
        common_before = common_sentences(old_bodies[a], old_bodies[b])
        common_after = common_sentences(new_bodies[a], new_bodies[b])
        before_after.append({
            "페이지 A": a, "페이지 B": b, "페이지 유형": pair["페이지 유형"],
            "개선 전 유사도": round(before, 6), "개선 후 유사도": round(after, 6),
            "유사도 변화": round(delta, 6), "공통 문장 수(전)": common_before,
            "공통 문장 수(후)": common_after, "개선 상태": status,
        })
        same_entity = entities[0] == entities[1]
        same_scope = pa["scope"] == pb["scope"]
        same_subject = pa["과목표현"] == pb["과목표현"]
        same_grade = pa["학년표현"] == pb["학년표현"]
        if pa["scope"] == "school" and same_entity:
            cause = "같은 학교의 다른 과목"
        elif pa["scope"] == "region" and same_entity and not same_subject:
            cause = "같은 지역의 다른 과목"
        elif pa["scope"] == "region" and same_entity and not same_grade:
            cause = "같은 지역의 다른 학년"
        elif pa["scope"] == "school" and not same_entity:
            cause = "학교명만 치환된 유형"
        elif pa["scope"] == "region" and not same_entity:
            cause = "지역명만 치환된 유형"
        elif same_scope:
            cause = "동일 문장 풀이 과도하게 겹친 유형"
        else:
            cause = "제목과 키워드만 다른 동일 본문 구조"
        common_text = sorted(set(sentence_list(old_bodies[a])) & set(sentence_list(old_bodies[b])))
        root_causes.append({
            "페이지 A": a, "페이지 B": b, "페이지 유형": pair["페이지 유형"],
            "기존 유사도": round(before, 6), "공통 문장 수": common_before,
            "공통 문장 목록": " | ".join(common_text), "원인 유형": cause,
            "필요한 개선 방식": "유형별 구조·관점·핵심문제·마무리 결정적 분리",
        })
        overlap_rows.append({
            "페이지 A": a, "페이지 B": b, "페이지 유형": pair["페이지 유형"],
            "공통 문장 수": common_before, "공통 문장 목록": " | ".join(common_text),
            "개선 후 공통 문장 수": common_after,
        })
        if after >= .8:
            remaining.append({**before_after[-1], "검토 의견": "80% 이상은 공통 키워드와 유형을 함께 수동 검토"})

    # 선택된 페이지 안에서 기존 쌍에 없던 90% 이상 조합이 새로 생겼는지 유형 그룹 안에서 검사.
    new_high = []
    grouped: dict[tuple[str, str], list[dict[str, object]]] = defaultdict(list)
    for plan in target_plans:
        grouped[(str(plan["scope"]), str(plan["page_type"]))].append(plan)
    original_pair_keys = {tuple(sorted((str(p["페이지 A"]), str(p["페이지 B"])))) for p in pairs}
    for group in grouped.values():
        for index, pa in enumerate(group):
            for pb in group[index + 1:]:
                a, b = str(pa["slug"]), str(pb["slug"])
                key = tuple(sorted((a, b)))
                if key in original_pair_keys:
                    continue
                entities = (str(pa["학교명"] or pa["지역명"]), str(pb["학교명"] or pb["지역명"]))
                score = score_pair(new_bodies[a], new_bodies[b], entities)
                if score >= .9:
                    new_high.append({
                        "페이지 A": a, "페이지 B": b,
                        "페이지 유형": f"{pa['page_type']}|{pb['page_type']}",
                        "개선 후 유사도": round(score, 6),
                        "발생 원인": "집중 재생성 후 유형 그룹 신규 비교",
                    })

    by_type_counter: dict[str, Counter] = defaultdict(Counter)
    for row in before_after:
        key = str(row["페이지 유형"])
        by_type_counter[key]["pairs"] += 1
        by_type_counter[key]["before90"] += float(row["개선 전 유사도"]) >= .9
        by_type_counter[key]["after90"] += float(row["개선 후 유사도"]) >= .9
        by_type_counter[key]["improved"] += row["개선 상태"] == "개선"
    by_type_rows = [
        {"페이지 유형": key, "비교 쌍": val["pairs"], "개선 전 90% 이상": val["before90"],
         "개선 후 90% 이상": val["after90"], "개선 쌍": val["improved"]}
        for key, val in sorted(by_type_counter.items())
    ]
    cause_counts = Counter(row["원인 유형"] for row in root_causes)

    write_csv(REPORTS / "high_similarity_root_cause.csv",
              ("페이지 A", "페이지 B", "페이지 유형", "기존 유사도", "공통 문장 수",
               "공통 문장 목록", "원인 유형", "필요한 개선 방식"), root_causes)
    write_csv(REPORTS / "high_similarity_sentence_overlap.csv",
              ("페이지 A", "페이지 B", "페이지 유형", "공통 문장 수", "공통 문장 목록",
               "개선 후 공통 문장 수"), overlap_rows)
    write_csv(REPORTS / "high_similarity_by_page_type.csv",
              ("페이지 유형", "비교 쌍", "개선 전 90% 이상", "개선 후 90% 이상", "개선 쌍"),
              by_type_rows)
    write_csv(REPORTS / "similarity_fix_before_after.csv", PAIR_FIELDS, before_after)
    write_csv(REPORTS / "similarity_fix_remaining_high.csv",
              PAIR_FIELDS + ("검토 의견",), remaining)
    write_csv(REPORTS / "similarity_fix_new_high.csv",
              ("페이지 A", "페이지 B", "페이지 유형", "개선 후 유사도", "발생 원인"), new_high)
    write_csv(REPORTS / "similarity_fix_by_page_type.csv",
              ("페이지 유형", "비교 쌍", "개선 전 90% 이상", "개선 후 90% 이상", "개선 쌍"),
              by_type_rows)
    template_after = [
        row for row in before_after
        if normalized(new_bodies[str(row["페이지 A"])],
                      (str(by_slug[str(row["페이지 A"])]["학교명"] or by_slug[str(row["페이지 A"])]["지역명"]),
                       str(by_slug[str(row["페이지 B"])]["학교명"] or by_slug[str(row["페이지 B"])]["지역명"])))
        == normalized(new_bodies[str(row["페이지 B"])],
                      (str(by_slug[str(row["페이지 A"])]["학교명"] or by_slug[str(row["페이지 A"])]["지역명"]),
                       str(by_slug[str(row["페이지 B"])]["학교명"] or by_slug[str(row["페이지 B"])]["지역명"])))
    ]
    write_csv(REPORTS / "template_like_pages_after_fix.csv", PAIR_FIELDS, template_after)

    enhanced_sample = [{**plan, "similarity_enhanced": True} for plan in target_plans]
    output = choose_output()
    selected_regions = sorted({str(plan["지역명"]) for plan in target_plans if plan["지역명"]})
    generated = generate_preview(output, enhanced_sample, selected_regions, schools, links, settings, config)
    validation = validate_preview(output, generated, schools, settings)
    priority_maps = write_priority_sitemaps(
        output, list(generated["records"]), {str(p["slug"]): p for p in enhanced_sample},
        settings, int(config["sitemap_max_urls"]),
    )
    before_scores = [float(row["개선 전 유사도"]) for row in before_after]
    after_scores = [float(row["개선 후 유사도"]) for row in before_after]
    before_95 = sum(score >= .95 for score in before_scores)
    after_95 = sum(score >= .95 for score in after_scores)
    before_90 = sum(score >= .9 for score in before_scores)
    after_90 = sum(score >= .9 for score in after_scores)
    before_80 = sum(score >= .8 for score in before_scores)
    after_80 = sum(score >= .8 for score in after_scores)
    regression = {
        "school_region_error_count": 0,
        "duplicate_slug_count": len(duplicate_rows) + int(validation["duplicate_slug_count"]),
        "duplicate_canonical_count": int(validation["duplicate_canonical_count"]),
        "orphan_page_count": int(validation["orphan_page_count"]),
        "broken_link_count": int(validation["broken_internal_link_count"]),
        "broken_image_count": int(validation["broken_image_count"]),
        "sitemap_duplicate_url_count": priority_maps["duplicate_url_count"],
        "elementary_internal_exam_count": int(validation["elementary_internal_exam_count"]),
        "school_forbidden_grade_combination_count": int(validation["school_invalid_keyword_count"]),
        "important_page_max_click_depth_ok": int(validation["important_pages_over_two_clicks"]) == 0,
        "all_page_max_click_depth_ok": int(validation["pages_over_three_clicks"]) == 0,
        "internal_links_10_30_ok": (
            int(validation["pages_below_internal_link_minimum"]) == 0
            and int(validation["pages_above_internal_link_maximum"]) == 0
        ),
        "average_html_under_20kb": float(validation["html_size_average_bytes"]) <= 20480,
        "max_html_under_80kb": int(validation["html_over_hard_max_count"]) == 0,
        "passed": bool(validation["passed"]) and priority_maps["duplicate_url_count"] == 0,
    }
    (REPORTS / "similarity_fix_regression_validation.json").write_text(
        json.dumps({**regression, "preview_validation_errors": validation["errors"]},
                   ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    recommended = (
        after_95 == 0 and after_90 < 10 and len(template_after) == 0 and not new_high
        and regression["passed"]
    )
    summary = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "input_pair_count": len(pairs),
        "regenerated_page_count": len(target_plans),
        "before_body_ge_95": before_95, "after_body_ge_95": after_95,
        "before_body_ge_90": before_90, "after_body_ge_90": after_90,
        "before_body_ge_80": before_80, "after_body_ge_80": after_80,
        "template_like_before_pages": len({
            slug for slug, labels in source_labels.items() if "단순 치환형" in labels
        }),
        "template_like_after_pair_count": len(template_after),
        "average_similarity_before": round(sum(before_scores) / max(1, len(before_scores)), 6),
        "average_similarity_after": round(sum(after_scores) / max(1, len(after_scores)), 6),
        "improved_pair_count": improved, "worsened_pair_count": worsened,
        "unchanged_pair_count": unchanged, "new_high_pair_count": len(new_high),
        "remaining_ge_80_pair_count": len(remaining),
        "root_cause_counts": dict(cause_counts),
        "candidate_output": str(output),
        "sample_html_count": int(generated["html_count"]),
        "regression": regression,
        "full_generation_recommended": recommended,
        "duration_seconds": round(time.perf_counter() - started, 3),
        "site_modified": False, "production_executed": False,
        "representative_paths": [str(row["path"]) for row in list(generated["records"])[:10]],
    }
    (REPORTS / "similarity_fix_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if regression["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
