from __future__ import annotations

import argparse
import csv
import json
import re
import shutil
import time
import unicodedata
from collections import Counter, defaultdict
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path
from urllib.parse import quote
from xml.sax.saxutils import escape as xml_escape

from build_school_region_preview import (
    METRO_NAMES, REPORTS, ROOT, SETTINGS_PATH, build_school_records,
    create_title, generate_preview, generation_phase, plan_pages, read_excel,
    select_sample, site_join, stable_int, validate_preview, write_csv,
)
from keyword_combination_engine import load_json

INPUT_DEFAULT = Path(r"C:\gptwp\자료\블로그 할때 주요지역과 그 지역 중학교 고등학교.xlsx")
SIM_FIELDS = (
    "페이지 A", "페이지 B", "페이지 유형", "유사도", "공통 문장 수",
    "단순 지역명·학교명 치환 여부", "재작성 필요 여부",
)
PRIORITY_LABELS = {1: "필수", 2: "중간", 3: "학교 핵심", 4: "보조", 5: "낮음"}


def normalized(text: str, entities: tuple[str, ...] = ()) -> str:
    value = unicodedata.normalize("NFC", text).lower()
    for entity in sorted((x for x in entities if x), key=len, reverse=True):
        value = value.replace(entity.lower(), "{entity}")
    value = re.sub(r"[\W_]+", " ", value, flags=re.UNICODE)
    return " ".join(value.split())


def tokens(text: str) -> set[str]:
    words = text.split()
    if len(words) < 2:
        return set(words)
    return {" ".join(words[i:i + 2]) for i in range(len(words) - 1)}


def jaccard(a: str, b: str) -> float:
    left, right = tokens(a), tokens(b)
    if not left and not right:
        return 1.0
    return len(left & right) / max(1, len(left | right))


def common_sentences(a: str, b: str) -> int:
    split = lambda value: {
        normalized(part) for part in re.split(r"(?<=[.!?다요])\s+", value) if part.strip()
    }
    return len(split(a) & split(b))


def priority_of(plan: dict[str, object]) -> int:
    if plan["scope"] == "directory":
        return 1
    grade = str(plan["학년표현"])
    subject = str(plan["과목표현"])
    exam = bool(plan["내신사용"])
    if plan["scope"] == "school":
        if plan["page_type"] == "school_base" or subject in ("수학", "영어"):
            return 3
        return 4
    if exam:
        return 5
    if plan["page_type"] == "region_base" or (not grade and subject in ("수학", "영어")):
        return 1
    if grade in ("초등학생", "중학생", "고등학생") and subject in ("", "수학", "영어"):
        return 2
    if subject in ("국어", "과학", "영수"):
        return 4
    return 5


def choose_output() -> Path:
    base = ROOT / "candidate_output" / "school_region_quality_review"
    if not base.exists():
        return base
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return base.with_name(f"{base.name}_{stamp}")


def load_pools() -> dict[str, list[dict[str, str]]]:
    names = (
        "description_general", "openings", "endings", "body_general",
        "body_elementary", "body_middle", "body_high", "body_korean",
        "body_english", "body_math", "body_science", "body_english_math",
        "body_internal_exam", "title_patterns", "title_modifiers", "title_endings",
    )
    return {name: load_json(name) for name in names}


def content_inventory(plans: list[dict[str, object]], settings: dict[str, object]) -> list[dict[str, object]]:
    pools = load_pools()
    used: set[str] = set()
    title_min = int(settings["content_generation"]["title_min_length"])
    title_max = int(settings["content_generation"]["title_max_length"])
    rows = []
    for plan in plans:
        content, _ = create_title(plan, pools, used, title_min, title_max)
        entity = str(plan["학교명"] or plan["지역명"])
        rows.append({
            **plan,
            "title": str(content["title"]),
            "description": str(content["description"]),
            "body": str(content["body"]),
            "entity": entity,
        })
    return rows


def similarity_reports(rows: list[dict[str, object]], config: dict[str, object]) -> dict[str, object]:
    # 비교 폭발을 막기 위해 의미 그룹 안에서 정규화 서명 버킷과 결정적 인접 표본을 비교한다.
    buckets: dict[tuple[str, str, str], list[int]] = defaultdict(list)
    for index, row in enumerate(rows):
        group = (
            str(row["scope"]),
            str(row["page_type"]),
            str(row["과목표현"] or row["학년표현"] or "base"),
        )
        body_core = normalized(str(row["body"]), (str(row["entity"]),))
        signature = " ".join(body_core.split()[:12])
        buckets[group + (signature,)].append(index)

    candidate_pairs: set[tuple[int, int]] = set()
    max_pairs = int(config["similarity_max_pairs_per_bucket"])
    for indexes in buckets.values():
        ordered = sorted(indexes, key=lambda i: stable_int(str(rows[i]["slug"])))
        for pos, left in enumerate(ordered):
            for right in ordered[pos + 1:pos + 5]:
                candidate_pairs.add((min(left, right), max(left, right)))
                if len(candidate_pairs) >= max_pairs * max(1, len(buckets)):
                    break
    # 같은 지역/학교 안의 조합도 인접 비교한다.
    entity_groups: dict[tuple[str, str], list[int]] = defaultdict(list)
    for index, row in enumerate(rows):
        entity_groups[(str(row["scope"]), str(row["entity"]))].append(index)
    for indexes in entity_groups.values():
        ordered = sorted(indexes, key=lambda i: str(rows[i]["slug"]))
        candidate_pairs.update(zip(ordered, ordered[1:]))

    title_high, desc_high, body_high, template_rows = [], [], [], []
    threshold_counts = Counter()
    by_type: dict[str, list[float]] = defaultdict(list)
    simple_replacement_pages: set[str] = set()
    for left_index, right_index in sorted(candidate_pairs):
        a, b = rows[left_index], rows[right_index]
        entities = (str(a["entity"]), str(b["entity"]))
        values = {}
        for field in ("title", "description", "body"):
            left = normalized(str(a[field]), entities)
            right = normalized(str(b[field]), entities)
            values[field] = max(jaccard(left, right), SequenceMatcher(None, left, right).ratio())
        body_score = values["body"]
        simple = normalized(str(a["body"]), entities) == normalized(str(b["body"]), entities)
        common = common_sentences(str(a["body"]), str(b["body"]))
        page_type = f"{a['page_type']}|{b['page_type']}"
        by_type[page_type].append(body_score)
        rewrite = body_score >= .80 or values["description"] >= .85 or values["title"] >= .90 or simple
        base = {
            "페이지 A": a["slug"], "페이지 B": b["slug"], "페이지 유형": page_type,
            "공통 문장 수": common, "단순 지역명·학교명 치환 여부": simple,
            "재작성 필요 여부": rewrite,
        }
        for field, target in (("title", title_high), ("description", desc_high), ("body", body_high)):
            if values[field] >= .70:
                target.append({**base, "유사도": round(values[field], 6)})
        if simple:
            simple_replacement_pages.update((str(a["slug"]), str(b["slug"])))
            template_rows.append({**base, "유사도": round(body_score, 6)})
        for threshold in (.95, .90, .80, .70):
            if body_score >= threshold:
                threshold_counts[str(threshold)] += 1

    write_csv(REPORTS / "title_similarity_high.csv", SIM_FIELDS, title_high)
    write_csv(REPORTS / "description_similarity_high.csv", SIM_FIELDS, desc_high)
    write_csv(REPORTS / "body_similarity_high.csv", SIM_FIELDS, body_high)
    write_csv(REPORTS / "template_like_pages.csv", SIM_FIELDS, template_rows)
    type_rows = [{
        "페이지 유형": key,
        "비교 쌍 수": len(scores),
        "평균 body 유사도": round(sum(scores) / len(scores), 6),
        "80% 이상": sum(score >= .8 for score in scores),
        "90% 이상": sum(score >= .9 for score in scores),
    } for key, scores in sorted(by_type.items())]
    write_csv(
        REPORTS / "similarity_by_page_type.csv",
        ("페이지 유형", "비교 쌍 수", "평균 body 유사도", "80% 이상", "90% 이상"),
        type_rows,
    )
    summary = {
        "method": "의미 그룹 + 엔터티 그룹 후보 추출, 엔터티 치환 정규화, bigram Jaccard와 SequenceMatcher 최댓값",
        "planned_page_count": len(rows),
        "candidate_pair_count": len(candidate_pairs),
        "body_pairs_ge_95": threshold_counts["0.95"],
        "body_pairs_ge_90": threshold_counts["0.9"],
        "body_pairs_ge_80": threshold_counts["0.8"],
        "body_pairs_ge_70": threshold_counts["0.7"],
        "simple_replacement_page_count": len(simple_replacement_pages),
        "title_rewrite_candidates": sum(float(row["유사도"]) >= .9 for row in title_high),
        "description_rewrite_candidates": sum(float(row["유사도"]) >= .85 for row in desc_high),
        "body_rewrite_candidates": sum(float(row["유사도"]) >= .8 for row in body_high),
        "comparison_scope_note": "64,104개 전 쌍이 아닌 고위험 그룹 후보 비교 결과",
    }
    (REPORTS / "similarity_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    return summary


def matching_reports(
    links: list[dict[str, object]], regions: list[dict[str, object]],
    unmatched: list[dict[str, object]], schools: dict[str, dict[str, object]],
    plans: list[dict[str, object]],
) -> dict[str, int]:
    region_names = {str(row["지역명"]) for row in regions}
    school_names = set(schools)
    mismatch = []
    integrity = []
    grade_mismatch = []
    duplicate = []
    primary = []
    seen_school_slugs = set()
    for row in links:
        if str(row["지역명"]) not in region_names:
            mismatch.append({**row, "오류": "지역 요약에 없는 연결 지역"})
        if str(row["학교명"]) not in school_names:
            integrity.append({**row, "오류": "학교 마스터에 없는 학교명"})
        expected = "중학교" if str(row["학교명"]).endswith("중학교") else "고등학교" if str(row["학교명"]).endswith("고등학교") else ""
        if expected and expected != row["학교급"]:
            grade_mismatch.append({**row, "오류": "학교 공식명과 학교급 불일치"})
    for plan in plans:
        if plan["scope"] == "school":
            slug = str(plan["slug"])
            if slug in seen_school_slugs:
                duplicate.append({**plan, "오류": "학교 페이지 중복"})
            seen_school_slugs.add(slug)
    for name, school in schools.items():
        if school["대표지역"] not in school["관련지역"]:
            primary.append({"학교명": name, **school, "오류": "대표지역이 관련지역에 없음"})
    # 미매칭은 오류가 아니라 원본 상태로 유지되었는지 확인한다.
    for row in unmatched:
        if row["지역명"] not in region_names:
            mismatch.append({**row, "오류": "미매칭 지역이 지역 요약에 없음"})
    report_specs = (
        ("region_school_mismatch.csv", mismatch),
        ("school_name_integrity.csv", integrity),
        ("school_grade_mismatch.csv", grade_mismatch),
        ("school_duplicate_page_check.csv", duplicate),
        ("school_primary_region_check.csv", primary),
    )
    for filename, items in report_specs:
        fields = sorted({key for item in items for key in item}) or ["오류"]
        write_csv(REPORTS / filename, fields, items)
    return {filename: len(items) for filename, items in report_specs}


def priority_reports(plans: list[dict[str, object]], estimate_bytes: int, estimate_seconds: float) -> dict[str, object]:
    counts = Counter(priority_of(plan) for plan in plans)
    type_counts = Counter(str(plan["page_type"]) for plan in plans)
    rows = []
    redundant = []
    for plan in plans:
        priority = priority_of(plan)
        reason = PRIORITY_LABELS[priority]
        rows.append({
            "priority": priority, "priority_label": reason, "page_type": plan["page_type"],
            "scope": plan["scope"], "keyword": plan["keyword"], "slug": plan["slug"],
            "generation_phase": generation_phase(plan),
        })
        grade = str(plan["학년표현"])
        if grade in ("중등", "고등") or bool(plan["내신사용"]):
            redundant.append({
                "page_type": plan["page_type"], "keyword": plan["keyword"],
                "slug": plan["slug"], "candidate_reason": (
                    "중등/중학생·고등/고등학생 검색 의도 중첩" if grade in ("중등", "고등")
                    else "내신 세부 조합으로 낮은 우선순위"
                ),
                "action": "삭제하지 않음; 검색 수요 확인 후 결정",
            })
    write_csv(
        REPORTS / "page_priority_plan.csv",
        ("priority", "priority_label", "page_type", "scope", "keyword", "slug", "generation_phase"),
        rows,
    )
    write_csv(
        REPORTS / "redundant_combination_candidates.csv",
        ("page_type", "keyword", "slug", "candidate_reason", "action"),
        redundant,
    )
    avg = estimate_bytes / max(1, len(plans) + 12)
    modes = {
        "plan": 0,
        "preview": 284,
        "production-priority-1": counts[1],
        "production-priority-1-2": counts[1] + counts[2],
        "production-priority-1-3": counts[1] + counts[2] + counts[3],
        "production-all": len(plans),
    }
    scope = {
        "page_type_counts": dict(sorted(type_counts.items())),
        "priority_counts": {str(k): counts[k] for k in range(1, 6)},
        "redundant_candidate_count": len(redundant),
        "recommended_keep_count_priority_1_4": sum(counts[k] for k in range(1, 5)),
        "full_count": len(plans),
        "modes": {
            name: {
                "expected_content_pages": count,
                "expected_bytes": round(count * avg),
                "expected_seconds": round(estimate_seconds * count / max(1, len(plans)), 1),
            } for name, count in modes.items()
        },
        "note": "조합은 삭제하지 않았으며 우선순위와 삭제 검토 후보만 분류함",
    }
    (REPORTS / "recommended_page_scope.json").write_text(
        json.dumps(scope, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    return scope


def write_priority_sitemaps(
    output: Path, records: list[dict[str, object]], plans_by_slug: dict[str, dict[str, object]],
    settings: dict[str, object], max_urls: int,
) -> dict[str, object]:
    grouped: dict[int, list[str]] = defaultdict(list)
    for record in records:
        path = str(record["path"])
        slug = str(record.get("slug", ""))
        plan = plans_by_slug.get(slug)
        priority = 1 if path == "/" or plan is None else priority_of(plan)
        grouped[priority].append(site_join(str(settings["site_url"]), path))
    files, duplicates = [], 0
    all_seen = set()
    for priority in range(1, 6):
        urls = []
        for url in grouped[priority]:
            if url in all_seen:
                duplicates += 1
                continue
            all_seen.add(url)
            urls.append(url)
        for offset in range(0, len(urls), max_urls):
            chunk = urls[offset:offset + max_urls]
            part_count = (len(urls) + max_urls - 1) // max_urls
            name = (
                f"sitemap-priority-{priority}.xml" if part_count <= 1
                else f"sitemap-priority-{priority}-{offset // max_urls + 1:03d}.xml"
            )
            xml = ['<?xml version="1.0" encoding="UTF-8"?>',
                   '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">']
            xml.extend(f"  <url><loc>{xml_escape(url)}</loc></url>" for url in chunk)
            xml.append("</urlset>")
            (output / name).write_text("\n".join(xml) + "\n", encoding="utf-8")
            files.append(name)
    index = ['<?xml version="1.0" encoding="UTF-8"?>',
             '<sitemapindex xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">']
    index.extend(
        f"  <sitemap><loc>{xml_escape(site_join(str(settings['site_url']), '/' + name))}</loc></sitemap>"
        for name in files
    )
    index.append("</sitemapindex>")
    (output / "sitemap_index.xml").write_text("\n".join(index) + "\n", encoding="utf-8")
    return {"files": files, "counts": {str(k): len(grouped[k]) for k in range(1, 6)},
            "duplicate_url_count": duplicates}


def link_reports(validation: dict[str, object]) -> None:
    metrics = list(validation["page_metrics"])
    write_csv(REPORTS / "click_depth_report.csv",
              ("url_path", "click_depth", "is_orphan"), metrics)
    write_csv(REPORTS / "internal_link_count_report.csv",
              ("url_path", "internal_link_count", "click_depth"), metrics)
    write_csv(REPORTS / "orphan_page_report.csv",
              ("url_path", "click_depth", "internal_link_count", "is_orphan"),
              [row for row in metrics if row["is_orphan"]])
    write_csv(REPORTS / "overlinked_pages.csv",
              ("url_path", "internal_link_count", "click_depth"),
              [row for row in metrics if int(row["internal_link_count"]) > 30])


def rewrite_sample_report(input_path: Path) -> dict[str, object]:
    """고위험 쌍만 실제로 다시 생성해 유형별 보강 전후를 비교한다."""
    settings = json.loads(SETTINGS_PATH.read_text(encoding="utf-8-sig"))
    config = dict(settings["school_region_generation"])
    links, regions, _, _ = read_excel(input_path)
    schools = build_school_records(links)
    plans, _, _ = plan_pages(regions, links, schools, config)
    plan_by_slug = {str(plan["slug"]): plan for plan in plans}
    with (REPORTS / "body_similarity_high.csv").open(encoding="utf-8-sig", newline="") as handle:
        candidates = [row for row in csv.DictReader(handle) if float(row["유사도"]) >= .7]
    high, selected_pages = [], set()
    for row in candidates:
        pair_pages = {row["페이지 A"], row["페이지 B"]}
        if len(selected_pages | pair_pages) > 300:
            continue
        high.append(row)
        selected_pages.update(pair_pages)
        if len(selected_pages) >= 100 and len(high) >= 200:
            break
    wanted = {row[key] for row in high for key in ("페이지 A", "페이지 B")}
    pools = load_pools()
    used: set[str] = set()
    contents: dict[str, str] = {}
    for slug in sorted(wanted):
        plan = plan_by_slug.get(slug)
        if not plan:
            continue
        content, _ = create_title(
            plan, pools, used,
            int(settings["content_generation"]["title_min_length"]),
            int(settings["content_generation"]["title_max_length"]),
        )
        contents[slug] = str(content["body"])

    structures = {
        "region_base": "지역의 학교 분포와 이동 동선을 기준으로 상담 범위를 나누어 확인합니다.",
        "region_subject": "과목별 진단 순서와 주간 복습 간격을 지역 학습 일정에 맞춰 조정합니다.",
        "region_grade": "학년 전환 시기의 교과 범위와 평가 방식을 먼저 구분해 계획합니다.",
        "region_exam": "학교별 시험 일정과 수행평가 반영 방식을 확인한 뒤 내신 계획을 세웁니다.",
        "school_base": "학교 교육과정과 평가 공지를 토대로 개인별 학습 순서를 점검합니다.",
        "school_subject": "해당 학교의 과목별 평가 범위와 서술형 출제 흐름을 함께 살펴봅니다.",
        "school_exam": "교과서 진도와 수행평가 일정을 기준으로 내신 준비 시점을 구체화합니다.",
    }
    output_rows = []
    for row in high:
        slug_a, slug_b = row["페이지 A"], row["페이지 B"]
        if slug_a not in contents or slug_b not in contents:
            continue
        plan_a, plan_b = plan_by_slug[slug_a], plan_by_slug[slug_b]
        before_a, before_b = contents[slug_a], contents[slug_b]

        def enhance(plan: dict[str, object], body: str) -> str:
            page_type = str(plan["page_type"])
            base = structures.get(page_type, structures.get(page_type.rsplit("_", 1)[0], structures["region_base"]))
            variants = (
                "진단 결과는 월별 목표와 연결해 점검합니다.",
                "상담에서는 현재 성취도와 학습 가능 시간을 함께 반영합니다.",
                "복습 기록은 다음 수업의 문제 난이도를 결정하는 기준으로 활용합니다.",
                "평가 후에는 오답 원인을 개념·계산·해석 단계로 나누어 기록합니다.",
                "학습 계획은 학교 일정 변화에 맞춰 주 단위로 다시 조정합니다.",
            )
            variant = variants[stable_int(str(plan["slug"])) % len(variants)]
            entity = str(plan["학교명"] or plan["지역명"])
            return f"{body} {entity} 학습에서는 {base} {variant}"

        after_a, after_b = enhance(plan_a, before_a), enhance(plan_b, before_b)
        entities = (
            str(plan_a["학교명"] or plan_a["지역명"]),
            str(plan_b["학교명"] or plan_b["지역명"]),
        )
        before_score = max(
            jaccard(normalized(before_a, entities), normalized(before_b, entities)),
            SequenceMatcher(None, normalized(before_a, entities), normalized(before_b, entities)).ratio(),
        )
        after_score = max(
            jaccard(normalized(after_a, entities), normalized(after_b, entities)),
            SequenceMatcher(None, normalized(after_a, entities), normalized(after_b, entities)).ratio(),
        )
        output_rows.append({
            "페이지 A": slug_a, "페이지 B": slug_b,
            "페이지 유형": row["페이지 유형"],
            "보강 전 유사도": round(before_score, 6),
            "보강 후 유사도": round(after_score, 6),
            "유사도 감소": round(before_score - after_score, 6),
            "적용 규칙": "유형별 문장 구조 + 결정적 보조 문장",
            "재작성 상태": "샘플 콘텐츠 재생성 완료",
        })
    write_csv(
        REPORTS / "similarity_rewrite_sample.csv",
        ("페이지 A", "페이지 B", "페이지 유형", "보강 전 유사도", "보강 후 유사도",
         "유사도 감소", "적용 규칙", "재작성 상태"),
        output_rows,
    )
    result = {
        "sample_pair_count": len(output_rows),
        "pages_regenerated": len({r[k] for r in output_rows for k in ("페이지 A", "페이지 B")}),
        "average_before": round(sum(float(r["보강 전 유사도"]) for r in output_rows) / max(1, len(output_rows)), 6),
        "average_after": round(sum(float(r["보강 후 유사도"]) for r in output_rows) / max(1, len(output_rows)), 6),
        "improved_pair_count": sum(float(r["유사도 감소"]) > 0 for r in output_rows),
        "note": "샘플 재작성 규칙만 검증했으며 전체 문장 풀에는 아직 반영하지 않음",
    }
    (REPORTS / "similarity_rewrite_sample_summary.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    return result


def run(input_path: Path, mode: str) -> dict[str, object]:
    started = time.perf_counter()
    settings = json.loads(SETTINGS_PATH.read_text(encoding="utf-8-sig"))
    config = dict(settings["school_region_generation"])
    quality = dict(config["quality_review"])
    links, regions, unmatched, normalization = read_excel(input_path)
    schools = build_school_records(links)
    plans, duplicate_rows, school_keyword_validation = plan_pages(regions, links, schools, config)
    plans = [plan for plan in plans if generation_phase(plan) in set(config["enabled_generation_phases"])]
    REPORTS.mkdir(parents=True, exist_ok=True)
    match = matching_reports(links, regions, unmatched, schools, plans)
    previous = json.loads((REPORTS / "page_plan_summary.json").read_text(encoding="utf-8"))
    scope = priority_reports(
        plans, int(previous["estimated_full_output_size_bytes"]),
        float(previous["estimated_full_generation_seconds"]),
    )
    inventory = content_inventory(plans, settings)
    similarity = similarity_reports(inventory, quality)

    sample, selected_regions, selected_schools = select_sample(plans, regions, schools, config)
    output = choose_output()
    generated = generate_preview(output, sample, selected_regions, schools, links, settings, config)
    validation = validate_preview(output, generated, schools, settings)
    priority_maps = write_priority_sitemaps(
        output, list(generated["records"]), {str(p["slug"]): p for p in plans},
        settings, int(config["sitemap_max_urls"]),
    )
    link_reports(validation)
    school_error_count = sum(match.values())
    body_massive = similarity["body_pairs_ge_90"] >= 100
    replacement_massive = similarity["simple_replacement_page_count"] >= 100
    blocking = {
        "school_region_errors": school_error_count,
        "orphan_pages": int(validation["orphan_page_count"]),
        "broken_links": int(validation["broken_internal_link_count"]) + int(validation["broken_image_count"]),
        "duplicate_slugs": len(duplicate_rows),
        "duplicate_canonical": int(validation["duplicate_canonical_count"]),
        "elementary_exam": int(validation["elementary_internal_exam_count"]),
        "school_forbidden_keyword": int(validation["school_invalid_keyword_count"]),
        "mass_body_similarity": body_massive,
        "mass_simple_replacement": replacement_massive,
    }
    recommended = not any(bool(value) for value in blocking.values())
    summary = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "mode": mode,
        "input": str(input_path),
        "normalization": normalization,
        "planned_content_pages": len(plans),
        "sample_output": str(output),
        "sample_html_count": int(generated["html_count"]),
        "similarity": similarity,
        "school_region_validation": match,
        "school_region_error_count": school_error_count,
        "page_scope": scope,
        "priority_sitemaps": priority_maps,
        "sample_validation": {k: v for k, v in validation.items() if k != "page_metrics"},
        "blocking_conditions": blocking,
        "full_generation_recommended": recommended,
        "duration_seconds": round(time.perf_counter() - started, 3),
        "production_executed": False,
        "site_modified": False,
    }
    (REPORTS / "school_region_quality_review_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return summary


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="전체 생성 전 지역·학교 품질 검수")
    parser.add_argument("--input", default=str(INPUT_DEFAULT))
    parser.add_argument("--mode", choices=("plan", "preview"), default="preview")
    parser.add_argument("--rewrite-only", action="store_true")
    args = parser.parse_args(argv)
    if args.rewrite_only:
        print(json.dumps(rewrite_sample_report(Path(args.input)), ensure_ascii=False, indent=2))
        return 0
    summary = run(Path(args.input), args.mode)
    return 0 if summary["sample_validation"]["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
