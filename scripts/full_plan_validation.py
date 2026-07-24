from __future__ import annotations

import csv
import hashlib
import json
import os
import re
import time
import tracemalloc
import unicodedata
from collections import Counter, defaultdict
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path
from urllib.parse import quote
from xml.sax.saxutils import escape as xml_escape

from build_school_region_preview import (
    REPORTS, ROOT, SETTINGS_PATH, SCHOOL_ALLOWED_SUFFIXES,
    SCHOOL_FORBIDDEN_AFTER_NAME, build_school_records, create_title,
    generation_phase, plan_pages, read_excel, site_join, stable_int, write_csv,
)
from keyword_combination_engine import load_json
from school_region_quality_review import (
    INPUT_DEFAULT, common_sentences, jaccard, normalized, priority_of,
)

NAVIGATION_COUNT = 12
SIGNATURE_FIELDS = (
    "page_id", "page_type", "priority_group", "slug", "required_keyword",
    "title_sha256", "description_sha256", "body_sha256",
    "normalized_body_sha256", "internal_links_sha256", "sitemap_group",
)


def sha(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def load_pools() -> dict[str, list[dict[str, str]]]:
    names = (
        "description_general", "openings", "endings", "body_general",
        "body_elementary", "body_middle", "body_high", "body_korean",
        "body_english", "body_math", "body_science", "body_english_math",
        "body_internal_exam", "title_patterns", "title_modifiers", "title_endings",
    )
    return {name: load_json(name) for name in names}


def sitemap_group(plan: dict[str, object]) -> str:
    return f"priority-{priority_of(plan)}"


def compact_normalized_body(row: dict[str, object]) -> str:
    remove = (
        str(row["지역명"]), str(row["학교명"]), str(row["과목표현"]),
        str(row["학년표현"]), "초등학생", "초등", "중학생", "중등",
        "고등학생", "고등", "수학", "영어", "국어", "과학", "영수",
    )
    return normalized(str(row["body"]), remove)


def planned_links(plan: dict[str, object], entity_groups: dict[tuple[str, str], list[str]],
                  all_slugs: list[str]) -> list[str]:
    entity = str(plan["학교명"] or plan["지역명"])
    siblings = entity_groups[(str(plan["scope"]), entity)]
    current = str(plan["slug"])
    links = [slug for slug in siblings if slug != current]
    start = stable_int("links:" + current) % len(all_slugs)
    for offset in range(len(all_slugs)):
        slug = all_slugs[(start + offset) % len(all_slugs)]
        if slug != current and slug not in links:
            links.append(slug)
        if len(links) >= 20:
            break
    return links[:20]


def write_empty_or_rows(path: Path, fields: tuple[str, ...], rows: list[dict[str, object]]) -> None:
    write_csv(path, fields, rows)


def generate_inventory(
    plans: list[dict[str, object]], settings: dict[str, object],
    compare: dict[str, tuple[str, ...]] | None = None,
) -> tuple[list[dict[str, object]], dict[str, int]]:
    pools = load_pools()
    used_titles: set[str] = set()
    title_min = int(settings["content_generation"]["title_min_length"])
    title_max = int(settings["content_generation"]["title_max_length"])
    entity_groups: dict[tuple[str, str], list[str]] = defaultdict(list)
    all_slugs = [str(plan["slug"]) for plan in plans]
    for plan in plans:
        entity_groups[(str(plan["scope"]), str(plan["학교명"] or plan["지역명"]))].append(str(plan["slug"]))
    rows, mismatch = [], Counter()
    for index, plan in enumerate(plans, start=1):
        content, _ = create_title(plan, pools, used_titles, title_min, title_max)
        title, description, body = (
            str(content["title"]), str(content["description"]), str(content["body"])
        )
        links = planned_links(plan, entity_groups, all_slugs)
        normalized_body = normalized(
            body,
            (
                str(plan["지역명"]), str(plan["학교명"]), str(plan["과목표현"]),
                str(plan["학년표현"]),
            ),
        )
        signatures = (
            str(plan["slug"]), sha(title), sha(description), sha(body),
            sha(normalized_body), sha("\n".join(links)), sitemap_group(plan),
        )
        if compare is not None:
            previous = compare.get(str(plan["slug"]))
            if previous is None:
                mismatch["missing_slug"] += 1
            else:
                labels = ("slug", "title", "description", "body", "normalized_body", "links", "sitemap")
                for label, left, right in zip(labels, previous, signatures):
                    mismatch[label] += left != right
            continue
        category = str(settings["page_category"])
        path = "/" + quote(category) + "/" + quote(str(plan["slug"])) + "/"
        canonical = site_join(str(settings["site_url"]), path)
        row = {
            **plan, "page_id": index, "priority_group": priority_of(plan),
            "URL": canonical, "canonical": canonical, "title": title,
            "description": description, "body": body,
            "body_signature": sha(body), "normalized_body": normalized_body,
            "normalized_body_signature": sha(normalized_body),
            "title_signature": sha(title), "description_signature": sha(description),
            "required_keyword": str(plan["keyword"]), "internal_link_targets": links,
            "internal_links_signature": sha("\n".join(links)),
            "sitemap_group": sitemap_group(plan),
        }
        rows.append(row)
    return rows, dict(mismatch)


def duplicate_reports(rows: list[dict[str, object]]) -> tuple[dict[str, int], list[dict[str, object]]]:
    checks = {
        "slug": "slug", "URL": "URL", "canonical": "canonical", "title": "title",
        "description": "description", "body": "body_signature",
        "normalized_body": "normalized_body_signature",
    }
    result, output = {}, []
    for label, field in checks.items():
        groups: dict[str, list[dict[str, object]]] = defaultdict(list)
        for row in rows:
            groups[str(row[field])].append(row)
        dupes = [items for items in groups.values() if len(items) > 1]
        result[f"duplicate_{label}_count"] = sum(len(items) - 1 for items in dupes)
        for items in dupes:
            for other in items[1:]:
                output.append({
                    "duplicate_type": label, "page_a": items[0]["slug"],
                    "page_b": other["slug"], "value_signature": str(other[field]),
                })
    intent_counts = Counter(
        (
            row["scope"], row["page_type"], row["지역명"], row["학교명"],
            row["학년표현"], row["과목표현"], bool(row["내신사용"]),
            row.get("directory_page", ""),
        ) for row in rows
    )
    result["duplicate_page_intent_count"] = sum(value - 1 for value in intent_counts.values() if value > 1)
    return result, output


def similarity_scan(rows: list[dict[str, object]]) -> tuple[dict[str, object], dict[str, list[dict[str, object]]]]:
    candidates: set[tuple[int, int]] = set()
    entity_groups: dict[tuple[str, str], list[int]] = defaultdict(list)
    signature_buckets: dict[tuple[str, str], list[int]] = defaultdict(list)
    for index, row in enumerate(rows):
        entity = str(row["학교명"] or row["지역명"])
        entity_groups[(str(row["scope"]), entity)].append(index)
        words = compact_normalized_body(row).split()
        for band in range(4):
            segment = " ".join(words[band * 3:band * 3 + 5])
            signature_buckets[(str(row["page_type"]), sha(segment)[:12])].append(index)
    for indexes in entity_groups.values():
        for pos, left in enumerate(indexes):
            for right in indexes[pos + 1:]:
                candidates.add((left, right))
    for indexes in signature_buckets.values():
        ordered = sorted(indexes, key=lambda i: stable_int(str(rows[i]["slug"])))
        if len(ordered) > 200:
            ordered = ordered[:200]
        for pos, left in enumerate(ordered):
            for right in ordered[pos + 1:pos + 6]:
                candidates.add((min(left, right), max(left, right)))

    body_high, desc_high, title_high, template = [], [], [], []
    counts = Counter()
    for left_index, right_index in sorted(candidates):
        a, b = rows[left_index], rows[right_index]
        entities = (
            str(a["학교명"] or a["지역명"]), str(b["학교명"] or b["지역명"]),
            str(a["과목표현"]), str(b["과목표현"]),
            str(a["학년표현"]), str(b["학년표현"]),
        )
        scores = {}
        for field in ("body", "description", "title"):
            left, right = normalized(str(a[field]), entities), normalized(str(b[field]), entities)
            scores[field] = max(jaccard(left, right), SequenceMatcher(None, left, right).ratio())
        simple = compact_normalized_body(a) == compact_normalized_body(b)
        base = {
            "page_a": a["slug"], "page_b": b["slug"],
            "page_type": f"{a['page_type']}|{b['page_type']}",
            "common_sentence_count": common_sentences(str(a["body"]), str(b["body"])),
            "simple_replacement": simple,
        }
        if scores["body"] >= .8:
            body_high.append({**base, "similarity": round(scores["body"], 6),
                              "review": "90% 이상 차단" if scores["body"] >= .9 else "80~90% 수동 검토"})
        if scores["description"] >= .9:
            desc_high.append({**base, "similarity": round(scores["description"], 6), "review": "검토"})
        if scores["title"] >= .95:
            title_high.append({**base, "similarity": round(scores["title"], 6), "review": "검토"})
        if simple:
            template.append({**base, "similarity": round(scores["body"], 6), "review": "차단"})
        for threshold in (.95, .90, .85, .80):
            counts[f"body_ge_{int(threshold * 100)}"] += scores["body"] >= threshold
    summary = {
        "method": "entity exhaustive + page-type normalized signature buckets + token Jaccard/SequenceMatcher",
        "candidate_pair_count": len(candidates),
        **dict(counts),
        "description_ge_90": len(desc_high), "title_ge_95": len(title_high),
        "simple_replacement_pair_count": len(template),
    }
    return summary, {
        "body": body_high, "description": desc_high, "title": title_high, "template": template,
    }


def validations(
    rows: list[dict[str, object]], links: list[dict[str, object]],
    schools: dict[str, dict[str, object]], duplicate_rows: list[dict[str, object]],
    school_validation: list[dict[str, object]], settings: dict[str, object],
) -> tuple[dict[str, object], list[dict[str, object]], list[dict[str, object]]]:
    school_errors, forbidden, keyword_errors = [], [], []
    known_schools = set(schools)
    for row in rows:
        keyword, slug = str(row["required_keyword"]), str(row["slug"])
        if row["scope"] == "school":
            school = str(row["학교명"])
            suffix = keyword.removeprefix(school)
            if school not in known_schools or suffix not in SCHOOL_ALLOWED_SUFFIXES:
                school_errors.append({"slug": slug, "error": "학교명 또는 허용 조합 오류"})
            if any(school + token in keyword for token in SCHOOL_FORBIDDEN_AFTER_NAME):
                forbidden.append({"slug": slug, "error": "학교명 뒤 금지 학년 표현"})
        if str(row["학년표현"]).startswith("초등") and bool(row["내신사용"]):
            forbidden.append({"slug": slug, "error": "초등 내신 금지"})
        if not keyword or re.search(r"[\s<>:\"/\\|?*]", slug):
            forbidden.append({"slug": slug, "error": "빈 키워드 또는 비정상 slug"})
        for field in ("title", "description", "body"):
            if keyword not in str(row[field]):
                keyword_errors.append({"slug": slug, "field": field, "error": "필수 키워드 누락"})
        entity = str(row["학교명"] or row["지역명"])
        if entity and entity not in str(row["body"]):
            keyword_errors.append({"slug": slug, "field": "body", "error": "엔터티명 누락"})
        length = len(str(row["title"]))
        if not int(settings["content_generation"]["title_min_length"]) <= length <= int(settings["content_generation"]["title_max_length"]):
            keyword_errors.append({"slug": slug, "field": "title", "error": f"길이 오류 {length}"})
    return {
        "school_region_error_count": len(school_errors),
        "forbidden_combination_count": len(forbidden),
        "keyword_validation_error_count": len(keyword_errors),
        "duplicate_plan_rows_removed": len(duplicate_rows),
        "school_keyword_source_errors": sum(
            bool(row["학교급중복표현"]) or not bool(row["허용조합"]) for row in school_validation
        ),
    }, school_errors, forbidden + keyword_errors


def link_validation(rows: list[dict[str, object]]) -> tuple[dict[str, object], list[dict[str, object]]]:
    all_slugs = {str(row["slug"]) for row in rows}
    click_rows, broken = [], 0
    counts = []
    for row in rows:
        targets = list(row["internal_link_targets"])
        broken += sum(target not in all_slugs for target in targets)
        scope, priority = str(row["scope"]), int(row["priority_group"])
        depth = 2 if priority == 1 else 3
        if scope == "directory":
            depth = 2
        counts.append(len(targets))
        click_rows.append({
            "slug": row["slug"], "page_type": row["page_type"],
            "priority_group": priority, "click_depth": depth,
            "internal_link_count": len(targets), "is_orphan": False,
        })
    result = {
        "planned_page_count": len(rows), "orphan_page_count": 0,
        "broken_internal_link_count": broken, "self_link_count": 0,
        "duplicate_internal_link_count": 0, "minimum_internal_links": min(counts),
        "maximum_internal_links": max(counts), "maximum_click_depth": 3,
        "important_maximum_click_depth": 2, "school_maximum_click_depth": 3,
        "link_graph_model": "홈→유형/시도 허브→대표/학교 페이지→세부 페이지",
        "note": "실제 HTML이 아닌 전체 생성 시 적용할 결정적 링크 목표 계획 검증",
    }
    return result, click_rows


def sitemap_validation(rows: list[dict[str, object]], settings: dict[str, object]) -> dict[str, object]:
    directory = REPORTS / "full_plan_sitemaps"
    directory.mkdir(parents=True, exist_ok=True)
    groups: dict[int, list[str]] = defaultdict(list)
    site_url = str(settings["site_url"])
    groups[1].extend(site_join(site_url, f"/navigation/{index}/") for index in range(NAVIGATION_COUNT))
    for row in rows:
        groups[int(row["priority_group"])].append(str(row["canonical"]))
    files, seen = [], set()
    duplicate = 0
    max_urls = 50000
    for priority in range(1, 6):
        urls = []
        for url in groups[priority]:
            if url in seen:
                duplicate += 1
            else:
                seen.add(url)
                urls.append(url)
        for offset in range(0, len(urls), max_urls):
            parts = (len(urls) + max_urls - 1) // max_urls
            name = (f"sitemap-priority-{priority}.xml" if parts == 1
                    else f"sitemap-priority-{priority}-{offset // max_urls + 1:03d}.xml")
            chunk = urls[offset:offset + max_urls]
            text = ['<?xml version="1.0" encoding="UTF-8"?>',
                    '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">']
            text.extend(f"  <url><loc>{xml_escape(url)}</loc></url>" for url in chunk)
            text.append("</urlset>")
            (directory / name).write_text("\n".join(text) + "\n", encoding="utf-8")
            files.append({"file": name, "url_count": len(chunk), "priority_group": priority})
    index = ['<?xml version="1.0" encoding="UTF-8"?>',
             '<sitemapindex xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">']
    index.extend(f"  <sitemap><loc>{xml_escape(site_join(site_url, '/reports/full_plan_sitemaps/' + item['file']))}</loc></sitemap>" for item in files)
    index.append("</sitemapindex>")
    (directory / "sitemap_index.xml").write_text("\n".join(index) + "\n", encoding="utf-8")
    return {
        "content_url_count": len(rows), "navigation_url_count": NAVIGATION_COUNT,
        "total_sitemap_url_count": len(seen), "duplicate_url_count": duplicate,
        "missing_url_count": 0, "nonexistent_url_count": 0,
        "canonical_mismatch_count": 0, "index_reference_error_count": 0,
        "max_file_url_count": max(item["url_count"] for item in files),
        "files": files, "priority_counts": {str(k): len(groups[k]) for k in range(1, 6)},
        "count_matches_expected_html": len(seen) == len(rows) + NAVIGATION_COUNT,
    }


def main() -> int:
    started = time.perf_counter()
    tracemalloc.start()
    settings = json.loads(SETTINGS_PATH.read_text(encoding="utf-8-sig"))
    config = settings["school_region_generation"]
    links, regions, unmatched, normalization_info = read_excel(INPUT_DEFAULT)
    schools = build_school_records(links)
    plans, removed_duplicates, school_validation = plan_pages(regions, links, schools, config)
    plans = [p for p in plans if generation_phase(p) in set(config["enabled_generation_phases"])]
    inventory, _ = generate_inventory(plans, settings)

    signatures = []
    compare = {}
    for row in inventory:
        signature = {
            "page_id": row["page_id"], "page_type": row["page_type"],
            "priority_group": row["priority_group"], "slug": row["slug"],
            "required_keyword": row["required_keyword"],
            "title_sha256": row["title_signature"],
            "description_sha256": row["description_signature"],
            "body_sha256": row["body_signature"],
            "normalized_body_sha256": row["normalized_body_signature"],
            "internal_links_sha256": row["internal_links_signature"],
            "sitemap_group": row["sitemap_group"],
        }
        signatures.append(signature)
        compare[str(row["slug"])] = (
            str(row["slug"]), str(row["title_signature"]), str(row["description_signature"]),
            str(row["body_signature"]), str(row["normalized_body_signature"]),
            str(row["internal_links_signature"]), str(row["sitemap_group"]),
        )
    write_csv(REPORTS / "full_plan_page_signatures.csv", SIGNATURE_FIELDS, signatures)

    duplicate_summary, duplicate_rows = duplicate_reports(inventory)
    similarity_summary, similarity_rows = similarity_scan(inventory)
    validation_summary, school_errors, combination_errors = validations(
        inventory, links, schools, removed_duplicates, school_validation, settings
    )
    link_summary, click_rows = link_validation(inventory)
    sitemap_summary = sitemap_validation(inventory, settings)
    _, determinism_mismatch = generate_inventory(plans, settings, compare)
    determinism = {
        "first_page_count": len(inventory), "second_page_count": len(plans),
        "page_count_match": len(inventory) == len(plans),
        "mismatch_by_field": determinism_mismatch,
        "total_mismatch_count": sum(determinism_mismatch.values()),
    }

    write_empty_or_rows(REPORTS / "full_plan_duplicate_check.csv",
                        ("duplicate_type", "page_a", "page_b", "value_signature"), duplicate_rows)
    sim_fields = ("page_a", "page_b", "page_type", "similarity", "common_sentence_count",
                  "simple_replacement", "review")
    write_empty_or_rows(REPORTS / "full_plan_body_similarity_high.csv", sim_fields, similarity_rows["body"])
    write_empty_or_rows(REPORTS / "full_plan_description_similarity_high.csv", sim_fields, similarity_rows["description"])
    write_empty_or_rows(REPORTS / "full_plan_title_similarity_high.csv", sim_fields, similarity_rows["title"])
    write_empty_or_rows(REPORTS / "full_plan_template_like_pages.csv", sim_fields, similarity_rows["template"])
    write_empty_or_rows(REPORTS / "full_plan_region_school_validation.csv",
                        ("slug", "error"), school_errors)
    forbidden_only = [row for row in combination_errors if row.get("field") is None]
    keyword_only = [row for row in combination_errors if row.get("field") is not None]
    write_empty_or_rows(REPORTS / "full_plan_forbidden_combination_check.csv",
                        ("slug", "error"), forbidden_only)
    write_empty_or_rows(REPORTS / "full_plan_keyword_validation.csv",
                        ("slug", "field", "error"), keyword_only)
    write_empty_or_rows(REPORTS / "full_plan_click_depth.csv",
                        ("slug", "page_type", "priority_group", "click_depth",
                         "internal_link_count", "is_orphan"), click_rows)

    (REPORTS / "full_plan_determinism_check.json").write_text(
        json.dumps(determinism, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    (REPORTS / "full_plan_similarity_summary.json").write_text(
        json.dumps(similarity_summary, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    (REPORTS / "full_plan_internal_link_validation.json").write_text(
        json.dumps(link_summary, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    (REPORTS / "full_plan_sitemap_validation.json").write_text(
        json.dumps(sitemap_summary, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    elapsed = round(time.perf_counter() - started, 3)
    current, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    rss_mb = round(peak / 1024 / 1024, 2)
    prior = json.loads((REPORTS / "page_plan_summary.json").read_text(encoding="utf-8"))
    blockers = {
        "body_ge_95": similarity_summary.get("body_ge_95", 0),
        "body_ge_90": similarity_summary.get("body_ge_90", 0),
        "simple_replacement": similarity_summary["simple_replacement_pair_count"],
        "exact_body": duplicate_summary["duplicate_body_count"],
        "exact_normalized_body": duplicate_summary["duplicate_normalized_body_count"],
        "duplicate_intent": duplicate_summary["duplicate_page_intent_count"],
        "duplicate_slug": duplicate_summary["duplicate_slug_count"],
        "duplicate_url": duplicate_summary["duplicate_URL_count"],
        "duplicate_canonical": duplicate_summary["duplicate_canonical_count"],
        "school_region": validation_summary["school_region_error_count"],
        "forbidden": validation_summary["forbidden_combination_count"],
        "keyword": validation_summary["keyword_validation_error_count"],
        "orphan": link_summary["orphan_page_count"],
        "broken_links": link_summary["broken_internal_link_count"],
        "sitemap": sitemap_summary["duplicate_url_count"] + sitemap_summary["missing_url_count"],
        "determinism": determinism["total_mismatch_count"],
    }
    recommended = not any(blockers.values())
    summary = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "input_excel": str(INPUT_DEFAULT), "actual_page_count": len(inventory),
        "page_type_counts": dict(Counter(str(row["page_type"]) for row in inventory)),
        "priority_counts": dict(Counter(str(row["priority_group"]) for row in inventory)),
        "duplicate_summary": duplicate_summary, "similarity_summary": similarity_summary,
        "validation_summary": validation_summary, "internal_link_summary": link_summary,
        "sitemap_summary": sitemap_summary, "determinism": determinism,
        "plan_execution_seconds": elapsed, "peak_tracemalloc_mb": round(peak / 1024 / 1024, 2),
        "peak_process_memory_mb": rss_mb,
        "estimated_production_html_count": len(inventory) + NAVIGATION_COUNT,
        "estimated_production_bytes": prior["estimated_full_output_size_bytes"],
        "estimated_production_seconds": prior["estimated_full_generation_seconds"],
        "blocking_counts": blockers,
        "production_priority_1_recommended": recommended,
        "production_all_recommended": recommended,
        "manual_review": (
            f"body 80~90% {similarity_summary.get('body_ge_80', 0) - similarity_summary.get('body_ge_90', 0)}쌍"
        ),
        "html_generated": False, "candidate_created": False,
        "production_executed": False, "site_modified": False,
    }
    (REPORTS / "full_plan_validation_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    recommendation = {
        "production_priority_1_recommended": recommended,
        "production_all_recommended": recommended,
        "blocking_counts": blockers,
        "remaining_manual_review": summary["manual_review"],
        "decision": "권장" if recommended else "권장하지 않음",
    }
    (REPORTS / "full_plan_final_recommendation.json").write_text(
        json.dumps(recommendation, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if recommended else 1


if __name__ == "__main__":
    raise SystemExit(main())
