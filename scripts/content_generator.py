from __future__ import annotations

import hashlib
import json
import random
import re
from collections import Counter
from pathlib import Path


POOL_FILES = {
    "hooks": "hooks.json",
    "description_general": "description_general.json",
    "body_general": "body_general.json",
    "초등학생": "body_elementary.json",
    "중학생": "body_middle.json",
    "고등학생": "body_high.json",
    "국어": "body_korean.json",
    "영어": "body_english.json",
    "수학": "body_math.json",
    "국영수": "body_combined.json",
    "openings": "openings.json",
    "endings": "endings.json",
    "title_patterns": "title_patterns.json",
    "title_modifiers": "title_modifiers.json",
    "title_endings": "title_endings.json",
}
SUBJECT_MAP = {
    "국어": "국어", "국어과외": "국어",
    "영어": "영어", "영어과외": "영어",
    "수학": "수학", "수학과외": "수학",
    "국영수": "국영수", "국영수과외": "국영수", "종합": "국영수",
}
TARGET_EXPRESSIONS = {
    "초등학생": ["초등", "초등학생", "초등과정", "초등 공부", "초등 학습"],
    "중학생": ["중등", "중학생", "중등과정", "중등 내신", "중학생 공부", "중등 학습"],
    "고등학생": ["고등", "고등학생", "고등과정", "고등 내신", "고등학생 공부", "고등 학습"],
}
SUBJECT_EXPRESSIONS = {
    "국어": ["국어과외", "국어 학습", "국어 내신", "국어 독해", "국어 서술형"],
    "영어": ["영어과외", "영어 학습", "영어 내신", "영어 독해", "영어 문법"],
    "수학": ["수학과외", "수학 학습", "수학 내신", "수학 개념", "수학 문제풀이"],
    "국영수": ["국영수과외", "국영수 학습", "국영수 내신", "주요 과목 학습", "국어 영어 수학 학습"],
}
FORBIDDEN = ["학원", "지점", "강의실", "칠판", "상담", "전화", "가격", "최강", "1위", "압도적", "무조건", "반드시 성적 향상"]
ABNORMAL_RE = re.compile(r"[\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff\ufffd]")


def stable_random(slug: str, content_type: str, retry: int = 0) -> random.Random:
    digest = hashlib.sha256(f"{slug}:{content_type}:{retry}".encode("utf-8")).hexdigest()
    return random.Random(int(digest, 16))


def load_pools(pool_dir: Path) -> dict[str, list[dict[str, str]]]:
    pools: dict[str, list[dict[str, str]]] = {}
    for key, filename in POOL_FILES.items():
        items = json.loads((pool_dir / filename).read_text(encoding="utf-8-sig"))
        ids = [item["id"] for item in items]
        texts = [item["text"].strip() for item in items]
        if len(ids) != len(set(ids)) or len(texts) != len(set(texts)):
            raise ValueError(f"{filename}에 중복 ID 또는 중복 문장이 있습니다.")
        if any(not item_id or not text for item_id, text in zip(ids, texts)):
            raise ValueError(f"{filename}에 빈 ID 또는 문장이 있습니다.")
        pools[key] = [{"id": item_id, "text": text} for item_id, text in zip(ids, texts)]
    return pools


def choose_unique(items: list[dict[str, str]], count: int, rng: random.Random) -> list[dict[str, str]]:
    if count > len(items):
        raise ValueError(f"문장 풀이 부족합니다: 요청 {count}, 보유 {len(items)}")
    return rng.sample(items, count)


def normalize_meaning(value: str) -> str:
    return re.sub(r"[\s,]+", "", value)


def make_keywords(keyword: str, region: str, target: str, subject: str) -> list[str]:
    candidates = [
        keyword,
        f"{region}{subject}",
        f"{region}{target}과외",
        f"{target}{subject}",
    ]
    result: list[str] = []
    meanings: set[str] = set()
    for item in candidates:
        meaning = normalize_meaning(item)
        if meaning and meaning not in meanings:
            meanings.add(meaning)
            result.append(item)
        if len(result) == 4:
            break
    return result


def insert_description_keywords(sentences: list[str], keyword: str) -> list[str]:
    output = list(sentences)
    output[0] = f"{keyword}를 살펴볼 때, {output[0]}"
    middle = min(2, len(output) - 1)
    output[middle] = f"{output[middle]} 이 과정은 {keyword} 학습에도 연결됩니다."
    output[-1] = f"{output[-1]} {keyword}의 학습 흐름도 같은 기준으로 점검할 수 있습니다."
    return output


def insert_body_keywords(sentences: list[str], keyword: str) -> list[str]:
    output = list(sentences)
    if len(output) > 2:
        output[2] = f"{keyword} 학습에서는 {output[2]}"
    if len(output) > 4:
        output[4] = f"이러한 과정은 {keyword}를 알아볼 때에도 중요하며, {output[4]}"
    return output


def create_title(
    slug: str,
    keyword: str,
    region: str,
    target: str,
    subject_key: str,
    pools: dict[str, list[dict[str, str]]],
    used_titles: set[str],
    config: dict[str, object],
) -> tuple[str, dict[str, str], int]:
    minimum = int(config.get("title_min_length", 45))
    target_length = int(config.get("title_target_length", 65))
    maximum = int(config.get("title_max_length", 90))
    candidates: list[tuple[int, int, str, dict[str, str], int]] = []
    for retry in range(30):
        rng = stable_random(slug, "title", retry)
        pattern = rng.choice(pools["title_patterns"])
        target_expression = rng.choice(TARGET_EXPRESSIONS[target])
        subject_expression = rng.choice(SUBJECT_EXPRESSIONS[subject_key])
        modifier = rng.choice(pools["title_modifiers"])
        ending = rng.choice(pools["title_endings"])
        components = {
            "pattern_id": pattern["id"],
            "pattern": pattern["text"],
            "target_expression": target_expression,
            "subject_expression": subject_expression,
            "modifier": modifier["text"],
            "ending": ending["text"],
        }
        values = {
            "메인키워드": keyword,
            "지역명": region,
            "대상표현": target_expression,
            "과목표현": subject_expression,
            "수식어": modifier["text"],
            "마무리": ending["text"],
        }
        title = re.sub(r"\s+", " ", pattern["text"].format(**values)).strip()
        if len(title) > maximum:
            title = f"{keyword} {modifier['text']} {ending['text']} | {region} {target_expression} {subject_expression} 정보"
        if len(title) > maximum:
            title = f"{keyword} {ending['text']} | {region} {target_expression} {subject_expression} 정보"
        if len(title) < minimum:
            title = f"{title} | {region} 맞춤 {target_expression} {subject_expression} 학습 정보"
        if minimum <= len(title) <= maximum:
            duplicate_penalty = 1 if title in used_titles else 0
            candidates.append((duplicate_penalty, abs(len(title) - target_length), title, components, retry))
    if not candidates:
        fallback = f"{keyword} 학습 흐름과 공부 기준을 정리합니다 | {region} 맞춤 {target} {subject_key} 정보"
        if len(fallback) > maximum:
            fallback = f"{keyword} 학습 흐름과 공부 기준을 정리합니다"
        return fallback, {
            "pattern_id": "fallback",
            "pattern": "fallback",
            "target_expression": target,
            "subject_expression": subject_key,
            "modifier": "학습 흐름을 살펴보는",
            "ending": "공부 기준을 정리합니다",
        }, 30
    candidates.sort(key=lambda item: (item[0], item[1], item[4]))
    _, _, title, components, retry = candidates[0]
    return title, components, retry


def generate_page_content(
    row: dict[str, str],
    pools: dict[str, list[dict[str, str]]],
    config: dict[str, object],
    previous_body_sets: list[set[str]],
    used_titles: set[str],
    used_descriptions: set[str],
    used_bodies: set[str],
    used_combinations: set[tuple[str, ...]],
) -> dict[str, object]:
    region = row["지역명"].strip()
    target = row["대상"].strip()
    subject_raw = row["과목"].strip()
    subject_key = SUBJECT_MAP.get(subject_raw)
    if target not in ("초등학생", "중학생", "고등학생"):
        raise ValueError(f"지원하지 않는 대상: {target}")
    if not subject_key:
        raise ValueError(f"지원하지 않는 과목: {subject_raw}")
    keyword = row.get("메인키워드", "").strip() or f"{region} {target} {subject_raw}"
    slug = row.get("슬러그", "").strip() or re.sub(r"\s+", "", keyword)
    description_count = int(config.get("description_sentence_count", 4))
    body_count = int(config.get("body_sentence_count", 7))
    max_share = float(config.get("max_similarity_share", 0.7))
    max_retry = int(config.get("max_retry", 20))
    retry_used = 0
    similarity_warning = False

    for retry in range(max_retry + 1):
        title, title_components, title_retry = create_title(
            slug, keyword, region, target, subject_key, pools, used_titles, config
        )

        description_items = choose_unique(
            pools["description_general"], description_count, stable_random(slug, "description", retry)
        )
        description = " ".join(insert_description_keywords([item["text"] for item in description_items], keyword))
        description = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", description)).strip()

        body_rng = stable_random(slug, "body", retry)
        middle_count = max(0, body_count - 2)
        general_count = max(1, middle_count - 3)
        target_count = min(2, max(1, middle_count - general_count - 1))
        subject_count = max(1, middle_count - general_count - target_count)
        body_items = (
            choose_unique(pools["openings"], 1, body_rng)
            + choose_unique(pools["body_general"], general_count, body_rng)
            + choose_unique(pools[target], target_count, body_rng)
            + choose_unique(pools[subject_key], subject_count, body_rng)
            + choose_unique(pools["endings"], 1, body_rng)
        )
        body_items = body_items[:body_count]
        body_sentences = insert_body_keywords([item["text"] for item in body_items], keyword)
        body_text = " ".join(body_sentences)
        body_html = f"{keyword}<br>\n" + body_text
        body_ids = tuple(item["id"] for item in body_items)
        body_id_set = set(body_ids)
        max_observed_share = max(
            ((len(body_id_set & other) / max(1, len(body_id_set))) for other in previous_body_sets),
            default=0.0,
        )
        duplicate = (
            title in used_titles
            or description in used_descriptions
            or body_text in used_bodies
            or body_ids in used_combinations
        )
        too_similar = max_observed_share >= max_share
        if not duplicate and not too_similar:
            retry_used = retry
            break
        if retry == max_retry:
            retry_used = retry
            similarity_warning = too_similar

    pool_names = []
    id_to_pool: dict[str, str] = {}
    for pool_name, items in pools.items():
        for item in items:
            id_to_pool[item["id"]] = pool_name
    all_ids = [item["id"] for item in description_items] + list(body_ids)
    for item_id in all_ids:
        pool_name = id_to_pool[item_id]
        if pool_name not in pool_names:
            pool_names.append(pool_name)

    return {
        "region": region,
        "target": target,
        "subject": subject_raw,
        "keyword": keyword,
        "slug": slug,
        "title": title,
        "title_components": title_components,
        "title_retry_count": title_retry,
        "description": description,
        "body_text": body_text,
        "body_html": body_html,
        "keywords": make_keywords(keyword, region, target, subject_raw),
        "description_ids": [item["id"] for item in description_items],
        "body_ids": list(body_ids),
        "all_sentence_ids": all_ids,
        "used_pools": pool_names,
        "pool_counts": dict(Counter(id_to_pool[item_id] for item_id in all_ids)),
        "retry_count": retry_used,
        "similarity_warning": similarity_warning,
        "body_id_set": body_id_set,
    }


def validate_content(content: dict[str, object], all_regions: set[str]) -> dict[str, int]:
    target = str(content["target"])
    subject = str(content["subject"])
    region = str(content["region"])
    text = " ".join((str(content["title"]), str(content["description"]), str(content["body_text"])))
    grade_mismatch = 0
    if target == "초등학생" and any(term in text for term in ("수능", "모의고사", "고등학교 내신")):
        grade_mismatch += 1
    if target == "중학생" and any(term in text for term in ("초등 저학년", "놀이학습")):
        grade_mismatch += 1
    if target == "고등학생" and any(term in text for term in ("초등 기초", "놀이학습")):
        grade_mismatch += 1
    subject_mismatch = 0
    if subject.startswith("국어") and any(term in text for term in ("수학 공식", "수학 계산", "영어 단어")):
        subject_mismatch += 1
    if subject.startswith("영어") and any(term in text for term in ("수학 공식", "수학 계산")):
        subject_mismatch += 1
    if subject.startswith("수학") and any(term in text for term in ("영어 단어", "영어 어휘")):
        subject_mismatch += 1
    other_region = sum(
        1 for candidate in all_regions
        if candidate != region
        and candidate not in region
        and region not in candidate
        and candidate in text
    )
    forbidden = sum(text.count(term) for term in FORBIDDEN)
    abnormal = len(ABNORMAL_RE.findall(text)) + text.count("�")
    return {
        "grade_mismatch": grade_mismatch,
        "subject_mismatch": subject_mismatch,
        "other_region": other_region,
        "forbidden": forbidden,
        "abnormal": abnormal,
    }
