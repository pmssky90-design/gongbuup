from __future__ import annotations

import csv
import hashlib
import html
import json
import math
import re
import sys
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor
from difflib import SequenceMatcher
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from build_school_region_preview import (  # noqa: E402
    SETTINGS_PATH,
    build_school_records,
    generation_phase,
    plan_pages,
    read_excel,
)
from full_plan_validation import compact_normalized_body  # noqa: E402
from school_region_quality_review import INPUT_DEFAULT, jaccard, normalized  # noqa: E402


REPORT_DIR = ROOT / "reports" / "production_full_repair"
CONTENT_TARGETS = REPORT_DIR / "repair_content_targets.csv"
EXECUTION = REPORT_DIR / "repair_execution_summary.json"
CONTENT_RE = re.compile(
    r'<div class="content">.*?<br>\s*(.*?)</div>', re.I | re.S
)


def sha(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def simhash(value: str) -> int:
    words = value.split()
    shingles = (
        [" ".join(words[index:index + 2]) for index in range(len(words) - 1)]
        if len(words) > 1 else words
    )
    vector = [0] * 64
    for shingle in shingles:
        number = int(hashlib.sha256(shingle.encode("utf-8")).hexdigest()[:16], 16)
        for bit in range(64):
            vector[bit] += 1 if number & (1 << bit) else -1
    result = 0
    for bit, score in enumerate(vector):
        if score >= 0:
            result |= 1 << bit
    return result


def read_target_slugs() -> set[str]:
    with CONTENT_TARGETS.open("r", encoding="utf-8-sig", newline="") as source:
        return {str(row["slug"]) for row in csv.DictReader(source)}


def extract_body(path: Path) -> tuple[str, str]:
    text = path.read_text(encoding="utf-8")
    match = CONTENT_RE.search(text)
    if not match:
        raise RuntimeError(f"본문 영역 누락: {path}")
    body = html.unescape(re.sub(r"<[^>]+>", " ", match.group(1)))
    return path.parent.name, " ".join(body.split())


def main() -> int:
    execution = json.loads(EXECUTION.read_text(encoding="utf-8"))
    output = Path(execution["repaired_candidate"])
    modified = read_target_slugs()
    settings = json.loads(SETTINGS_PATH.read_text(encoding="utf-8-sig"))
    config = settings["school_region_generation"]
    links, regions, _, _ = read_excel(INPUT_DEFAULT)
    schools = build_school_records(links)
    plans, _, _ = plan_pages(regions, links, schools, config)
    enabled = set(str(value) for value in config["enabled_generation_phases"])
    plans = [row for row in plans if generation_phase(row) in enabled]
    by_slug = {str(row["slug"]): row for row in plans}
    files = [
        output / "과외" / str(row["slug"]) / "index.html"
        for row in plans
    ]
    with ThreadPoolExecutor(max_workers=8) as pool:
        bodies = dict(pool.map(extract_body, files, chunksize=64))

    rows: list[dict[str, object]] = []
    exact_body = defaultdict(list)
    exact_normalized = defaultdict(list)
    compact_groups = defaultdict(list)
    entity_groups = defaultdict(list)
    lsh = defaultdict(list)
    for index, plan in enumerate(plans):
        slug = str(plan["slug"])
        body = bodies[slug]
        normal = normalized(
            body,
            (
                str(plan["지역명"]), str(plan["학교명"]),
                str(plan["과목표현"]), str(plan["학년표현"]),
            ),
        )
        row = {
            **plan,
            "body": body,
            "normalized_body": normal,
            "compact": compact_normalized_body({**plan, "body": body}),
        }
        rows.append(row)
        exact_body[sha(body)].append(index)
        exact_normalized[sha(normal)].append(index)
        compact_groups[sha(str(row["compact"]))].append(index)
        entity = str(plan["학교명"] or plan["지역명"])
        entity_groups[(str(plan["scope"]), entity)].append(index)
        fingerprint = simhash(str(row["compact"]))
        length_bucket = len(str(row["compact"])) // 80
        for band in range(4):
            key = (
                str(plan["page_type"]), length_bucket,
                band, (fingerprint >> (band * 16)) & 0xFFFF,
            )
            lsh[key].append(index)

    duplicate_body = sum(len(group) - 1 for group in exact_body.values() if len(group) > 1)
    duplicate_normalized = sum(
        len(group) - 1 for group in exact_normalized.values() if len(group) > 1
    )
    simple_replacement = sum(
        len(group) - 1 for group in compact_groups.values() if len(group) > 1
    )

    candidates: set[tuple[int, int]] = set()
    modified_indexes = {
        index for index, row in enumerate(rows) if str(row["slug"]) in modified
    }
    for index in modified_indexes:
        row = rows[index]
        entity = str(row["학교명"] or row["지역명"])
        for other in entity_groups[(str(row["scope"]), entity)]:
            if other != index:
                candidates.add((min(index, other), max(index, other)))
        fingerprint = simhash(str(row["compact"]))
        length_bucket = len(str(row["compact"])) // 80
        for nearby in (length_bucket - 1, length_bucket, length_bucket + 1):
            for band in range(4):
                key = (
                    str(row["page_type"]), nearby,
                    band, (fingerprint >> (band * 16)) & 0xFFFF,
                )
                for other in lsh.get(key, []):
                    if other != index:
                        candidates.add((min(index, other), max(index, other)))

    counts = Counter()
    high_rows: list[dict[str, object]] = []
    for left_index, right_index in sorted(candidates):
        if left_index not in modified_indexes and right_index not in modified_indexes:
            continue
        left_row, right_row = rows[left_index], rows[right_index]
        remove = (
            str(left_row["학교명"] or left_row["지역명"]),
            str(right_row["학교명"] or right_row["지역명"]),
            str(left_row["과목표현"]), str(right_row["과목표현"]),
            str(left_row["학년표현"]), str(right_row["학년표현"]),
        )
        left = normalized(str(left_row["body"]), remove)
        right = normalized(str(right_row["body"]), remove)
        token_score = jaccard(left, right)
        # 95% 후보는 SimHash/엔터티 버킷에서 정밀 SequenceMatcher로 확인한다.
        sequence_score = SequenceMatcher(None, left, right).ratio()
        score = max(token_score, sequence_score)
        if score >= .80:
            counts["body_ge_80"] += 1
        if score >= .90:
            counts["body_ge_90"] += 1
        if score >= .95:
            counts["body_ge_95"] += 1
            high_rows.append({
                "page_a": left_row["slug"],
                "page_b": right_row["slug"],
                "page_type": (
                    str(left_row["page_type"]) + "|" + str(right_row["page_type"])
                ),
                "similarity": round(score, 6),
                "left_modified": left_index in modified_indexes,
                "right_modified": right_index in modified_indexes,
            })

    fields = (
        "page_a", "page_b", "page_type", "similarity",
        "left_modified", "right_modified",
    )
    with (REPORT_DIR / "repair_similarity_remaining_high.csv").open(
        "w", encoding="utf-8-sig", newline=""
    ) as target:
        writer = csv.DictWriter(target, fieldnames=fields)
        writer.writeheader()
        writer.writerows(high_rows)
    summary = {
        "candidate": str(output),
        "content_page_count": len(rows),
        "modified_page_count": len(modified_indexes),
        "candidate_pair_count": len(candidates),
        "comparison_policy": (
            "modified-to-all first; entity groups + page_type/length/SimHash LSH; "
            "token Jaccard and SequenceMatcher precision check"
        ),
        "exact_body_duplicate_count": duplicate_body,
        "exact_normalized_body_duplicate_count": duplicate_normalized,
        "simple_replacement_count": simple_replacement,
        "body_ge_95": counts["body_ge_95"],
        "body_ge_90": counts["body_ge_90"],
        "body_ge_80": counts["body_ge_80"],
        "body_90_to_95": counts["body_ge_90"] - counts["body_ge_95"],
        "body_80_to_90": counts["body_ge_80"] - counts["body_ge_90"],
        "passed": (
            duplicate_body == 0
            and duplicate_normalized == 0
            and simple_replacement == 0
            and counts["body_ge_95"] == 0
        ),
    }
    (REPORT_DIR / "repair_similarity_validation.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if summary["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
