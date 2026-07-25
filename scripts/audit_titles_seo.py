from __future__ import annotations

import html
import json
import os
import re
from collections import Counter
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SITE_DIR = ROOT / "site"
REPORT_PATH = ROOT / "reports" / "title_seo_audit_report.json"

TITLE_RE = re.compile(r"<title>(.*?)</title>", re.IGNORECASE | re.DOTALL)
H1_RE = re.compile(r"<h1[^>]*>(.*?)</h1>", re.IGNORECASE | re.DOTALL)
TAG_RE = re.compile(r"<[^>]+>")


def clean_text(value: str) -> str:
    return re.sub(r"\s+", " ", html.unescape(TAG_RE.sub("", value))).strip()


def html_files(root: Path) -> list[Path]:
    paths: list[Path] = []
    for current, dirs, files in os.walk(root):
        dirs.sort()
        for filename in sorted(files):
            if filename.lower().endswith(".html"):
                paths.append(Path(current) / filename)
    return paths


def shingles(value: str, size: int = 3) -> set[str]:
    compact = re.sub(r"\s+", "", value)
    if len(compact) <= size:
        return {compact}
    return {compact[index : index + size] for index in range(len(compact) - size + 1)}


def jaccard(left: set[str], right: set[str]) -> float:
    union = left | right
    return len(left & right) / len(union) if union else 1.0


def pattern_similarity(patterns: list[str]) -> dict[str, object]:
    unique_patterns = sorted(set(patterns))
    shingle_sets = [(pattern, shingles(pattern)) for pattern in unique_patterns]
    max_score = 0.0
    max_pair: tuple[str, str] = ("", "")
    high_pairs = 0

    for i, (left_pattern, left_set) in enumerate(shingle_sets):
        for right_pattern, right_set in shingle_sets[i + 1 :]:
            score = jaccard(left_set, right_set)
            if score >= 0.8:
                high_pairs += 1
            if score > max_score:
                max_score = score
                max_pair = (left_pattern, right_pattern)

    pair_count = len(unique_patterns) * (len(unique_patterns) - 1) // 2
    return {
        "unique_pattern_count": len(unique_patterns),
        "pattern_pair_count": pair_count,
        "max_pattern_jaccard": round(max_score, 4),
        "max_pattern_pair": max_pair,
        "high_similarity_pattern_pairs_0_8": high_pairs,
        "high_similarity_pattern_pair_ratio_0_8": round(high_pairs / pair_count, 6) if pair_count else 0,
    }


def main() -> None:
    files = html_files(SITE_DIR)
    records: list[dict[str, object]] = []
    missing_title: list[str] = []
    missing_h1: list[str] = []

    for path in files:
        text = path.read_text(encoding="utf-8")
        title_match = TITLE_RE.search(text)
        h1_match = H1_RE.search(text)
        relative = str(path.relative_to(ROOT)).replace("\\", "/")
        if not title_match:
            missing_title.append(relative)
            continue
        if not h1_match:
            missing_h1.append(relative)
            continue
        title = clean_text(title_match.group(1))
        keyword = clean_text(h1_match.group(1))
        records.append(
            {
                "path": relative,
                "title": title,
                "keyword": keyword,
                "length": len(title),
                "starts_with_keyword": title.startswith(keyword),
                "pattern": title[len(keyword) :].strip() if title.startswith(keyword) else title,
            }
        )

    titles = [str(record["title"]) for record in records]
    lengths = [int(record["length"]) for record in records]
    title_counts = Counter(titles)
    pattern_counts = Counter(str(record["pattern"]) for record in records)
    duplicate_titles = {title: count for title, count in title_counts.items() if count > 1}
    top_pattern_count = pattern_counts.most_common(1)[0][1] if pattern_counts else 0
    similarity = pattern_similarity([str(record["pattern"]) for record in records])

    checks = {
        "title_duplicates_zero": len(duplicate_titles) == 0,
        "all_titles_start_with_h1_keyword": all(bool(record["starts_with_keyword"]) for record in records),
        "all_titles_35_to_60": all(35 <= length <= 60 for length in lengths),
        "missing_title_zero": not missing_title,
        "missing_h1_zero": not missing_h1,
        "top_pattern_ratio_below_1_percent": (top_pattern_count / len(records)) < 0.01 if records else False,
        "high_similarity_pattern_ratio_below_5_percent": similarity["high_similarity_pattern_pair_ratio_0_8"] < 0.05,
    }
    status = "PASS" if all(checks.values()) else "FAIL"

    report = {
        "status": status,
        "checks": checks,
        "html_files": len(files),
        "processed": len(records),
        "missing_title_count": len(missing_title),
        "missing_h1_count": len(missing_h1),
        "unique_titles": len(title_counts),
        "duplicate_title_count": sum(count - 1 for count in title_counts.values() if count > 1),
        "title_average_length": round(sum(lengths) / len(lengths), 2) if lengths else 0,
        "title_min_length": min(lengths, default=0),
        "title_max_length": max(lengths, default=0),
        "title_below_35": sum(length < 35 for length in lengths),
        "title_above_60": sum(length > 60 for length in lengths),
        "keyword_prefix_mismatch_count": sum(not record["starts_with_keyword"] for record in records),
        "unique_patterns": len(pattern_counts),
        "top_pattern_count": top_pattern_count,
        "top_pattern_ratio": round(top_pattern_count / len(records), 6) if records else 0,
        "similarity": similarity,
        "top_patterns": pattern_counts.most_common(20),
        "duplicate_title_samples": list(duplicate_titles.items())[:20],
        "sample": records[:20],
    }
    REPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
    REPORT_PATH.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
