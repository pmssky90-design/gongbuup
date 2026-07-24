from __future__ import annotations

import argparse
import csv
import hashlib
import re
from collections import defaultdict
from pathlib import Path
from datetime import datetime


ROOT = Path(__file__).resolve().parents[1]
REGION_SOURCE = ROOT / "data" / "gongbuup_region_master.csv"
SCHOOL_SOURCE = ROOT / "data" / "gongbuup_school_master.csv"
DEFAULT_OUTPUT = ROOT / "data" / "quality_sample_200.csv"
SUBJECTS = ["국영수과외", "수학과외", "영어과외"]
TARGETS = ["초등학생", "중학생", "고등학생"]
FIELDS = [
    "page_type", "시도", "시군구", "읍면동", "리", "지역명", "학교명", "학교급",
    "대상", "과목", "메인키워드", "슬러그", "제목", "설명", "본문", "source_html",
]


def digest(value: str) -> int:
    return int(hashlib.sha256(value.encode("utf-8")).hexdigest(), 16)


def choose(values: list[str], identity: str, kind: str) -> str:
    return values[digest(f"{identity}:{kind}") % len(values)]


def balanced_sample(rows: list[dict[str, str]], count: int, group_field: str) -> list[dict[str, str]]:
    groups: dict[str, list[dict[str, str]]] = defaultdict(list)
    for row in rows:
        groups[row[group_field]].append(row)
    for values in groups.values():
        values.sort(key=lambda row: digest(row["source_html"]))
    keys = sorted(groups)
    selected: list[dict[str, str]] = []
    while len(selected) < count and keys:
        remaining: list[str] = []
        for key in keys:
            if groups[key] and len(selected) < count:
                selected.append(groups[key].pop(0))
            if groups[key]:
                remaining.append(key)
        keys = remaining
    if len(selected) != count:
        raise ValueError(f"요청 {count}건 중 {len(selected)}건만 선택했습니다.")
    return selected


def clean_slug(value: str) -> str:
    return re.sub(r"\s+", "", value)


def valid_region_name(value: str) -> bool:
    return bool(value) and not re.search(r"(수학|영어|국영수|초등|중등|고등)$", value)


def main() -> None:
    parser = argparse.ArgumentParser(description="전국 분산 품질검수 입력 200건 선택")
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT.relative_to(ROOT)))
    args = parser.parse_args()
    output = Path(args.output)
    if not output.is_absolute():
        output = ROOT / output
    if output.exists():
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output = output.with_name(f"{output.stem}_{timestamp}{output.suffix}")
    with REGION_SOURCE.open("r", encoding="utf-8-sig", newline="") as handle:
        regions = [
            row for row in csv.DictReader(handle)
            if row["is_valid"].lower() == "true" and valid_region_name(row["지역표시명"])
        ]
    with SCHOOL_SOURCE.open("r", encoding="utf-8-sig", newline="") as handle:
        schools = [row for row in csv.DictReader(handle) if row["is_valid"].lower() == "true"]

    selected_regions = balanced_sample(regions, 100, "시도")
    middle = balanced_sample([row for row in schools if row["학교급"] == "중학교"], 50, "시도")
    high = balanced_sample([row for row in schools if row["학교급"] == "고등학교"], 50, "시도")

    output_rows: list[dict[str, str]] = []
    used_slugs: set[str] = set()
    for row in selected_regions:
        name = row["지역표시명"]
        identity = "|".join(row.get(field, "") for field in ("시도", "시군구", "읍면동", "리", "지역표시명"))
        target = choose(TARGETS, identity, "target")
        subject = choose(SUBJECTS, identity, "subject")
        keyword = f"{name} {target} {subject}"
        slug = clean_slug(keyword)
        if slug in used_slugs:
            slug = clean_slug(f"{row['시도']} {row['시군구']} {keyword}")
        used_slugs.add(slug)
        output_rows.append({
            "page_type": "region", "시도": row["시도"], "시군구": row["시군구"],
            "읍면동": row["읍면동"], "리": row["리"], "지역명": name,
            "학교명": "", "학교급": "", "대상": target, "과목": subject,
            "메인키워드": keyword, "슬러그": slug, "제목": "", "설명": "", "본문": "",
            "source_html": row["source_html"],
        })

    for row in middle + high:
        school = row["학교명"]
        identity = "|".join((row["시도"], row["시군구"], school))
        subject = choose(SUBJECTS, identity, "subject")
        target = "중학생" if row["학교급"] == "중학교" else "고등학생"
        keyword = f"{school} {target} {subject}"
        slug = clean_slug(keyword)
        if slug in used_slugs:
            slug = clean_slug(f"{row['시도']} {row['시군구']} {keyword}")
        used_slugs.add(slug)
        output_rows.append({
            "page_type": "school", "시도": row["시도"], "시군구": row["시군구"],
            "읍면동": row["읍면동"], "리": row["리"], "지역명": school,
            "학교명": school, "학교급": row["학교급"], "대상": target, "과목": subject,
            "메인키워드": keyword, "슬러그": slug, "제목": "", "설명": "", "본문": "",
            "source_html": row["source_html"],
        })

    with output.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=FIELDS)
        writer.writeheader()
        writer.writerows(output_rows)
    print(f"OUTPUT={output}")
    print(f"지역 {len(selected_regions)}, 중학교 {len(middle)}, 고등학교 {len(high)}, 합계 {len(output_rows)}")


if __name__ == "__main__":
    main()
