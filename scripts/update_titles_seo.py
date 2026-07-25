from __future__ import annotations

import argparse
import hashlib
import html
import json
import os
import re
from collections import Counter
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SITE_DIR = ROOT / "site"
REPORT_PATH = ROOT / "reports" / "title_seo_update_report.json"

TITLE_RE = re.compile(r"(<title>)(.*?)(</title>)", re.IGNORECASE | re.DOTALL)
H1_RE = re.compile(r"<h1[^>]*>(.*?)</h1>", re.IGNORECASE | re.DOTALL)
TAG_RE = re.compile(r"<[^>]+>")

BASE_FOCI = [
    "학생에게 맞는 학습 방법",
    "내신 흐름을 잡는 공부 계획",
    "시험 전 점검할 학습 기준",
    "오답을 줄이는 복습 루틴",
    "개념부터 다지는 수업 방향",
    "성적 향상을 위한 학습 전략",
    "공부 습관을 세우는 접근",
    "단원별 약점을 찾는 점검법",
    "꾸준히 이어가는 학습 관리",
    "기초와 응용을 잇는 공부법",
    "학교 진도에 맞춘 복습 기준",
    "문제 풀이력을 키우는 학습법",
    "학습 공백을 줄이는 준비 과정",
    "목표별로 정리하는 공부 방향",
    "이해와 적용을 높이는 학습 설계",
    "수업 전후로 챙길 복습 포인트",
    "개인별 속도에 맞춘 학습 흐름",
    "실수를 줄이는 문제 접근법",
    "단기 목표를 세우는 공부 기준",
    "현재 수준을 살피는 학습 점검",
    "반복 학습을 이어가는 계획",
    "과목별 우선순위를 잡는 방법",
    "스스로 설명하는 복습 습관",
    "시험 범위를 나누는 준비법",
    "취약 단원을 보완하는 학습법",
    "학교 과제와 복습을 잇는 전략",
    "매일 실천할 공부 루틴",
    "학습 기록을 활용하는 방법",
    "질문을 정리하는 공부 방식",
    "수행평가까지 보는 준비 과정",
    "집중력을 유지하는 학습 흐름",
    "기본기를 확인하는 점검 기준",
    "목표 점수에 맞춘 복습 방향",
    "학년별 변화에 맞춘 공부법",
    "학습 태도를 다지는 관리법",
    "교과서 중심으로 잡는 준비법",
    "서술형 답안을 다듬는 학습법",
    "풀이 과정을 점검하는 공부법",
    "누적 복습을 설계하는 전략",
    "핵심 개념을 연결하는 학습법",
]

SUBJECT_FOCI = {
    "수학": [
        "수학 개념과 풀이 과정을 잡는 학습법",
        "수학 오답을 줄이는 단계별 복습",
        "수학 문제풀이 감각을 키우는 전략",
        "수학 내신 대비를 위한 개념 점검",
    ],
    "영어": [
        "영어 독해와 문법을 함께 다지는 학습법",
        "영어 어휘와 문장 구조를 잡는 공부법",
        "영어 내신 준비를 위한 복습 흐름",
        "영어 읽기와 쓰기를 잇는 학습 전략",
    ],
    "국어": [
        "국어 독해와 서술형 답안을 다듬는 학습법",
        "국어 지문 이해력을 키우는 공부법",
        "국어 내신 대비를 위한 작품 정리",
        "국어 어휘와 문장 흐름을 잡는 전략",
    ],
    "과학": [
        "과학 개념과 탐구 자료를 연결하는 학습법",
        "과학 단원별 원리를 정리하는 복습",
        "과학 내신 준비를 위한 실험 해석",
        "과학 용어와 개념 적용을 잡는 전략",
    ],
    "사회": [
        "사회 개념과 자료 해석을 잇는 학습법",
        "사회 단원 흐름을 정리하는 복습",
        "사회 내신 대비를 위한 핵심 점검",
        "사회 암기와 이해를 함께 잡는 전략",
    ],
    "영수": [
        "영어와 수학 균형을 맞추는 학습 전략",
        "영수 과목별 약점을 나누는 공부법",
        "영수 내신 준비를 함께 보는 복습",
        "영어 수학 기본기를 함께 다지는 흐름",
    ],
    "국영수": [
        "국영수 기본기를 균형 있게 다지는 학습법",
        "주요 과목 우선순위를 잡는 공부 전략",
        "국어 영어 수학 흐름을 정리하는 복습",
        "국영수 내신 대비를 나누어 보는 계획",
    ],
    "내신": [
        "내신 시험 범위에 맞춘 복습 전략",
        "내신 성적 향상을 위한 학습 계획",
        "내신 오답과 서술형을 함께 보는 공부법",
        "내신 대비를 차분히 정리하는 학습 흐름",
    ],
}

TARGET_FOCI = {
    "초등": [
        "초등학생 공부 습관을 세우는 학습법",
        "초등 기초를 차근차근 다지는 공부법",
        "초등 과정 이해를 높이는 복습 전략",
        "초등학생에게 맞춘 학습 루틴",
    ],
    "중등": [
        "중학생 내신 흐름을 잡는 학습법",
        "중등 개념과 시험 준비를 잇는 공부법",
        "중학생 학습 습관을 다지는 전략",
        "중등 과정 약점을 줄이는 복습 기준",
    ],
    "고등": [
        "고등학생 내신과 수능 기초를 잇는 전략",
        "고등 과정에 맞춘 학습 계획",
        "고등학생 시험 준비를 돕는 복습법",
        "고등 내신 성적 향상을 위한 공부법",
    ],
    "학교": [
        "학교 진도에 맞춰 살피는 학습 전략",
        "학교 시험 준비를 위한 복습 기준",
        "학교별 학습 흐름을 고려한 공부법",
        "학교 수업과 이어지는 학습 관리",
    ],
}

CONNECTORS = [
    " 중심의 성적 향상 전략",
    " 기준을 차분히 정리",
    " 흐름으로 공부 방향 잡기",
    " 관점의 시험 준비 포인트",
    " 중심으로 약점 보완하기",
    " 기준과 선택 포인트 정리",
    " 흐름에 맞춘 복습 방향",
    " 중심의 학습 습관 만들기",
    " 기준으로 내신 대비하기",
    " 관점의 오답 관리 방법",
    " 중심으로 기초 다지기",
    " 기준의 실전 점검 포인트",
    " 흐름에 맞춘 공부 계획",
    " 기준의 단계별 관리법",
    " 중심으로 이해력 높이기",
    " 관점의 목표별 준비법",
    " 흐름으로 꾸준함 만들기",
    " 기준에 맞춘 학습 전략",
    " 중심의 시험 범위 정리",
    " 관점으로 복습 루틴 잡기",
    " 흐름에 맞춘 오답 정리",
    " 기준으로 공부 습관 세우기",
    " 중심의 단원별 점검",
    " 관점의 내신 준비 전략",
    " 흐름으로 학습 공백 줄이기",
    " 기준의 수업 전후 복습",
    " 중심으로 문제 접근 다듬기",
    " 관점의 학습 기록 활용",
    " 흐름에 맞춘 성적 관리",
    " 기준으로 개념 연결하기",
    " 중심의 시험 전 점검",
    " 관점의 과목별 우선순위",
    " 흐름으로 공부 집중력 높이기",
    " 기준에 맞춘 수행평가 준비",
    " 중심으로 풀이 과정 확인",
    " 관점의 질문 정리 방법",
    " 흐름에 맞춘 기본기 점검",
    " 기준의 맞춤 복습 전략",
    " 중심으로 학습 속도 맞추기",
    " 관점의 학교 진도 복습",
    " 흐름으로 실수 줄이기",
    " 기준에 맞춘 주간 계획",
    " 중심의 자기주도 학습법",
    " 관점의 개념 적용 연습",
    " 흐름으로 목표 점수 준비",
    " 기준의 취약 단원 보완",
    " 중심으로 서술형 대비하기",
    " 관점의 반복 학습 설계",
    " 흐름에 맞춘 학습 관리",
    " 기준으로 실전 감각 키우기",
    " 중심의 교과서 복습",
    " 관점의 공부 계획 조정",
    " 흐름으로 이해 과정 점검",
    " 기준에 맞춘 단계별 복습",
]

SHORT_FOCI = [
    "학습 방법",
    "공부 전략",
    "내신 준비",
    "복습 기준",
    "오답 관리",
    "학습 습관",
    "시험 대비",
    "성적 향상",
    "개념 정리",
    "선택 기준",
    "공부 계획",
    "기초 학습",
]

SHORT_CONNECTORS = [
    "중심의 성적 향상 전략",
    "기준과 시험 준비 포인트",
    "흐름에 맞춘 복습 방향",
    "관점의 학습 습관 만들기",
    "중심으로 약점 보완하기",
    "기준의 오답 관리 방법",
    "관점의 선택 기준 정리",
    "흐름으로 공부 계획 세우기",
    "중심의 내신 대비 전략",
    "기준으로 기초 다지기",
    "관점의 실전 점검 기준",
    "흐름에 맞춘 목표별 준비법",
]


def stable_index(value: str, length: int, salt: str = "") -> int:
    digest = hashlib.sha256(f"{value}:{salt}".encode("utf-8")).hexdigest()
    return int(digest[:16], 16) % length


def clean_text(value: str) -> str:
    return re.sub(r"\s+", " ", html.unescape(TAG_RE.sub("", value))).strip()


def detect_focus(keyword: str) -> list[str]:
    candidates: list[str] = []
    for token, phrases in SUBJECT_FOCI.items():
        if token in keyword:
            candidates.extend(phrases)
    for token, phrases in TARGET_FOCI.items():
        if token in keyword:
            candidates.extend(phrases)
    if not candidates:
        candidates.extend(BASE_FOCI)
    candidates.extend(BASE_FOCI)
    return candidates


def ordered_candidates(keyword: str, path_key: str) -> list[str]:
    foci = detect_focus(keyword)
    offset = stable_index(path_key, len(foci), "focus")
    connector_offset = stable_index(path_key, len(CONNECTORS), "connector")
    short_offset = stable_index(path_key, len(SHORT_FOCI), "short")
    titles: list[str] = []

    for i in range(len(foci)):
        focus = foci[(offset + i * 7) % len(foci)]
        connector = CONNECTORS[(connector_offset + i * 5) % len(CONNECTORS)]
        titles.append(f"{keyword} {focus}{connector}")

    for i in range(len(SHORT_FOCI) * len(SHORT_CONNECTORS)):
        focus = SHORT_FOCI[(short_offset + i) % len(SHORT_FOCI)]
        connector = SHORT_CONNECTORS[(connector_offset + i * 3) % len(SHORT_CONNECTORS)]
        titles.append(f"{keyword} {focus} {connector}")

    return titles


def make_title(keyword: str, path_key: str, used: set[str]) -> str:
    if keyword == "공부업":
        title = "공부업 지역과 학교별 과외 탐색을 돕는 학습 정보와 선택 기준 안내"
        if title not in used:
            return title

    candidates = ordered_candidates(keyword, path_key)
    expanded: list[str] = []
    for title in candidates:
        expanded.append(title)
        if len(title) < 35:
            expanded.append(f"{title} 차근차근 정리")
            expanded.append(f"{title} 학습 방향 살펴보기")
    candidates = expanded
    valid = [title for title in candidates if 35 <= len(title) <= 60]
    fallback = valid or [title for title in candidates if len(title) <= 60] or candidates
    for title in fallback:
        if title not in used:
            return title

    # Rare duplicate fallback for identical H1 values. Keep it human-readable and deterministic.
    suffixes = [
        "맞춤 학습 흐름",
        "세부 공부 기준",
        "단계별 복습 설계",
        "시험 준비 흐름",
        "오답 점검 루틴",
        "내신 대비 방향",
    ]
    token = suffixes[stable_index(path_key, len(suffixes), "unique")]
    title = f"{keyword} {token}"
    if len(title) < 35:
        title = f"{title}과 성적 향상 전략"
    n = 2
    while title in used:
        title = f"{keyword} {token} {n}"
        n += 1
    return title


def html_files(root: Path) -> list[Path]:
    paths: list[Path] = []
    for current, dirs, files in os.walk(root):
        dirs.sort()
        for filename in sorted(files):
            if filename.lower().endswith(".html"):
                paths.append(Path(current) / filename)
    return paths


def update_titles(dry_run: bool) -> dict[str, object]:
    files = html_files(SITE_DIR)
    used: set[str] = set()
    records: list[dict[str, object]] = []
    missing_title = 0
    missing_h1 = 0
    changed = 0

    for path in files:
        text = path.read_text(encoding="utf-8")
        title_match = TITLE_RE.search(text)
        h1_match = H1_RE.search(text)
        if not title_match:
            missing_title += 1
            continue
        if not h1_match:
            missing_h1 += 1
            continue

        keyword = clean_text(h1_match.group(1))
        old_title = clean_text(title_match.group(2))
        new_title = make_title(keyword, str(path.relative_to(SITE_DIR)).replace("\\", "/"), used)
        used.add(new_title)

        if old_title != new_title:
            changed += 1
            if not dry_run:
                escaped = html.escape(new_title, quote=False)
                text = TITLE_RE.sub(rf"\1{escaped}\3", text, count=1)
                path.write_text(text, encoding="utf-8", newline="")

        records.append(
            {
                "path": str(path.relative_to(ROOT)).replace("\\", "/"),
                "keyword": keyword,
                "old_title": old_title,
                "new_title": new_title,
                "length": len(new_title),
                "pattern": new_title.removeprefix(keyword).strip(),
            }
        )

    title_counts = Counter(record["new_title"] for record in records)
    pattern_counts = Counter(record["pattern"] for record in records)
    lengths = [int(record["length"]) for record in records]
    report = {
        "status": "DRY_RUN" if dry_run else "UPDATED",
        "html_files": len(files),
        "processed": len(records),
        "changed": changed,
        "missing_title": missing_title,
        "missing_h1": missing_h1,
        "unique_titles": len(title_counts),
        "duplicate_title_count": sum(count - 1 for count in title_counts.values() if count > 1),
        "title_average_length": round(sum(lengths) / len(lengths), 2) if lengths else 0,
        "title_min_length": min(lengths, default=0),
        "title_max_length": max(lengths, default=0),
        "title_below_35": sum(length < 35 for length in lengths),
        "title_above_60": sum(length > 60 for length in lengths),
        "unique_patterns": len(pattern_counts),
        "top_pattern_ratio": round((pattern_counts.most_common(1)[0][1] / len(records)), 6) if records else 0,
        "top_patterns": pattern_counts.most_common(20),
        "sample": records[:20],
    }
    REPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
    REPORT_PATH.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return report


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    report = update_titles(args.dry_run)
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
