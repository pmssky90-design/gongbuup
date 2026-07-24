from __future__ import annotations

import csv
import html
import json
import re
import sys
from collections import Counter, defaultdict
from datetime import datetime
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import quote

from openpyxl import Workbook


ROOT = Path(__file__).resolve().parents[1]
SOURCE_ROOT = Path(r"C:\Users\user\Documents\GitHub\studynavi")
DATA_DIR = ROOT / "data"
REPORT_DIR = ROOT / "reports"

MASTER_FIELDS = [
    "source_html", "source_relative_path", "source_url", "page_type",
    "시도", "시군구", "읍면동", "리", "지역표시명", "학교명", "학교급",
    "title", "h1", "canonical", "description", "slug", "confidence",
    "is_valid", "validation_note",
]
REGION_FIELDS = [
    "page_type", "시도", "시군구", "읍면동", "리", "지역표시명",
    "source_html", "confidence", "is_valid", "validation_note",
]
SCHOOL_FIELDS = [
    "page_type", "시도", "시군구", "읍면동", "리", "학교명", "학교급",
    "대상", "source_html", "confidence", "is_valid", "validation_note",
]
EXCLUDED_FIELDS = MASTER_FIELDS + ["exclusion_reason"]

PROVINCES = {
    "서울": "서울특별시", "부산": "부산광역시", "대구": "대구광역시",
    "인천": "인천광역시", "광주": "광주광역시", "대전": "대전광역시",
    "울산": "울산광역시", "세종": "세종특별자치시", "경기": "경기도",
    "강원": "강원특별자치도", "충북": "충청북도", "충남": "충청남도",
    "전북": "전북특별자치도", "전남": "전라남도", "경북": "경상북도",
    "경남": "경상남도", "제주": "제주특별자치도",
}
TARGET_BY_GRADE = {
    "초등학교": "초등학생", "중학교": "중학생", "고등학교": "고등학생",
}
ABNORMAL_RE = re.compile(r"[\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff\ufffd]")
TAG_RE = re.compile(r"<[^>]+>")
DERIVED_RE = re.compile(r"(수학|영어|초등|중등|고등).*(과외)$")


class MetadataParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.in_title = False
        self.in_h1 = False
        self.in_anchor = False
        self.title_parts: list[str] = []
        self.h1_parts: list[str] = []
        self.anchor_parts: list[str] = []
        self.anchor_texts: list[str] = []
        self.canonical = ""
        self.description = ""

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        values = {key.lower(): (value or "") for key, value in attrs}
        lowered = tag.lower()
        if lowered == "title":
            self.in_title = True
        elif lowered == "h1" and not self.h1_parts:
            self.in_h1 = True
        elif lowered == "a":
            self.in_anchor = True
            self.anchor_parts = []
        elif lowered == "link" and "canonical" in values.get("rel", "").lower():
            self.canonical = values.get("href", "").strip()
        elif lowered == "meta" and values.get("name", "").lower() == "description":
            self.description = values.get("content", "").strip()

    def handle_endtag(self, tag: str) -> None:
        lowered = tag.lower()
        if lowered == "title":
            self.in_title = False
        elif lowered == "h1":
            self.in_h1 = False
        elif lowered == "a":
            self.in_anchor = False
            text = clean_text(" ".join(self.anchor_parts))
            if text and len(self.anchor_texts) < 100:
                self.anchor_texts.append(text)

    def handle_data(self, data: str) -> None:
        if self.in_title:
            self.title_parts.append(data)
        if self.in_h1:
            self.h1_parts.append(data)
        if self.in_anchor:
            self.anchor_parts.append(data)

    @property
    def title(self) -> str:
        return clean_text(" ".join(self.title_parts))

    @property
    def h1(self) -> str:
        return clean_text(" ".join(self.h1_parts))


def clean_text(value: str) -> str:
    return re.sub(r"\s+", " ", html.unescape(value or "")).strip()


def strip_tutor(value: str) -> str:
    value = clean_text(value)
    return value[:-2] if value.endswith("과외") else value


def normalize_locality(value: str, province_short: str) -> str:
    result = strip_tutor(value)
    if result.startswith(province_short):
        candidate = result[len(province_short):]
        if candidate.endswith(("동", "읍", "면", "리")) and len(candidate) >= 2:
            return candidate
    return result


def school_identity(folder: str) -> tuple[str, str]:
    raw = strip_tutor(folder)
    for suffix, grade in (
        ("초등학교", "초등학교"), ("중학교", "중학교"), ("고등학교", "고등학교"),
    ):
        if raw.endswith(suffix):
            return raw, grade
    if raw.endswith("초") and len(raw) > 1:
        return raw + "등학교", "초등학교"
    if raw.endswith("중") and len(raw) > 1:
        return raw + "학교", "중학교"
    if raw.endswith("고") and len(raw) > 1:
        return raw + "등학교", "고등학교"
    return raw, ""


def read_html(path: Path) -> tuple[str, str]:
    data = path.read_bytes()
    for encoding in ("utf-8-sig", "cp949"):
        try:
            return data.decode(encoding), encoding
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace"), "utf-8-replace"


def quoted_relative_url(relative: Path) -> str:
    parts = [quote(part, safe="") for part in relative.parent.parts]
    return "/" + "/".join(parts) + ("/" if parts else "")


def classify(relative: Path, parser: MetadataParser, raw_text: str) -> dict[str, object]:
    directories = list(relative.parent.parts)
    title = parser.title
    h1 = parser.h1
    canonical = parser.canonical
    description = parser.description
    slug = directories[-1] if directories else ""
    page_type = "other"
    confidence = "low"
    valid = False
    notes: list[str] = []
    province = district = locality = ri = display = school = grade = ""

    province_folder = directories[0] if directories else ""
    province_short = strip_tutor(province_folder)
    recognized_province = province_folder.endswith("과외") and province_short in PROVINCES
    if recognized_province:
        province = PROVINCES[province_short]
        if len(directories) >= 2:
            district = strip_tutor(directories[1])
        if len(directories) >= 3:
            local_value = normalize_locality(directories[2], province_short)
            if local_value.endswith("리"):
                ri = local_value
            else:
                locality = local_value

    if recognized_province and 1 <= len(directories) <= 3:
        page_type = "region"
        display = province if len(directories) == 1 else district if len(directories) == 2 else (ri or locality)
        path_label = strip_tutor(directories[-1])
        evidence = sum(
            bool(value and path_label in strip_tutor(value))
            for value in (h1, title, canonical)
        )
        confidence = "high" if evidence >= 2 else "medium" if evidence == 1 else "low"
        valid = bool(display) and confidence != "low"
        if not display:
            notes.append("지역명 누락")
        if confidence == "low":
            notes.append("경로와 메타데이터의 지역명 일치 근거 부족")
    elif recognized_province and len(directories) == 4:
        leaf = directories[-1]
        raw_school, normalized_grade = school_identity(leaf)
        is_school_shape = (
            normalized_grade
            or raw_school.endswith("학교")
        ) and not DERIVED_RE.search(leaf)
        if is_school_shape:
            page_type = "school"
            school = raw_school
            grade = normalized_grade
            leaf_label = strip_tutor(leaf)
            evidence = sum(
                bool(value and leaf_label in strip_tutor(value))
                for value in (h1, title, canonical)
            )
            confidence = "high" if evidence >= 2 and grade else "medium" if evidence >= 1 else "low"
            valid = bool(school and grade and (district or locality or ri)) and confidence != "low"
            if not grade:
                notes.append("학교급 누락 또는 대상 외 학교")
            if not (district or locality or ri):
                notes.append("학교 지역 매핑 누락")
            if confidence == "low":
                notes.append("학교명 판별 근거 부족")

    if page_type == "other":
        if not directories:
            notes.append("홈 페이지")
        elif any(term in slug for term in ("개인정보", "이용안내", "연락처", "오류")):
            notes.append("안내 또는 제외 페이지")
        elif recognized_province:
            notes.append("과목·학년 파생 페이지 또는 비대상 계층")
        else:
            notes.append("지역·학교와 무관한 일반 콘텐츠")

    combined = " ".join((title, h1, canonical, description))
    if ABNORMAL_RE.search(combined) or "\ufffd" in combined:
        notes.append("비정상 문자 또는 한글 깨짐")
        valid = False
    if TAG_RE.search(title) or TAG_RE.search(h1) or TAG_RE.search(description):
        notes.append("추출 텍스트에 HTML 태그 포함")
    if page_type in ("region", "school") and not recognized_province:
        notes.append("시도 계층 불일치")
        valid = False

    return {
        "source_html": str(SOURCE_ROOT / relative),
        "source_relative_path": relative.as_posix(),
        "source_url": canonical or quoted_relative_url(relative),
        "page_type": page_type,
        "시도": province,
        "시군구": district,
        "읍면동": locality,
        "리": ri,
        "지역표시명": display,
        "학교명": school,
        "학교급": grade,
        "title": title,
        "h1": h1,
        "canonical": canonical,
        "description": description,
        "slug": slug,
        "confidence": confidence,
        "is_valid": "true" if valid else "false",
        "validation_note": "; ".join(dict.fromkeys(notes)) or "정상",
    }


def write_csv(path: Path, fields: list[str], rows: list[dict[str, object]]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def write_xlsx(path: Path, fields: list[str], rows: list[dict[str, object]]) -> None:
    workbook = Workbook(write_only=True)
    sheet = workbook.create_sheet("studynavi_master")
    sheet.append(fields)
    for row in rows:
        sheet.append([row.get(field, "") for field in fields])
    workbook.save(path)


def round_robin_sample(rows: list[dict[str, object]], count: int, keys: tuple[str, ...]) -> list[dict[str, object]]:
    groups: dict[tuple[str, ...], list[dict[str, object]]] = defaultdict(list)
    for row in sorted(rows, key=lambda item: (item.get("confidence") != "high", str(item.get("source_html", "")))):
        groups[tuple(str(row.get(key, "")) for key in keys)].append(row)
    result: list[dict[str, object]] = []
    group_keys = sorted(groups)
    while len(result) < count and group_keys:
        remaining: list[tuple[str, ...]] = []
        for key in group_keys:
            if groups[key] and len(result) < count:
                result.append(groups[key].pop(0))
            if groups[key]:
                remaining.append(key)
        group_keys = remaining
    return result


def main() -> int:
    if not SOURCE_ROOT.is_dir():
        print(f"STUDYNAVI 경로를 찾을 수 없습니다: {SOURCE_ROOT}", file=sys.stderr)
        return 1
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    REPORT_DIR.mkdir(parents=True, exist_ok=True)

    html_files = sorted(SOURCE_ROOT.rglob("*.html"), key=lambda path: path.as_posix().casefold())
    master_rows: list[dict[str, object]] = []
    excluded_rows: list[dict[str, object]] = []
    parse_errors: list[dict[str, str]] = []
    encoding_counts: Counter[str] = Counter()
    title_missing = h1_missing = canonical_missing = abnormal_count = 0

    for index, path in enumerate(html_files, start=1):
        relative = path.relative_to(SOURCE_ROOT)
        try:
            raw_text, encoding = read_html(path)
            encoding_counts[encoding] += 1
            parser = MetadataParser()
            parser.feed(raw_text)
            parser.close()
            row = classify(relative, parser, raw_text)
        except Exception as exc:
            row = {field: "" for field in MASTER_FIELDS}
            row.update({
                "source_html": str(path),
                "source_relative_path": relative.as_posix(),
                "source_url": quoted_relative_url(relative),
                "page_type": "other",
                "confidence": "low",
                "is_valid": "false",
                "validation_note": f"HTML 파싱 실패: {exc}",
            })
            parse_errors.append({"file": str(path), "error": str(exc)})
        master_rows.append(row)
        title_missing += int(not row["title"])
        h1_missing += int(not row["h1"])
        canonical_missing += int(not row["canonical"])
        abnormal_count += int("비정상 문자" in str(row["validation_note"]))
        if row["page_type"] == "other" or row["is_valid"] != "true" or row["confidence"] == "low":
            excluded = dict(row)
            excluded["exclusion_reason"] = row["validation_note"]
            excluded_rows.append(excluded)
        if index % 1000 == 0 or index == len(html_files):
            print(f"{index} / {len(html_files)}")

    region_candidates = [
        row for row in master_rows if row["page_type"] == "region" and row["is_valid"] == "true"
    ]
    school_candidates = [
        row for row in master_rows if row["page_type"] == "school" and row["is_valid"] == "true"
    ]

    region_unique: dict[tuple[str, ...], dict[str, object]] = {}
    for row in region_candidates:
        key = tuple(str(row[field]) for field in ("시도", "시군구", "읍면동", "리", "지역표시명"))
        region_unique.setdefault(key, row)
    school_unique: dict[tuple[str, ...], dict[str, object]] = {}
    for row in school_candidates:
        key = tuple(str(row[field]) for field in ("학교명", "시도", "시군구"))
        school_unique.setdefault(key, row)

    region_rows = [
        {**row, "page_type": "region"} for row in region_unique.values()
    ]
    school_rows = [
        {**row, "page_type": "school", "대상": TARGET_BY_GRADE.get(str(row["학교급"]), "")}
        for row in school_unique.values()
    ]

    canonical_counts = Counter(str(row["canonical"]) for row in master_rows if row["canonical"])
    canonical_duplicate_count = sum(count - 1 for count in canonical_counts.values() if count > 1)
    confidence_counts = Counter(str(row["confidence"]) for row in master_rows)
    type_counts = Counter(str(row["page_type"]) for row in master_rows)
    regions_by_province = Counter(str(row["시도"]) for row in region_rows)
    schools_by_province = Counter(str(row["시도"]) for row in school_rows)
    schools_by_grade = Counter(str(row["학교급"]) for row in school_rows)

    region_expected = len(region_rows) * 3 * 3
    school_expected = len(school_rows) * 3
    review_rows = (
        round_robin_sample(region_rows, 100, ("시도",))
        + round_robin_sample(school_rows, 100, ("시도", "학교급"))
    )

    output_files = [
        DATA_DIR / "studynavi_extracted_master.csv",
        DATA_DIR / "studynavi_extracted_master.xlsx",
        DATA_DIR / "gongbuup_region_master.csv",
        DATA_DIR / "gongbuup_school_master.csv",
        DATA_DIR / "studynavi_extracted_excluded.csv",
        DATA_DIR / "studynavi_review_sample_200.csv",
        REPORT_DIR / "studynavi_html_extraction_report.txt",
        REPORT_DIR / "studynavi_html_extraction_report.json",
    ]
    existing = [str(path) for path in output_files if path.exists()]
    if existing:
        print("기존 출력 파일을 덮어쓰지 않기 위해 중단합니다:\n" + "\n".join(existing), file=sys.stderr)
        return 1

    write_csv(output_files[0], MASTER_FIELDS, master_rows)
    write_xlsx(output_files[1], MASTER_FIELDS, master_rows)
    write_csv(output_files[2], REGION_FIELDS, region_rows)
    write_csv(output_files[3], SCHOOL_FIELDS, school_rows)
    write_csv(output_files[4], EXCLUDED_FIELDS, excluded_rows)
    write_csv(output_files[5], MASTER_FIELDS + ["대상"], review_rows)

    report = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "source_root": str(SOURCE_ROOT),
        "total_html_count": len(html_files),
        "parse_success_count": len(html_files) - len(parse_errors),
        "parse_failure_count": len(parse_errors),
        "page_type_count": dict(type_counts),
        "valid_region_count": len(region_rows),
        "valid_school_count": len(school_rows),
        "school_grade_count": dict(schools_by_grade),
        "confidence_count": dict(confidence_counts),
        "region_duplicate_removed_count": len(region_candidates) - len(region_rows),
        "school_duplicate_removed_count": len(school_candidates) - len(school_rows),
        "canonical_duplicate_count": canonical_duplicate_count,
        "title_missing_count": title_missing,
        "h1_missing_count": h1_missing,
        "canonical_missing_count": canonical_missing,
        "school_grade_missing_count": sum(
            1 for row in master_rows if row["page_type"] == "school" and not row["학교급"]
        ),
        "region_name_missing_count": sum(
            1 for row in master_rows if row["page_type"] == "region" and not row["지역표시명"]
        ),
        "abnormal_character_count": abnormal_count,
        "regions_by_province": dict(sorted(regions_by_province.items())),
        "schools_by_province": dict(sorted(schools_by_province.items())),
        "region_expected_page_count": region_expected,
        "school_expected_page_count": school_expected,
        "total_expected_page_count": region_expected + school_expected,
        "encoding_count": dict(encoding_counts),
        "url_analysis": {
            "region_pattern": "/{시도}과외/{시군구}과외/{읍면동 또는 리}과외/",
            "school_pattern": "/{시도}과외/{시군구}과외/{읍면동 또는 리}과외/{학교명}과외/",
            "subject_pattern": ".../{지역명}{수학|영어}과외/",
            "grade_pattern": ".../{지역명}{초등|중등|고등}{수학|영어}과외/",
            "uses_korean_urls": True,
            "index_structure": "각 URL 폴더에 index.html을 두는 디렉터리 인덱스 구조",
            "derived_page_count": type_counts.get("other", 0),
            "duplicate_slug_type": "동일 지역·학교에서 과목 및 학년 접미사가 붙는 파생 슬러그",
        },
        "extraction_rules": [
            "URL 폴더 계층을 최우선으로 사용",
            "h1, title, canonical 일치 수로 confidence 산정",
            "17개 시도 최상위 폴더만 지역 계층으로 인정",
            "시도/시군구/읍면동 다음 직접 하위 폴더만 학교 후보로 판정",
            "초·중·고 축약 학교명을 초등학교·중학교·고등학교로 정규화",
            "과목·학년 파생 페이지는 other로 분리",
            "Python 표준 HTMLParser로 파일을 순차 처리",
        ],
        "parse_errors": parse_errors,
        "generated_files": [str(path.relative_to(ROOT)) for path in output_files],
        "modified_files": ["scripts/extract_studynavi.py"],
    }
    output_files[7].write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    text_lines = [
        "STUDYNAVI HTML 추출 보고서",
        f"생성 시각: {report['generated_at']}",
        f"원본 경로: {SOURCE_ROOT}",
        f"전체 HTML 수: {len(html_files)}",
        f"파싱 성공 수: {report['parse_success_count']}",
        f"파싱 실패 수: {report['parse_failure_count']}",
        f"region 판별 수: {type_counts.get('region', 0)}",
        f"school 판별 수: {type_counts.get('school', 0)}",
        f"other 판별 수: {type_counts.get('other', 0)}",
        f"유효 지역 수: {len(region_rows)}",
        f"유효 학교 수: {len(school_rows)}",
        f"초등학교 수: {schools_by_grade.get('초등학교', 0)}",
        f"중학교 수: {schools_by_grade.get('중학교', 0)}",
        f"고등학교 수: {schools_by_grade.get('고등학교', 0)}",
        f"confidence high 수: {confidence_counts.get('high', 0)}",
        f"confidence medium 수: {confidence_counts.get('medium', 0)}",
        f"confidence low 수: {confidence_counts.get('low', 0)}",
        f"지역 중복 제거 수: {report['region_duplicate_removed_count']}",
        f"학교 중복 제거 수: {report['school_duplicate_removed_count']}",
        f"canonical 중복 수: {canonical_duplicate_count}",
        f"title 누락 수: {title_missing}",
        f"h1 누락 수: {h1_missing}",
        f"학교급 누락 수: {report['school_grade_missing_count']}",
        f"지역명 누락 수: {report['region_name_missing_count']}",
        f"비정상 문자 수: {abnormal_count}",
        f"지역 페이지 예상 조합 수: {region_expected}",
        f"학교 페이지 예상 조합 수: {school_expected}",
        f"전체 예상 페이지 수: {region_expected + school_expected}",
        "",
        "시도별 지역 수:",
        *[f"- {name}: {count}" for name, count in sorted(regions_by_province.items())],
        "",
        "시도별 학교 수:",
        *[f"- {name}: {count}" for name, count in sorted(schools_by_province.items())],
        "",
        "사용한 추출 규칙:",
        *[f"- {rule}" for rule in report["extraction_rules"]],
        "",
        "생성한 파일:",
        *[f"- {path}" for path in report["generated_files"]],
        "",
        "수정한 파일:",
        "- scripts/extract_studynavi.py",
    ]
    output_files[6].write_text("\n".join(text_lines) + "\n", encoding="utf-8")

    print(json.dumps({
        "total_html": len(html_files),
        "parse_success": report["parse_success_count"],
        "parse_failure": report["parse_failure_count"],
        "region": type_counts.get("region", 0),
        "school": type_counts.get("school", 0),
        "other": type_counts.get("other", 0),
        "valid_region": len(region_rows),
        "valid_school": len(school_rows),
        "review_sample": len(review_rows),
        "total_expected_pages": region_expected + school_expected,
    }, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
