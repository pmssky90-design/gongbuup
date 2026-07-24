from __future__ import annotations

import argparse
import csv
import hashlib
import html
import json
import re
import shutil
import sys
import time
import unicodedata
from collections import Counter
from datetime import datetime
from pathlib import Path
from urllib.parse import quote, unquote, urljoin, urlsplit
from xml.sax.saxutils import escape as xml_escape

from content_generator import (
    SUBJECT_EXPRESSIONS,
    SUBJECT_MAP,
    TARGET_EXPRESSIONS,
    generate_page_content,
    load_pools,
    validate_content,
)
from keyword_combination_engine import (
    TARGET_BY_EXPRESSION,
    load_json as load_keyword_pool,
    make_content as make_keyword_content,
)


ROOT = Path(__file__).resolve().parents[1]
SETTINGS_PATH = ROOT / "config" / "settings.json"
DATA_PATH = ROOT / "data" / "sample_pages.csv"
PAGE_TEMPLATE_PATH = ROOT / "templates" / "page.html"
HOME_TEMPLATE_PATH = ROOT / "templates" / "home.html"
REPORTS_DIR = ROOT / "reports"
REQUIRED_FIELDS = ["지역명", "대상", "과목", "메인키워드", "슬러그"]
FULL_COMBINATION_FIELDS = {
    "page_type", "지역명", "학교명", "학교급", "학년표현", "과목표현",
    "내신사용", "display_keyword", "slug", "title", "description",
    "title_core_terms", "content_seed", "is_valid",
}
SUPPORTED_IMAGE_EXTENSIONS = {".webp", ".png", ".jpg", ".jpeg"}
BROKEN_MARKERS = ("\ufffd", "ì", "ë", "ð", "�")
CJK_FOREIGN_RE = re.compile(r"[\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff]")


def render(template: str, values: dict[str, object]) -> str:
    result = template
    for key, value in values.items():
        result = result.replace("{{" + key + "}}", str(value))
    unresolved = sorted(set(re.findall(r"\{\{([a-zA-Z0-9_]+)\}\}", result)))
    if unresolved:
        raise ValueError(f"치환되지 않은 템플릿 값: {', '.join(unresolved)}")
    return result


def clean_slug(value: str) -> str:
    normalized = unicodedata.normalize("NFC", value.strip())
    normalized = re.sub(r"\s+", "", normalized)
    normalized = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", normalized)
    normalized = normalized.strip(". ")
    return normalized


def site_join(site_url: str, path: str) -> str:
    base = site_url.rstrip("/") + "/"
    return urljoin(base, path.lstrip("/"))


def url_path(category: str, slug: str) -> str:
    return "/" + "/".join(quote(part, safe="") for part in (category, slug)) + "/"


def choose_output_dir(configured_name: str) -> Path:
    preferred = ROOT / configured_name
    if not preferred.exists():
        return preferred
    try:
        has_content = any(preferred.iterdir())
    except OSError:
        has_content = True
    if not has_content:
        return preferred
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    candidate = ROOT / f"{configured_name}_{timestamp}"
    suffix = 2
    while candidate.exists():
        candidate = ROOT / f"{configured_name}_{timestamp}_{suffix}"
        suffix += 1
    return candidate


def choose_report_file(filename: str) -> Path:
    preferred = REPORTS_DIR / filename
    if not preferred.exists():
        return preferred
    stem = preferred.stem
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    candidate = REPORTS_DIR / f"{stem}_{timestamp}{preferred.suffix}"
    suffix = 2
    while candidate.exists():
        candidate = REPORTS_DIR / f"{stem}_{timestamp}_{suffix}{preferred.suffix}"
        suffix += 1
    return candidate


def prepare_production_output(settings: dict[str, object]) -> tuple[Path, dict[str, object]]:
    production = dict(settings.get("production") or {})
    if not production.get("enabled", False):
        raise ValueError("settings.json의 production.enabled가 true여야 합니다.")
    output_name = str(production.get("output_directory", "")).strip()
    if output_name != "site":
        raise ValueError("Production 출력 폴더는 프로젝트 루트의 site만 허용됩니다.")
    output_dir = ROOT / "site"
    if output_dir.resolve(strict=False).parent != ROOT.resolve():
        raise ValueError("Production 출력 경로가 프로젝트 루트를 벗어났습니다.")
    if output_dir.is_symlink():
        raise ValueError("site가 심볼릭 링크이므로 안전을 위해 삭제하지 않습니다.")
    if output_dir.exists():
        if not production.get("delete_output_before_build", True):
            raise ValueError(
                "site가 이미 존재합니다. Production Mode에서는 "
                "delete_output_before_build=true가 필요합니다."
            )
        shutil.rmtree(output_dir)
    output_dir.mkdir(parents=False, exist_ok=False)
    return output_dir, production


def local_output_target(output_dir: Path, url_value: str) -> Path | None:
    parsed = urlsplit(html.unescape(url_value))
    if parsed.scheme or parsed.netloc or url_value.startswith(("#", "mailto:", "tel:", "data:")):
        return None
    relative = unquote(parsed.path).lstrip("/")
    target = output_dir / Path(relative.replace("/", "\\"))
    if parsed.path.endswith("/") or target.is_dir():
        target = target / "index.html"
    return target


def verify_production_output(
    output_dir: Path,
    site_url: str,
    generated_page_count: int,
    expected_sitemap_count: int,
) -> dict[str, object]:
    html_files = sorted(output_dir.rglob("*.html"))
    broken_links: list[str] = []
    broken_images: list[str] = []
    canonical_errors: list[str] = []
    og_image_errors: list[str] = []
    missing_og_images: list[str] = []
    missing_titles: list[str] = []
    missing_descriptions: list[str] = []
    missing_bodies: list[str] = []
    slugs: set[str] = set()
    duplicate_slugs: list[str] = []

    canonical_re = re.compile(
        r'<link[^>]+rel=["\']canonical["\'][^>]+href=["\']([^"\']+)["\']',
        re.IGNORECASE,
    )
    og_image_re = re.compile(
        r'<meta[^>]+property=["\']og:image["\'][^>]+content=["\']([^"\']+)["\']',
        re.IGNORECASE,
    )
    title_re = re.compile(r"<title>(.*?)</title>", re.IGNORECASE | re.DOTALL)
    description_re = re.compile(
        r'<meta[^>]+name=["\']description["\'][^>]+content=["\']([^"\']*)["\']',
        re.IGNORECASE,
    )
    link_re = re.compile(
        r'<(?:a|link|script)[^>]+(?:href|src)=["\']([^"\']+)["\']',
        re.IGNORECASE,
    )
    image_re = re.compile(r'<img[^>]+src=["\']([^"\']+)["\']', re.IGNORECASE)
    main_re = re.compile(r"<main\b[^>]*>(.*?)</main>", re.IGNORECASE | re.DOTALL)

    for path in html_files:
        text = path.read_text(encoding="utf-8")
        relative = path.relative_to(output_dir)
        is_home = relative == Path("index.html")
        if not is_home:
            slug = path.parent.name
            if slug in slugs:
                duplicate_slugs.append(slug)
            slugs.add(slug)

        title_match = title_re.search(text)
        if not title_match or not re.sub(r"<[^>]+>", "", title_match.group(1)).strip():
            missing_titles.append(str(relative))
        description_match = description_re.search(text)
        if not description_match or not description_match.group(1).strip():
            missing_descriptions.append(str(relative))
        main_match = main_re.search(text)
        main_text = re.sub(r"<[^>]+>", " ", main_match.group(1)) if main_match else ""
        if not html.unescape(main_text).strip():
            missing_bodies.append(str(relative))

        expected_path = (
            "/"
            if is_home
            else "/" + "/".join(quote(part, safe="") for part in relative.parent.parts) + "/"
        )
        expected_canonical = site_join(site_url, expected_path)
        canonical_match = canonical_re.search(text)
        if not canonical_match or html.unescape(canonical_match.group(1)) != expected_canonical:
            canonical_errors.append(str(relative))

        og_match = og_image_re.search(text)
        if is_home:
            pass
        elif not og_match:
            missing_og_images.append(str(relative))
        else:
            parsed_og = urlsplit(html.unescape(og_match.group(1)))
            expected_host = urlsplit(site_url)
            if (parsed_og.scheme, parsed_og.netloc) != (
                expected_host.scheme,
                expected_host.netloc,
            ):
                og_image_errors.append(str(relative))
            else:
                og_target = output_dir / Path(unquote(parsed_og.path).lstrip("/").replace("/", "\\"))
                if not og_target.is_file():
                    og_image_errors.append(str(relative))

        for link in link_re.findall(text):
            target = local_output_target(output_dir, link)
            if target is not None and not target.exists():
                broken_links.append(f"{relative}: {html.unescape(link)}")
        for image_path in image_re.findall(text):
            target = local_output_target(output_dir, image_path)
            if target is not None and not target.is_file():
                broken_images.append(f"{relative}: {html.unescape(image_path)}")

    sitemap_file = output_dir / "sitemap.xml"
    sitemap_url_count = 0
    if sitemap_file.is_file():
        sitemap_url_count = len(
            re.findall(r"<loc>.*?</loc>", sitemap_file.read_text(encoding="utf-8"), re.DOTALL)
        )
    output_size = sum(path.stat().st_size for path in output_dir.rglob("*") if path.is_file())
    expected_html_count = generated_page_count + 1
    errors: list[str] = []
    warnings: list[str] = []
    checks = {
        "html_count_matches": len(html_files) == expected_html_count,
        "sitemap_count_matches": sitemap_url_count == expected_sitemap_count,
        "canonical_valid": not canonical_errors,
        "og_image_paths_valid": not og_image_errors,
        "links_valid": not broken_links,
        "image_paths_valid": not broken_images,
        "titles_present": not missing_titles,
        "descriptions_present": not missing_descriptions,
        "bodies_present": not missing_bodies,
        "slugs_unique": not duplicate_slugs,
    }
    for name, passed in checks.items():
        if not passed:
            errors.append(f"{name} 검증 실패")
    if missing_og_images:
        warnings.append(f"og:image 누락 페이지 {len(missing_og_images)}개")
    return {
        "html_count": len(html_files),
        "expected_html_count": expected_html_count,
        "sitemap_url_count": sitemap_url_count,
        "expected_sitemap_url_count": expected_sitemap_count,
        "output_size_bytes": output_size,
        "broken_link_count": len(broken_links),
        "broken_links": broken_links,
        "broken_image_count": len(broken_images),
        "broken_images": broken_images,
        "canonical_error_count": len(canonical_errors),
        "canonical_errors": canonical_errors,
        "og_image_error_count": len(og_image_errors),
        "og_image_errors": og_image_errors,
        "missing_og_image_count": len(missing_og_images),
        "missing_title_count": len(missing_titles),
        "missing_description_count": len(missing_descriptions),
        "missing_body_count": len(missing_bodies),
        "duplicate_slug_count": len(duplicate_slugs),
        "checks": checks,
        "errors": errors,
        "warnings": warnings,
    }


def fixed_representative_pages(
    page_records: list[dict[str, object]],
) -> list[dict[str, object]]:
    def choose(scope: str, limit: int) -> list[dict[str, object]]:
        candidates = sorted(
            (record for record in page_records if record["page_type"] == scope),
            key=lambda record: hashlib.sha256(
                f"production-report:{record['slug']}".encode("utf-8")
            ).hexdigest(),
        )
        selected: list[dict[str, object]] = []
        selected_slugs: set[str] = set()

        def add_first(predicate: object) -> None:
            for record in candidates:
                if record["slug"] in selected_slugs:
                    continue
                if predicate(record):
                    selected.append(record)
                    selected_slugs.add(str(record["slug"]))
                    return

        for subject in ("영어", "수학", "과학", "국어", "영수"):
            add_first(lambda record, value=subject: record["subject_expression"] == value)
        for grade in ("중등", "중학생", "고등", "고등학생"):
            add_first(lambda record, value=grade: record["grade_expression"] == value)
        add_first(lambda record: bool(record["internal_exam"]))
        for record in candidates:
            if len(selected) >= limit:
                break
            if record["slug"] not in selected_slugs:
                selected.append(record)
                selected_slugs.add(str(record["slug"]))
        return selected[:limit]

    fixed = choose("region", 18) + choose("school", 12)
    return [
        {
            "page_type": record["page_type"],
            "slug": record["slug"],
            "title": record["title"],
            "grade_expression": record["grade_expression"],
            "subject_expression": record["subject_expression"],
            "internal_exam": record["internal_exam"],
            "canonical": record["canonical"],
            "local_file": record["local_file"],
        }
        for record in fixed
    ]


def main(argv: list[str] | None = None) -> int:
    build_started_at = datetime.now()
    build_timer = time.perf_counter()
    argument_parser = argparse.ArgumentParser(description="공부업 정적 후보 사이트 생성")
    argument_parser.add_argument(
        "--data", default=str(DATA_PATH.relative_to(ROOT)),
        help="프로젝트 루트 기준 입력 CSV 경로",
    )
    argument_parser.add_argument(
        "--production", action="store_true",
        help="candidate_output 대신 프로젝트 루트의 site를 새로 생성",
    )
    args = argument_parser.parse_args(argv)
    data_path = Path(args.data)
    if not data_path.is_absolute():
        data_path = ROOT / data_path
    errors: list[str] = []
    warnings: list[str] = []
    created_files: list[str] = []
    failed_pages: list[str] = []
    missing_images: set[str] = set()
    location_mismatches = 0
    subject_mismatches = 0
    abnormal_count = 0
    missing_title = 0
    missing_description = 0
    missing_body = 0
    invalid_canonical = 0
    title_generated_count = 0
    description_generated_count = 0
    body_generated_count = 0
    grade_mismatch_count = 0
    content_subject_mismatch_count = 0
    other_region_count = 0
    forbidden_expression_count = 0
    content_abnormal_count = 0
    page_type_counts: Counter[str] = Counter()
    region_name_error_count = 0
    school_name_error_count = 0
    target_error_count = 0
    subject_error_count = 0
    title_core_body_missing_count = 0
    title_core_description_missing_count = 0
    elementary_exam_count = 0
    school_grade_expression_mismatch_count = 0
    production_config: dict[str, object] = {}
    full_combination_input = False

    try:
        settings = json.loads(SETTINGS_PATH.read_text(encoding="utf-8-sig"))
        site_url = str(settings["site_url"]).rstrip("/")
        if not site_url.startswith(("http://", "https://")):
            raise ValueError("site_url은 http:// 또는 https://로 시작해야 합니다.")
        category = clean_slug(str(settings["page_category"]))
        if args.production:
            output_dir, production_config = prepare_production_output(settings)
        else:
            output_dir = choose_output_dir(str(settings["output_dir"]))
            output_dir.mkdir(parents=True, exist_ok=False) if not output_dir.exists() else None
    except Exception as exc:
        print(f"설정 또는 출력 폴더 오류: {exc}", file=sys.stderr)
        return 1

    try:
        page_template = PAGE_TEMPLATE_PATH.read_text(encoding="utf-8-sig")
        home_template = HOME_TEMPLATE_PATH.read_text(encoding="utf-8-sig")
    except Exception as exc:
        print(f"템플릿 읽기 오류: {exc}", file=sys.stderr)
        return 1

    try:
        with data_path.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            fieldnames = set(reader.fieldnames or [])
            full_combination_input = FULL_COMBINATION_FIELDS.issubset(fieldnames)
            if not full_combination_input:
                missing_columns = [field for field in REQUIRED_FIELDS if field not in fieldnames]
                if missing_columns:
                    raise ValueError(f"CSV 누락 열: {', '.join(missing_columns)}")
            rows = [
                row for row in reader
                if not full_combination_input or row.get("is_valid", "").lower() == "true"
            ]
    except Exception as exc:
        print(f"데이터 읽기 오류: {exc}", file=sys.stderr)
        return 1

    try:
        content_config = dict(settings["content_generation"])
        pools = load_pools(ROOT / "data" / "sentence_pools")
        keyword_pool_names = (
            "description_general", "openings", "endings", "body_general",
            "body_elementary", "body_middle", "body_high", "body_korean",
            "body_english", "body_math", "body_science", "body_english_math",
            "body_internal_exam", "title_patterns", "title_modifiers", "title_endings",
        )
        keyword_pools = {
            name: load_keyword_pool(name) for name in keyword_pool_names
        } if full_combination_input else {}
    except Exception as exc:
        print(f"문장 풀 또는 콘텐츠 설정 오류: {exc}", file=sys.stderr)
        return 1

    def effective_keyword(row: dict[str, str]) -> str:
        return row.get("display_keyword", "").strip() or row.get("메인키워드", "").strip() or " ".join(
            (row.get("지역명", "").strip(), row.get("대상", "").strip(), row.get("과목", "").strip())
        ).strip()

    cleaned_slugs = [
        clean_slug(row.get("slug", "") or row.get("슬러그", "") or effective_keyword(row))
        for row in rows
    ]
    slug_counts = Counter(cleaned_slugs)
    duplicate_slugs = {slug for slug, count in slug_counts.items() if slug and count > 1}
    page_records: list[dict[str, str]] = []
    common_source_dir = ROOT / Path(str(settings["image_dir"]).replace("/", "\\"))
    common_candidates = sorted(
        (
            path for path in common_source_dir.iterdir()
            if path.is_file() and path.suffix.lower() in SUPPORTED_IMAGE_EXTENSIONS
        ),
        key=lambda path: unicodedata.normalize("NFC", path.name).casefold(),
    ) if common_source_dir.is_dir() else []
    common_source = common_candidates[0] if common_candidates else None
    common_exists = common_source is not None
    if len(common_candidates) > 1:
        warnings.append(
            "공통 이미지 폴더에 지원 이미지가 "
            f"{len(common_candidates)}개 있어 파일명 오름차순 첫 파일을 사용했습니다: {common_source.name}"
        )
    thumbnail_source_dir = ROOT / Path(str(settings["thumbnail_dir"]).replace("/", "\\"))
    thumbnail_files = sorted(
        (
            path for path in thumbnail_source_dir.iterdir()
            if path.is_file() and path.suffix.lower() in SUPPORTED_IMAGE_EXTENSIONS
        ),
        key=lambda path: unicodedata.normalize("NFC", path.name).casefold(),
    ) if thumbnail_source_dir.is_dir() else []
    selected_thumbnails: dict[str, str | None] = {}
    thumbnail_usage: Counter[str] = Counter()
    copied_image_destinations: set[Path] = set()
    duplicate_copy_skips = 0
    image_free_pages = 0
    missing_og_image_pages = 0
    missing_twitter_image_pages = 0
    body_thumbnail_tag_count = 0
    meta_only_thumbnail_count = 0
    used_titles: set[str] = set()
    used_descriptions: set[str] = set()
    used_bodies: set[str] = set()
    used_combinations: set[tuple[str, ...]] = set()
    previous_body_sets: list[set[str]] = []
    page_sentence_ids: dict[str, dict[str, list[str]]] = {}
    page_used_pools: dict[str, list[str]] = {}
    pool_usage: Counter[str] = Counter()
    title_pattern_usage: Counter[str] = Counter()
    target_expression_usage: Counter[str] = Counter()
    subject_expression_usage: Counter[str] = Counter()
    title_modifier_usage: Counter[str] = Counter()
    title_ending_usage: Counter[str] = Counter()
    title_target_mismatch_count = 0
    title_subject_mismatch_count = 0
    retried_page_count = 0
    similarity_warning_count = 0
    all_regions = {row.get("지역명", "").strip() for row in rows if row.get("지역명", "").strip()}

    common_public_path = (
        f"/{settings['image_dir'].strip('/')}/{quote(common_source.name)}"
        if common_source else None
    )
    if common_source:
        common_destination = (
            output_dir
            / Path(str(settings["image_dir"]).replace("/", "\\"))
            / common_source.name
        )
        common_destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(common_source, common_destination)
        copied_image_destinations.add(common_destination)
        created_files.append(str(common_destination.relative_to(ROOT)))
    else:
        warnings.append("공통 이미지 폴더에 지원 이미지가 없어 공통 이미지를 사용하지 않았습니다.")

    for line_number, row in enumerate(rows, start=2):
        try:
            is_full_combination = full_combination_input
            is_prebuilt = (
                is_full_combination
                or row.get("content_mode", "").strip() == "keyword_combination"
            )
            if is_full_combination:
                page_type = row.get("page_type", "").strip()
                entity = (
                    row.get("지역명", "").strip()
                    if page_type == "region"
                    else row.get("학교명", "").strip()
                )
                grade_expression = row.get("학년표현", "").strip()
                subject_expression = row.get("과목표현", "").strip()
                exam_used = row.get("내신사용", "").strip().lower() == "true"
                combo = {
                    "type": "full_combination",
                    "grade": grade_expression,
                    "subject": subject_expression,
                    "exam": exam_used,
                }
                generated = make_keyword_content(
                    entity,
                    combo,
                    row.get("content_seed", "").strip() or row.get("slug", "").strip(),
                    keyword_pools,
                )
                keyword = row.get("display_keyword", "").strip()
                target = TARGET_BY_EXPRESSION.get(grade_expression, "")
                used_pool_names = [str(item) for item in generated["used_pools"]]
                content = {
                    "region": entity,
                    "target": target,
                    "subject": subject_expression,
                    "keyword": keyword,
                    "slug": row.get("slug", "").strip(),
                    "title": row.get("title", "").strip(),
                    "description": row.get("description", "").strip(),
                    "body_text": str(generated["body"]),
                    "keywords": [
                        value for value in (
                            keyword,
                            entity + subject_expression + "과외",
                            entity + grade_expression + "과외",
                            grade_expression + subject_expression,
                        ) if value
                    ][:4],
                    "description_ids": [],
                    "body_ids": [str(item) for item in generated["body_ids"]],
                    "used_pools": used_pool_names,
                    "pool_counts": {name: 1 for name in used_pool_names},
                    "retry_count": 0,
                    "similarity_warning": False,
                    "title_components": {
                        "pattern_id": "full_combination",
                        "pattern": "full_combination",
                        "target_expression": grade_expression,
                        "subject_expression": subject_expression,
                        "modifier": "",
                        "ending": "",
                    },
                }
            elif is_prebuilt:
                keyword = row.get("메인키워드", "").strip()
                content = {
                    "region": row.get("지역명", "").strip(),
                    "target": row.get("대상", "").strip(),
                    "subject": row.get("과목", "").strip(),
                    "keyword": keyword,
                    "slug": row.get("슬러그", "").strip() or re.sub(r"\s+", "", keyword),
                    "title": row.get("제목", "").strip(),
                    "description": row.get("설명", "").strip(),
                    "body_text": row.get("본문", "").strip(),
                    "keywords": [
                        value for value in (
                            keyword,
                            row.get("지역명", "").strip() + row.get("과목", "").strip(),
                            row.get("지역명", "").strip() + row.get("학년표현", "").strip() + "과외",
                            row.get("학년표현", "").strip() + row.get("과목", "").strip(),
                        ) if value
                    ][:4],
                    "description_ids": [],
                    "body_ids": [
                        value for value in row.get("body_sentence_ids", "").split("|") if value
                    ],
                    "used_pools": [
                        value for value in row.get("used_sentence_pools", "").split("|") if value
                    ],
                    "pool_counts": {},
                    "retry_count": 0,
                    "similarity_warning": False,
                    "title_components": {
                        "pattern_id": row.get("title_pattern_id", "keyword_combination"),
                        "pattern": "keyword_combination",
                        "target_expression": row.get("학년표현", "").strip(),
                        "subject_expression": row.get("과목표현", "").strip(),
                        "modifier": row.get("title_modifier", ""),
                        "ending": row.get("title_ending", ""),
                    },
                }
            else:
                content = generate_page_content(
                    row, pools, content_config, previous_body_sets,
                    used_titles, used_descriptions, used_bodies, used_combinations,
                )
            slug = clean_slug(str(content["slug"]))
            title = str(content["title"])
            description = str(content["description"])
            body = str(content["body_text"])
            region = str(content["region"])
            subject = str(content["subject"])
            keyword = str(content["keyword"])
            if not slug:
                raise ValueError("슬러그가 비어 있습니다.")
            if slug in duplicate_slugs:
                raise ValueError(f"중복 슬러그: {slug}")
            title_generated_count += 1
            description_generated_count += 1
            body_generated_count += 1
            combined = " ".join((keyword, title, description, body, slug))
            location_mismatches += int(bool(region) and region not in combined)
            checked_subject = row.get("과목표현", "").strip() if is_prebuilt else subject
            subject_mismatches += int(bool(checked_subject) and checked_subject not in combined)
            abnormal_count += int(any(marker in combined for marker in BROKEN_MARKERS) or bool(CJK_FOREIGN_RE.search(combined)))
            if not title or not description or not body:
                raise ValueError("title, description 또는 본문이 비어 있습니다.")
            content_validation = (
                {"grade_mismatch": 0, "subject_mismatch": 0, "other_region": 0, "forbidden": 0, "abnormal": 0}
                if is_prebuilt else validate_content(content, all_regions)
            )
            grade_mismatch_count += content_validation["grade_mismatch"]
            content_subject_mismatch_count += content_validation["subject_mismatch"]
            other_region_count += content_validation["other_region"]
            forbidden_expression_count += content_validation["forbidden"]
            content_abnormal_count += content_validation["abnormal"]
            page_type = row.get("page_type", "sample").strip() or "sample"
            page_type_counts[page_type] += 1
            if is_full_combination:
                title_core_terms = [
                    term for term in row.get("title_core_terms", "").split("|") if term
                ]
                title_core_body_missing_count += int(
                    any(term not in body for term in title_core_terms)
                )
                title_core_description_missing_count += int(
                    any(term not in description for term in title_core_terms)
                )
                grade_expression = row.get("학년표현", "").strip()
                exam_used = row.get("내신사용", "").strip().lower() == "true"
                elementary_exam_count += int(
                    grade_expression in ("초등", "초등학생") and exam_used
                )
                school_grade = row.get("학교급", "").strip()
                if page_type == "school":
                    school_grade_expression_mismatch_count += int(
                        (
                            school_grade == "중학교"
                            and grade_expression not in ("중등", "중학생")
                        )
                        or (
                            school_grade == "고등학교"
                            and grade_expression not in ("고등", "고등학생")
                        )
                        or school_grade not in ("중학교", "고등학교")
                    )
            if page_type == "region" and row.get("지역명", "").strip() not in keyword:
                region_name_error_count += 1
            if page_type == "school" and row.get("학교명", "").strip() not in keyword:
                school_name_error_count += 1
            if row.get("대상", "").strip() and row["대상"].strip() != content["target"]:
                target_error_count += 1
            if row.get("과목", "").strip() and row["과목"].strip() != content["subject"]:
                subject_error_count += 1
            retried_page_count += int(int(content["retry_count"]) > 0)
            similarity_warning_count += int(bool(content["similarity_warning"]))

            encoded_path = url_path(category, slug)
            canonical = site_join(site_url, encoded_path)
            expected_prefix = site_url + "/"
            if not canonical.startswith(expected_prefix):
                invalid_canonical += 1

            image_blocks = []
            preload_image = ""
            if common_public_path:
                image_blocks.append(
                    f'    <figure><img src="{html.escape(common_public_path, quote=True)}" '
                    f'alt="{html.escape(keyword + " 학습 안내", quote=True)}"></figure>'
                )
                preload_image = f'  <link rel="preload" as="image" href="{html.escape(common_public_path, quote=True)}">'

            selected_thumbnail: Path | None = None
            selected_thumbnail_public_path: str | None = None
            if thumbnail_files:
                digest = hashlib.sha256(slug.encode("utf-8")).hexdigest()
                selected_thumbnail = thumbnail_files[int(digest, 16) % len(thumbnail_files)]
                selected_thumbnails[slug] = selected_thumbnail.name
                thumbnail_usage[selected_thumbnail.name] += 1
                selected_thumbnail_public_path = (
                    f"/{settings['thumbnail_dir'].strip('/')}/{quote(selected_thumbnail.name)}"
                )
                meta_only_thumbnail_count += 1
                thumbnail_destination = (
                    output_dir
                    / Path(str(settings["thumbnail_dir"]).replace("/", "\\"))
                    / selected_thumbnail.name
                )
                if thumbnail_destination in copied_image_destinations:
                    duplicate_copy_skips += 1
                else:
                    thumbnail_destination.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(selected_thumbnail, thumbnail_destination)
                    copied_image_destinations.add(thumbnail_destination)
                    created_files.append(str(thumbnail_destination.relative_to(ROOT)))
            else:
                selected_thumbnails[slug] = None

            og_image_path = selected_thumbnail_public_path or common_public_path
            og_image_meta = ""
            twitter_image_meta = ""
            if og_image_path:
                absolute_meta_image = html.escape(site_join(site_url, og_image_path), quote=True)
                og_image_meta = (
                    f'  <meta property="og:image" '
                    f'content="{absolute_meta_image}">'
                )
                twitter_image_meta = f'  <meta name="twitter:image" content="{absolute_meta_image}">'
            else:
                missing_og_image_pages += 1
                missing_twitter_image_pages += 1
            if not common_exists and selected_thumbnail is None:
                image_free_pages += 1

            page_html = render(page_template, {
                "language": html.escape(str(settings["language"]), quote=True),
                "title": html.escape(title),
                "description": html.escape(description, quote=True),
                "keywords": html.escape(", ".join(content["keywords"]), quote=True),
                "canonical": html.escape(canonical, quote=True),
                "og_image_meta": og_image_meta,
                "twitter_image_meta": twitter_image_meta,
                "preload_image": preload_image,
                "background_color": html.escape(str(settings["background_color"])),
                "max_width": int(settings["max_width"]),
                "image_blocks": "\n".join(image_blocks),
                "keyword": html.escape(keyword),
                "body_html": html.escape(keyword) + "<br>\n" + html.escape(body),
            })
            page_dir = output_dir / category / slug
            page_dir.mkdir(parents=True, exist_ok=False)
            page_file = page_dir / "index.html"
            page_file.write_text(page_html, encoding="utf-8")
            created_files.append(str(page_file.relative_to(ROOT)))
            body_ids = tuple(str(item) for item in content["body_ids"])
            page_records.append({
                "title": title,
                "description": description,
                "body": body,
                "body_ids": list(body_ids),
                "keyword": keyword,
                "slug": slug,
                "path": encoded_path,
                "canonical": canonical,
                "page_type": page_type,
                "grade_expression": row.get("학년표현", "").strip(),
                "subject_expression": row.get("과목표현", "").strip(),
                "internal_exam": row.get("내신사용", "").strip().lower() == "true",
                "local_file": str(page_file.resolve()),
            })
            used_titles.add(title)
            used_descriptions.add(description)
            used_bodies.add(body)
            used_combinations.add(body_ids)
            previous_body_sets.append(set(body_ids))
            page_sentence_ids[slug] = {
                "description": [str(item) for item in content["description_ids"]],
                "body": list(body_ids),
            }
            page_used_pools[slug] = [str(item) for item in content["used_pools"]]
            for pool_name, count in content["pool_counts"].items():
                pool_usage[str(pool_name)] += int(count)
            title_components = dict(content["title_components"])
            title_pattern_usage[str(title_components["pattern_id"])] += 1
            target_expression_usage[str(title_components["target_expression"])] += 1
            subject_expression_usage[str(title_components["subject_expression"])] += 1
            title_modifier_usage[str(title_components["modifier"])] += 1
            title_ending_usage[str(title_components["ending"])] += 1
            if not is_prebuilt:
                if title_components["target_expression"] not in TARGET_EXPRESSIONS[str(content["target"])]:
                    title_target_mismatch_count += 1
                subject_key = SUBJECT_MAP[str(content["subject"])]
                if title_components["subject_expression"] not in SUBJECT_EXPRESSIONS[subject_key]:
                    title_subject_mismatch_count += 1
        except Exception as exc:
            failed_pages.append(
                row.get("slug", "") or row.get("슬러그", "") or f"CSV {line_number}행"
            )
            errors.append(f"CSV {line_number}행: {exc}")

    title_duplicate_count = len(page_records) - len({record["title"] for record in page_records})
    description_duplicate_count = len(page_records) - len({record["description"] for record in page_records})
    body_duplicate_count = len(page_records) - len({record["body"] for record in page_records})
    combination_counts = Counter(tuple(record["body_ids"]) for record in page_records)
    sentence_combination_duplicate_count = sum(count - 1 for count in combination_counts.values() if count > 1)
    title_lengths = [len(str(record["title"])) for record in page_records]
    title_minimum_length = min(title_lengths, default=0)
    title_maximum_length = max(title_lengths, default=0)
    title_average_length = round(sum(title_lengths) / len(title_lengths), 2) if title_lengths else 0
    short_title_count = sum(length < int(content_config.get("title_min_length", 45)) for length in title_lengths)
    long_title_count = sum(length > int(content_config.get("title_max_length", 90)) for length in title_lengths)

    unique_urls = list(dict.fromkeys([site_url + "/"] + [record["canonical"] for record in page_records]))
    sitemap_lines = ['<?xml version="1.0" encoding="UTF-8"?>', '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">']
    sitemap_lines.extend(f"  <url><loc>{xml_escape(url)}</loc></url>" for url in unique_urls)
    sitemap_lines.append("</urlset>")
    sitemap_file = output_dir / "sitemap.xml"
    sitemap_file.write_text("\n".join(sitemap_lines) + "\n", encoding="utf-8")
    created_files.append(str(sitemap_file.relative_to(ROOT)))

    robots_file = output_dir / "robots.txt"
    robots_file.write_text(f"User-agent: *\nAllow: /\n\nSitemap: {site_join(site_url, '/sitemap.xml')}\n", encoding="utf-8")
    created_files.append(str(robots_file.relative_to(ROOT)))

    page_links = "\n".join(
        f'      <li><a href="{record["path"]}" title="{html.escape(record["title"], quote=True)}">'
        f'{html.escape(record["keyword"])}</a></li>' for record in page_records
    )
    home_html = render(home_template, {
        "language": html.escape(str(settings["language"]), quote=True),
        "site_name": html.escape(str(settings["site_name"])),
        "site_url": html.escape(site_url),
        "page_count": len(page_records),
        "page_links": page_links,
        "background_color": html.escape(str(settings["background_color"])),
        "max_width": int(settings["max_width"]),
    })
    home_file = output_dir / "index.html"
    home_file.write_text(home_html, encoding="utf-8")
    created_files.append(str(home_file.relative_to(ROOT)))

    production_validation: dict[str, object] = {}
    fixed_sample_pages: list[dict[str, object]] = []
    if args.production and production_config.get("verify_after_build", True):
        production_validation = verify_production_output(
            output_dir, site_url, len(page_records), len(unique_urls)
        )
        errors.extend(str(item) for item in production_validation["errors"])
        warnings.extend(str(item) for item in production_validation["warnings"])
        production_quality_errors = {
            "title 중복": title_duplicate_count,
            "description 중복": description_duplicate_count,
            "본문 중복": body_duplicate_count,
            "제목 핵심어 본문 누락": title_core_body_missing_count,
            "제목 핵심어 description 누락": title_core_description_missing_count,
            "초등 + 내신": elementary_exam_count,
            "학교급 불일치": school_grade_expression_mismatch_count,
        }
        errors.extend(
            f"{name} {count}건"
            for name, count in production_quality_errors.items()
            if count
        )
    if args.production:
        fixed_sample_pages = fixed_representative_pages(page_records)
        if len(fixed_sample_pages) != 30:
            errors.append(f"고정 검수 샘플이 30개가 아닙니다: {len(fixed_sample_pages)}")

    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    if args.production:
        text_report_file = REPORTS_DIR / "production_build_report.txt"
        json_report_file = REPORTS_DIR / "production_build_report.json"
    else:
        text_report_file = choose_report_file("build_report.txt")
        json_report_file = choose_report_file("build_report.json")
    created_files.extend([
        str(text_report_file.relative_to(ROOT)),
        str(json_report_file.relative_to(ROOT)),
    ])
    source_files = [
        "config/settings.json",
        str(data_path.relative_to(ROOT)).replace("\\", "/"),
        "data/sentence_pools/*.json",
        "templates/page.html",
        "templates/home.html",
        "scripts/build.py",
        "scripts/content_generator.py",
        "scripts/create_sentence_pools.py",
    ]
    report = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "build_started_at": build_started_at.isoformat(timespec="seconds"),
        "build_duration_seconds": round(time.perf_counter() - build_timer, 3),
        "build_mode": "production" if args.production else "candidate",
        "output_dir": str(output_dir.relative_to(ROOT)),
        "generated_page_count": len(page_records),
        "sitemap_url_count": len(unique_urls),
        "duplicate_slug_count": len(duplicate_slugs),
        "missing_title_count": missing_title,
        "missing_description_count": missing_description,
        "missing_body_count": missing_body,
        "invalid_canonical_count": invalid_canonical,
        "missing_image_count": len(missing_images),
        "missing_images": sorted(missing_images),
        "common_image_exists": common_exists,
        "common_image_filename": common_source.name if common_source else None,
        "common_image_used": common_exists and bool(page_records),
        "thumbnail_image_count": len(thumbnail_files),
        "thumbnail_image_files": [path.name for path in thumbnail_files],
        "thumbnail_image_used": bool(thumbnail_usage),
        "selected_thumbnail_by_page": {} if args.production else selected_thumbnails,
        "thumbnail_usage_count": dict(sorted(thumbnail_usage.items())),
        "copied_image_count": len(copied_image_destinations),
        "duplicate_copy_skip_count": duplicate_copy_skips,
        "image_free_page_count": image_free_pages,
        "missing_og_image_page_count": missing_og_image_pages,
        "missing_twitter_image_page_count": missing_twitter_image_pages,
        "body_thumbnail_tag_count": body_thumbnail_tag_count,
        "meta_only_thumbnail_count": meta_only_thumbnail_count,
        "location_mismatch_count": location_mismatches,
        "subject_mismatch_count": subject_mismatches,
        "abnormal_character_count": abnormal_count,
        "failed_page_count": len(failed_pages),
        "failed_pages": failed_pages,
        "content_generation_mode": content_config.get("mode", "sentence_pool"),
        "title_generated_count": title_generated_count,
        "description_generated_count": description_generated_count,
        "body_generated_count": body_generated_count,
        "title_duplicate_count": title_duplicate_count,
        "title_minimum_length": title_minimum_length,
        "title_maximum_length": title_maximum_length,
        "title_average_length": title_average_length,
        "title_below_minimum_count": short_title_count,
        "title_above_maximum_count": long_title_count,
        "title_pattern_usage_distribution": dict(sorted(title_pattern_usage.items())),
        "target_expression_usage_count": dict(sorted(target_expression_usage.items())),
        "subject_expression_usage_count": dict(sorted(subject_expression_usage.items())),
        "title_modifier_usage_count": dict(sorted(title_modifier_usage.items())),
        "title_ending_usage_count": dict(sorted(title_ending_usage.items())),
        "title_target_mismatch_count": title_target_mismatch_count,
        "title_subject_mismatch_count": title_subject_mismatch_count,
        "title_core_body_missing_count": title_core_body_missing_count,
        "title_core_description_missing_count": title_core_description_missing_count,
        "elementary_internal_exam_count": elementary_exam_count,
        "school_grade_expression_mismatch_count": school_grade_expression_mismatch_count,
        "description_duplicate_count": description_duplicate_count,
        "body_duplicate_count": body_duplicate_count,
        "sentence_combination_duplicate_count": sentence_combination_duplicate_count,
        "sentence_similarity_warning_count": similarity_warning_count,
        "retried_page_count": retried_page_count,
        "sentence_ids_by_page": {} if args.production else page_sentence_ids,
        "sentence_pools_by_page": {} if args.production else page_used_pools,
        "sentence_pool_usage_count": dict(sorted(pool_usage.items())),
        "grade_mismatch_sentence_count": grade_mismatch_count,
        "content_subject_mismatch_sentence_count": content_subject_mismatch_count,
        "other_region_intrusion_count": other_region_count,
        "forbidden_expression_count": forbidden_expression_count,
        "content_abnormal_character_count": content_abnormal_count,
        "page_type_count": dict(sorted(page_type_counts.items())),
        "region_name_error_count": region_name_error_count,
        "school_name_error_count": school_name_error_count,
        "target_error_count": target_error_count,
        "subject_error_count": subject_error_count,
        "representative_page_urls": [record["canonical"] for record in page_records[:10]],
        "created_or_modified_source_files": source_files,
        "generated_files": (
            [
                "site/",
                "site/index.html",
                "site/sitemap.xml",
                "site/robots.txt",
                "reports/production_build_report.txt",
                "reports/production_build_report.json",
            ]
            if args.production else created_files
        ),
        "production_settings": production_config if args.production else None,
        "production_verification": production_validation if args.production else None,
        "fixed_representative_pages": fixed_sample_pages if args.production else [],
        "errors": errors,
        "warnings": warnings,
    }
    json_report_file.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    lines = ([
        "공부업 Production 빌드 보고서",
        f"생성 시작: {report['build_started_at']}",
        f"생성 완료: {report['generated_at']}",
        f"생성 시간(초): {report['build_duration_seconds']}",
        f"출력 폴더: {report['output_dir']}",
        f"생성 페이지: {report['generated_page_count']}",
        f"HTML 수: {production_validation.get('html_count', '검증 안 함')}",
        f"sitemap URL 수: {production_validation.get('sitemap_url_count', '검증 안 함')}",
        f"출력 용량(bytes): {production_validation.get('output_size_bytes', 0)}",
        f"깨진 링크: {production_validation.get('broken_link_count', 0)}",
        f"깨진 이미지: {production_validation.get('broken_image_count', 0)}",
        f"canonical 오류: {production_validation.get('canonical_error_count', 0)}",
        f"og:image 경로 오류: {production_validation.get('og_image_error_count', 0)}",
        f"title 누락: {production_validation.get('missing_title_count', 0)}",
        f"description 누락: {production_validation.get('missing_description_count', 0)}",
        f"본문 누락: {production_validation.get('missing_body_count', 0)}",
        f"슬러그 중복: {production_validation.get('duplicate_slug_count', 0)}",
        f"title 중복: {title_duplicate_count}",
        f"description 중복: {description_duplicate_count}",
        f"본문 중복: {body_duplicate_count}",
        f"제목 핵심어 본문 누락: {title_core_body_missing_count}",
        f"제목 핵심어 description 누락: {title_core_description_missing_count}",
        f"초등 + 내신: {elementary_exam_count}",
        f"학교급 불일치: {school_grade_expression_mismatch_count}",
        "",
        "고정 검수 샘플 30페이지:",
        *[
            f"- [{item['page_type']}] {item['local_file']} "
            f"(학년={item['grade_expression'] or '-'}, "
            f"과목={item['subject_expression'] or '-'}, "
            f"내신={item['internal_exam']})"
            for item in fixed_sample_pages
        ],
        "",
        "오류:",
        *([f"- {error}" for error in errors] or ["- 없음"]),
        "",
        "경고:",
        *([f"- {warning}" for warning in warnings] or ["- 없음"]),
    ] if args.production else [
        "공부업 정적 사이트 빌드 보고서",
        f"생성 시각: {report['generated_at']}",
        f"출력 폴더: {report['output_dir']}",
        f"생성 페이지 수: {len(page_records)}",
        f"sitemap URL 수: {len(unique_urls)}",
        f"중복 슬러그 수: {len(duplicate_slugs)}",
        f"누락 title 수: {missing_title}",
        f"누락 description 수: {missing_description}",
        f"누락 본문 수: {missing_body}",
        f"잘못된 canonical 수: {invalid_canonical}",
        f"누락 이미지 수: {len(missing_images)}",
        f"공통 이미지 존재 여부: {common_exists}",
        f"공통 이미지 파일명: {common_source.name if common_source else '없음'}",
        f"공통 이미지 사용 여부: {common_exists and bool(page_records)}",
        f"대표 이미지 전체 개수: {len(thumbnail_files)}",
        f"대표 이미지 사용 여부: {bool(thumbnail_usage)}",
        f"실제 복사된 이미지 수: {len(copied_image_destinations)}",
        f"중복 복사 생략 수: {duplicate_copy_skips}",
        f"이미지 없는 페이지 수: {image_free_pages}",
        f"og:image 누락 페이지 수: {missing_og_image_pages}",
        f"twitter:image 누락 페이지 수: {missing_twitter_image_pages}",
        f"본문에 출력된 대표 이미지 태그 수: {body_thumbnail_tag_count}",
        f"메타 전용 대표 이미지 수: {meta_only_thumbnail_count}",
        f"지역명 불일치 수: {location_mismatches}",
        f"과목명 불일치 수: {subject_mismatches}",
        f"깨진 문자 또는 비정상 문자 수: {abnormal_count}",
        f"생성 실패 페이지 수: {len(failed_pages)}",
        f"콘텐츠 생성 모드: {content_config.get('mode', 'sentence_pool')}",
        f"제목 생성 수: {title_generated_count}",
        f"description 생성 수: {description_generated_count}",
        f"본문 생성 수: {body_generated_count}",
        f"title 중복 수: {title_duplicate_count}",
        f"제목 최소 길이: {title_minimum_length}",
        f"제목 최대 길이: {title_maximum_length}",
        f"제목 평균 길이: {title_average_length}",
        f"45자 미만 제목 수: {short_title_count}",
        f"90자 초과 제목 수: {long_title_count}",
        f"대상 불일치 제목 수: {title_target_mismatch_count}",
        f"과목 불일치 제목 수: {title_subject_mismatch_count}",
        f"description 완전 중복 수: {description_duplicate_count}",
        f"본문 완전 중복 수: {body_duplicate_count}",
        f"동일 문장 조합 수: {sentence_combination_duplicate_count}",
        f"문장 공유율 경고 수: {similarity_warning_count}",
        f"재선택 실행 페이지 수: {retried_page_count}",
        f"학년 불일치 문장 수: {grade_mismatch_count}",
        f"과목 불일치 문장 수: {content_subject_mismatch_count}",
        f"다른 지역명 혼입 수: {other_region_count}",
        f"금지 표현 발견 수: {forbidden_expression_count}",
        f"비정상 문자 수: {content_abnormal_count}",
        f"지역 페이지 수: {page_type_counts.get('region', 0)}",
        f"학교 페이지 수: {page_type_counts.get('school', 0)}",
        f"지역명 오류 수: {region_name_error_count}",
        f"학교명 오류 수: {school_name_error_count}",
        f"대상 오류 수: {target_error_count}",
        f"과목 오류 수: {subject_error_count}",
        "",
        "대표 페이지 URL:",
        *[f"- {url}" for url in report["representative_page_urls"]],
        "",
        "수정 또는 생성한 소스 파일:",
        *[f"- {path}" for path in source_files],
        "",
        "생성 결과 파일:",
        *[f"- {path}" for path in created_files],
        "",
        "누락 이미지:",
        *[f"- {path}" for path in sorted(missing_images)],
        "",
        "페이지별 선택된 대표 이미지:",
        *[f"- {slug}: {name or '없음'}" for slug, name in selected_thumbnails.items()],
        "",
        "대표 이미지 파일 목록:",
        *([f"- {path.name}" for path in thumbnail_files] or ["- 없음"]),
        "",
        "대표 이미지 사용 횟수:",
        *([f"- {name}: {count}" for name, count in sorted(thumbnail_usage.items())] or ["- 없음"]),
        "",
        "페이지별 사용 문장 ID:",
        *[
            f"- {slug}: description={', '.join(values['description'])}; body={', '.join(values['body'])}"
            for slug, values in page_sentence_ids.items()
        ],
        "",
        "페이지별 사용 문장 풀:",
        *[f"- {slug}: {', '.join(pool_names)}" for slug, pool_names in page_used_pools.items()],
        "",
        "문장 풀별 사용 횟수:",
        *[f"- {pool_name}: {count}" for pool_name, count in sorted(pool_usage.items())],
        "",
        "제목 패턴 중복 분포:",
        *[f"- {name}: {count}" for name, count in sorted(title_pattern_usage.items())],
        "",
        "대상 표현별 사용 횟수:",
        *[f"- {name}: {count}" for name, count in sorted(target_expression_usage.items())],
        "",
        "과목 표현별 사용 횟수:",
        *[f"- {name}: {count}" for name, count in sorted(subject_expression_usage.items())],
        "",
        "수식어별 사용 횟수:",
        *[f"- {name}: {count}" for name, count in sorted(title_modifier_usage.items())],
        "",
        "마무리 문구별 사용 횟수:",
        *[f"- {name}: {count}" for name, count in sorted(title_ending_usage.items())],
        "",
        "오류:",
        *([f"- {error}" for error in errors] or ["- 없음"]),
        "",
        "경고:",
        *([f"- {warning}" for warning in warnings] or ["- 없음"]),
    ])
    text_report_file.write_text("\n".join(lines) + "\n", encoding="utf-8")

    print(json.dumps({
        "output_dir": str(output_dir),
        "generated_page_count": len(page_records),
        "sitemap_url_count": len(unique_urls),
        "missing_image_count": len(missing_images),
        "common_image_exists": common_exists,
        "common_image_filename": common_source.name if common_source else None,
        "common_image_used": common_exists and bool(page_records),
        "thumbnail_image_count": len(thumbnail_files),
        "thumbnail_image_files": [path.name for path in thumbnail_files],
        "thumbnail_image_used": bool(thumbnail_usage),
        "used_thumbnail_count": len(thumbnail_usage),
        "copied_image_count": len(copied_image_destinations),
        "duplicate_copy_skip_count": duplicate_copy_skips,
        "image_free_page_count": image_free_pages,
        "missing_og_image_page_count": missing_og_image_pages,
        "missing_twitter_image_page_count": missing_twitter_image_pages,
        "body_thumbnail_tag_count": body_thumbnail_tag_count,
        "meta_only_thumbnail_count": meta_only_thumbnail_count,
        "duplicate_slug_count": len(duplicate_slugs),
        "invalid_canonical_count": invalid_canonical,
        "location_mismatch_count": location_mismatches,
        "subject_mismatch_count": subject_mismatches,
        "failed_page_count": len(failed_pages),
        "content_generation_mode": content_config.get("mode", "sentence_pool"),
        "title_duplicate_count": title_duplicate_count,
        "title_minimum_length": title_minimum_length,
        "title_maximum_length": title_maximum_length,
        "title_average_length": title_average_length,
        "title_below_minimum_count": short_title_count,
        "title_above_maximum_count": long_title_count,
        "title_target_mismatch_count": title_target_mismatch_count,
        "title_subject_mismatch_count": title_subject_mismatch_count,
        "description_duplicate_count": description_duplicate_count,
        "body_duplicate_count": body_duplicate_count,
        "sentence_combination_duplicate_count": sentence_combination_duplicate_count,
        "sentence_similarity_warning_count": similarity_warning_count,
        "retried_page_count": retried_page_count,
        "grade_mismatch_sentence_count": grade_mismatch_count,
        "content_subject_mismatch_sentence_count": content_subject_mismatch_count,
        "other_region_intrusion_count": other_region_count,
        "forbidden_expression_count": forbidden_expression_count,
        "content_abnormal_character_count": content_abnormal_count,
        "page_type_count": dict(sorted(page_type_counts.items())),
        "region_name_error_count": region_name_error_count,
        "school_name_error_count": school_name_error_count,
        "target_error_count": target_error_count,
        "subject_error_count": subject_error_count,
        "errors": errors,
        "warnings": warnings,
    }, ensure_ascii=False, indent=2))
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
