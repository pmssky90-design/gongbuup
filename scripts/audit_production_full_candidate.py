from __future__ import annotations

import argparse
import html
import json
import re
import sys
from collections import Counter, deque
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from urllib.parse import quote, unquote, urlsplit


ROOT = Path(__file__).resolve().parents[1]
REPORTS = ROOT / "reports"
OLD_NAME = "이화여자대학교사범대학부속이화?금란중학교"
OFFICIAL_NAME = "이화여자대학교사범대학부속이화·금란중학교"
TARGET_SUFFIXES = ("과외", "수학과외", "영어과외", "국어과외", "과학과외", "내신과외")

TITLE_RE = re.compile(r"<title>(.*?)</title>", re.I | re.S)
DESC_RE = re.compile(
    r'<meta\s+name=["\']description["\']\s+content=["\'](.*?)["\']', re.I | re.S
)
CANONICAL_RE = re.compile(
    r'<link\s+rel=["\']canonical["\']\s+href=["\'](.*?)["\']', re.I | re.S
)
H1_RE = re.compile(r"<h1\b[^>]*>(.*?)</h1>", re.I | re.S)
HREF_RE = re.compile(r'<a\b[^>]+href=["\']([^"\']+)["\']', re.I)
IMG_RE = re.compile(r'<img\b[^>]+src=["\']([^"\']+)["\']', re.I)
CSS_RE = re.compile(
    r'<link\b[^>]+rel=["\']stylesheet["\'][^>]+href=["\']([^"\']+)["\']', re.I
)
SCRIPT_RE = re.compile(r'<script\b[^>]+src=["\']([^"\']+)["\']', re.I)
JSONLD_RE = re.compile(
    r'<script\b[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
    re.I | re.S,
)


def url_path(root: Path, file: Path) -> str:
    relative = file.relative_to(root)
    if relative == Path("index.html"):
        return "/"
    return "/" + "/".join(quote(part, safe="") for part in relative.parent.parts) + "/"


def normalized_local_path(value: str) -> str:
    parsed = urlsplit(html.unescape(value))
    if parsed.scheme or parsed.netloc:
        return ""
    parts = [quote(unquote(part), safe="") for part in parsed.path.strip("/").split("/") if part]
    if not parts:
        return "/"
    return "/" + "/".join(parts) + ("/" if parsed.path.endswith("/") else "")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--candidate")
    parser.add_argument("--similarity-report")
    parser.add_argument("--output-report")
    args = parser.parse_args()
    source_report = json.loads(
        (REPORTS / "production_full_candidate_validation.json").read_text(
            encoding="utf-8"
        )
    )
    output = (
        Path(args.candidate).resolve()
        if args.candidate else Path(source_report["candidate_output"])
    )
    similarity = (
        json.loads(Path(args.similarity_report).read_text(encoding="utf-8"))
        if args.similarity_report else None
    )
    site_url = "https://example.local"
    files = list(output.rglob("*.html"))
    paths = {url_path(output, file) for file in files}
    incoming = Counter()
    graph: dict[str, set[str]] = {}

    def inspect(file: Path) -> dict[str, object]:
        text = file.read_text(encoding="utf-8")
        path = url_path(output, file)
        title = TITLE_RE.findall(text)
        description = DESC_RE.findall(text)
        canonical = CANONICAL_RE.findall(text)
        h1 = H1_RE.findall(text)
        links = [normalized_local_path(value) for value in HREF_RE.findall(text)]
        links = [value for value in links if value]
        images = [normalized_local_path(value) for value in IMG_RE.findall(text)]
        css = [normalized_local_path(value) for value in CSS_RE.findall(text)]
        scripts = [normalized_local_path(value) for value in SCRIPT_RE.findall(text)]
        json_errors = 0
        for block in JSONLD_RE.findall(text):
            try:
                json.loads(html.unescape(block))
            except (ValueError, TypeError):
                json_errors += 1
        expected_canonical = site_url + path
        return {
            "file": str(file),
            "path": path,
            "title_error": len(title) != 1 or not title[0].strip(),
            "description_error": len(description) != 1 or not description[0].strip(),
            "canonical_error": len(canonical) != 1 or canonical[0] != expected_canonical,
            "canonical": canonical[0] if len(canonical) == 1 else "",
            "h1_error": len(h1) != 1,
            "jsonld_error": json_errors,
            "links": links,
            "images": images,
            "css": css,
            "scripts": scripts,
            "old_name_count": text.count(OLD_NAME),
            "official_name_count": text.count(OFFICIAL_NAME),
        }

    with ThreadPoolExecutor(max_workers=8) as pool:
        rows = list(pool.map(inspect, files, chunksize=64))

    broken_links = 0
    broken_images = 0
    css_errors = 0
    js_errors = 0
    canonical_values: list[str] = []
    target_rows = []
    for row in rows:
        path = str(row["path"])
        valid_links: set[str] = set()
        for target in row["links"]:
            if target not in paths:
                broken_links += 1
            elif target != path:
                valid_links.add(target)
                incoming[target] += 1
        graph[path] = valid_links
        for target in row["images"]:
            disk = output / Path(unquote(target).strip("/").replace("/", "\\"))
            broken_images += int(not disk.is_file())
        for target in row["css"]:
            disk = output / Path(unquote(target).strip("/").replace("/", "\\"))
            css_errors += int(not disk.is_file())
        for target in row["scripts"]:
            disk = output / Path(unquote(target).strip("/").replace("/", "\\"))
            js_errors += int(not disk.is_file())
        if row["canonical"]:
            canonical_values.append(str(row["canonical"]))
        if OFFICIAL_NAME in unquote(path):
            target_rows.append(row)

    sitemap_urls: list[str] = []
    for sitemap in output.glob("sitemap-*.xml"):
        sitemap_urls.extend(re.findall(
            r"<loc>(.*?)</loc>",
            sitemap.read_text(encoding="utf-8"),
            re.S,
        ))
    sitemap_paths = {
        normalized_local_path(urlsplit(value).path) for value in sitemap_urls
    }
    sitemap_mismatch = len(paths.symmetric_difference(sitemap_paths))
    depths = {"/": 0}
    queue = deque(["/"])
    while queue:
        current = queue.popleft()
        for target in graph.get(current, set()):
            if target not in depths:
                depths[target] = depths[current] + 1
                queue.append(target)
    orphan_count = sum(path != "/" and incoming[path] == 0 for path in paths)
    unreachable = sum(path not in depths for path in paths)
    target_expected_paths = {
        "/" + quote("과외", safe="") + "/" + quote(OFFICIAL_NAME + suffix, safe="") + "/"
        for suffix in TARGET_SUFFIXES
    }
    target_consistency_errors = sum(
        bool(row["title_error"])
        or bool(row["description_error"])
        or bool(row["canonical_error"])
        or bool(row["h1_error"])
        or int(row["official_name_count"]) == 0
        for row in target_rows
    )
    summary = {
        "candidate_output": str(output),
        "corrected_school_name": OFFICIAL_NAME,
        "corrected_page_count": len(target_rows),
        "corrected_expected_page_count": 6,
        "corrected_path_set_match": {
            str(row["path"]) for row in target_rows
        } == target_expected_paths,
        "corrected_page_consistency_error_count": target_consistency_errors,
        "old_name_residual_count": sum(int(row["old_name_count"]) for row in rows),
        "question_mark_slug_count": source_report["correction_preflight"][
            "question_mark_slug_count"
        ],
        "content_page_count": (
            similarity["content_page_count"]
            if similarity else source_report["content_page_count"]
        ),
        "html_count": len(files),
        "sitemap_url_count": len(sitemap_urls),
        "sitemap_html_mismatch_count": sitemap_mismatch,
        "robots_exists": (output / "robots.txt").is_file(),
        "title_error_count": sum(bool(row["title_error"]) for row in rows),
        "description_error_count": sum(bool(row["description_error"]) for row in rows),
        "canonical_error_count": sum(bool(row["canonical_error"]) for row in rows),
        "canonical_duplicate_count": len(canonical_values) - len(set(canonical_values)),
        "h1_error_count": sum(bool(row["h1_error"]) for row in rows),
        "jsonld_error_count": sum(int(row["jsonld_error"]) for row in rows),
        "broken_internal_link_count": broken_links,
        "broken_image_count": broken_images,
        "css_reference_error_count": css_errors,
        "js_reference_error_count": js_errors,
        "orphan_page_count": orphan_count,
        "unreachable_from_home_count": unreachable,
        "home_max_click_depth": max(depths.values(), default=0),
        "body_ge_95": (
            similarity["body_ge_95"] if similarity
            else source_report["similarity_summary"].get("body_ge_95", 0)
        ),
        "body_90_to_95": (
            similarity["body_90_to_95"] if similarity
            else source_report["body_90_to_95"]
        ),
        "body_80_to_90": (
            similarity["body_80_to_90"] if similarity
            else source_report["body_80_to_90"]
        ),
        "exact_body": (
            similarity["exact_body_duplicate_count"] if similarity
            else source_report["duplicate_summary"]["duplicate_body_count"]
        ),
        "exact_normalized_body": (
            similarity["exact_normalized_body_duplicate_count"] if similarity
            else source_report["duplicate_summary"][
                "duplicate_normalized_body_count"
            ]
        ),
        "simple_replacement": (
            similarity["simple_replacement_count"] if similarity
            else source_report["similarity_summary"][
                "simple_replacement_pair_count"
            ]
        ),
        "slug_duplicate_count": source_report["duplicate_summary"][
            "duplicate_slug_count"
        ],
        "url_duplicate_count": source_report["duplicate_summary"][
            "duplicate_URL_count"
        ],
    }
    blockers = (
        "body_ge_95", "exact_body", "exact_normalized_body",
        "simple_replacement", "slug_duplicate_count", "url_duplicate_count",
        "canonical_duplicate_count", "sitemap_html_mismatch_count",
        "canonical_error_count", "h1_error_count", "jsonld_error_count",
        "broken_internal_link_count", "broken_image_count",
        "css_reference_error_count", "js_reference_error_count",
        "orphan_page_count", "unreachable_from_home_count",
        "corrected_page_consistency_error_count", "old_name_residual_count",
        "question_mark_slug_count",
    )
    summary["passed"] = all(summary[key] == 0 for key in blockers)
    summary["preview_started"] = False
    report = (
        Path(args.output_report).resolve()
        if args.output_report
        else REPORTS / "production_full_candidate_final_audit.json"
    )
    report.parent.mkdir(parents=True, exist_ok=True)
    report.write_text(
        json.dumps(summary, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if summary["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
