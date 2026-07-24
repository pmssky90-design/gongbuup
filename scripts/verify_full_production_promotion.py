from __future__ import annotations

import hashlib
import json
import re
import sys
from concurrent.futures import ThreadPoolExecutor
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
from school_region_quality_review import INPUT_DEFAULT  # noqa: E402


CANDIDATE = (
    ROOT / "preview" / "production-full-candidate-repaired_20260724_172243"
)
SITE = ROOT / "site"
REPORT = ROOT / "reports" / "full_production_promotion_integrity.json"


def digest(path: Path) -> str:
    value = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            value.update(chunk)
    return value.hexdigest()


def main() -> int:
    candidate_files = {
        path.relative_to(CANDIDATE).as_posix(): path
        for path in CANDIDATE.rglob("*") if path.is_file()
    }
    site_files = {
        path.relative_to(SITE).as_posix(): path
        for path in SITE.rglob("*") if path.is_file()
    }
    candidate_only = sorted(set(candidate_files) - set(site_files))
    site_only = sorted(set(site_files) - set(candidate_files))
    common = sorted(set(candidate_files) & set(site_files))

    def compare(relative: str) -> str | None:
        return (
            relative
            if digest(candidate_files[relative]) != digest(site_files[relative])
            else None
        )

    with ThreadPoolExecutor(max_workers=8) as pool:
        mismatches = [
            value for value in pool.map(compare, common, chunksize=32)
            if value is not None
        ]

    settings = json.loads(SETTINGS_PATH.read_text(encoding="utf-8-sig"))
    config = settings["school_region_generation"]
    links, regions, _, _ = read_excel(INPUT_DEFAULT)
    schools = build_school_records(links)
    plans, _, _ = plan_pages(regions, links, schools, config)
    enabled = set(str(value) for value in config["enabled_generation_phases"])
    plans = [row for row in plans if generation_phase(row) in enabled]
    missing_content = sum(
        not (SITE / str(settings["page_category"]) / str(row["slug"]) / "index.html").is_file()
        for row in plans
    )
    html_count = sum(path.suffix.lower() == ".html" for path in site_files.values())
    sitemap_url_count = 0
    for path in SITE.glob("sitemap-*.xml"):
        sitemap_url_count += len(re.findall(
            r"<loc>.*?</loc>", path.read_text(encoding="utf-8"), re.S
        ))
    result = {
        "candidate": str(CANDIDATE),
        "site": str(SITE),
        "candidate_file_count": len(candidate_files),
        "site_file_count": len(site_files),
        "candidate_only_file_count": len(candidate_only),
        "site_only_file_count": len(site_only),
        "sha256_mismatch_count": len(mismatches),
        "candidate_only_files": candidate_only[:100],
        "site_only_files": site_only[:100],
        "sha256_mismatches": mismatches[:100],
        "html_count": html_count,
        "content_page_count": len(plans),
        "missing_content_page_count": missing_content,
        "sitemap_url_count": sitemap_url_count,
        "robots_exists": (SITE / "robots.txt").is_file(),
    }
    result["passed"] = (
        len(candidate_files) == len(site_files)
        and not candidate_only
        and not site_only
        and not mismatches
        and html_count == 69140
        and len(plans) == 64104
        and missing_content == 0
        and sitemap_url_count == 69140
        and result["robots_exists"]
    )
    REPORT.write_text(
        json.dumps(result, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
