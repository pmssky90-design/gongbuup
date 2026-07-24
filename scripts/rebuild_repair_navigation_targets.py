from __future__ import annotations

import csv
import html
import json
import re
import sys
from collections import Counter, defaultdict, deque
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from urllib.parse import quote, unquote, urlsplit

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


SOURCE = ROOT / "preview" / "production-full-candidate_20260724_151919"
REPORT_DIR = ROOT / "reports" / "production_full_repair"
TARGET = REPORT_DIR / "repair_navigation_targets.csv"
HREF_RE = re.compile(r'<a\b[^>]+href=["\']([^"\']+)["\']', re.I)


def url_path(file: Path) -> str:
    relative = file.relative_to(SOURCE)
    if relative == Path("index.html"):
        return "/"
    return "/" + "/".join(
        quote(part, safe="") for part in relative.parent.parts
    ) + "/"


def normalize(value: str) -> str:
    parsed = urlsplit(html.unescape(value))
    if parsed.scheme or parsed.netloc:
        return ""
    parts = [
        quote(unquote(part), safe="")
        for part in parsed.path.strip("/").split("/") if part
    ]
    return "/" if not parts else "/" + "/".join(parts) + (
        "/" if parsed.path.endswith("/") else ""
    )


def inspect(file: Path) -> tuple[str, set[str]]:
    text = file.read_text(encoding="utf-8")
    return url_path(file), {
        target for target in (normalize(value) for value in HREF_RE.findall(text))
        if target
    }


def main() -> int:
    with TARGET.open("r", encoding="utf-8-sig", newline="") as source:
        proposed = {
            row["url"]: row["proposed_parent_link"]
            for row in csv.DictReader(source)
        }
    files = list(SOURCE.rglob("*.html"))
    paths = {url_path(file) for file in files}
    with ThreadPoolExecutor(max_workers=8) as pool:
        inspected = list(pool.map(inspect, files, chunksize=64))
    graph: dict[str, set[str]] = defaultdict(set)
    incoming = Counter()
    for source, targets in inspected:
        for target in targets:
            if target in paths and target != source:
                graph[source].add(target)
                incoming[target] += 1
    depths = {"/": 0}
    queue = deque(["/"])
    while queue:
        current = queue.popleft()
        for target in graph.get(current, set()):
            if target not in depths:
                depths[target] = depths[current] + 1
                queue.append(target)

    settings = json.loads(SETTINGS_PATH.read_text(encoding="utf-8-sig"))
    config = settings["school_region_generation"]
    links, regions, _, _ = read_excel(INPUT_DEFAULT)
    schools = build_school_records(links)
    plans, _, _ = plan_pages(regions, links, schools, config)
    enabled = set(str(value) for value in config["enabled_generation_phases"])
    plan_by_slug = {
        str(row["slug"]): row
        for row in plans if generation_phase(row) in enabled
    }
    rows = []
    for path in sorted(paths):
        if path == "/":
            continue
        orphan = incoming[path] == 0
        unreachable = path not in depths
        if not orphan and not unreachable:
            continue
        slug = unquote(path.strip("/").split("/")[-1])
        plan = plan_by_slug.get(slug, {})
        rows.append({
            "url": path,
            "slug": slug,
            "page_type": plan.get("page_type", "navigation"),
            "region": plan.get("지역명", ""),
            "parent_hub": "",
            "current_depth": "unreachable" if unreachable else depths[path],
            "proposed_parent_link": proposed.get(path, ""),
        })
    fields = (
        "url", "slug", "page_type", "region", "parent_hub",
        "current_depth", "proposed_parent_link",
    )
    with TARGET.open("w", encoding="utf-8-sig", newline="") as target:
        writer = csv.DictWriter(target, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)
    summary = {
        "source_candidate": str(SOURCE),
        "navigation_repair_target_count": len(rows),
        "original_orphan_count": sum(
            path != "/" and incoming[path] == 0 for path in paths
        ),
        "original_unreachable_count": sum(path not in depths for path in paths),
        "original_max_reachable_depth": max(depths.values(), default=0),
    }
    (REPORT_DIR / "repair_navigation_target_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
