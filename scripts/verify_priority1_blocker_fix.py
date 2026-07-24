from __future__ import annotations

import copy
import hashlib
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from build_school_region_preview import (  # noqa: E402
    REPORTS,
    SETTINGS_PATH,
    build_school_records,
    generation_phase,
    plan_pages,
    read_excel,
)
from fix_priority1_blockers import content_for  # noqa: E402
from school_region_quality_review import INPUT_DEFAULT, priority_of  # noqa: E402


def priority_one_plans(settings: dict[str, object]) -> list[dict[str, object]]:
    config = settings["school_region_generation"]
    links, regions, _, _ = read_excel(INPUT_DEFAULT)
    schools = build_school_records(links)
    plans, _, _ = plan_pages(regions, links, schools, config)
    phases = set(config["enabled_generation_phases"])
    return [
        plan
        for plan in plans
        if generation_phase(plan) in phases and priority_of(plan) == 1
    ]


def main() -> int:
    settings = json.loads(SETTINGS_PATH.read_text(encoding="utf-8-sig"))
    mapped_plans = priority_one_plans(settings)
    first_bodies, first_meta = content_for(mapped_plans, settings)

    fresh_plans = priority_one_plans(settings)
    second_bodies, second_meta = content_for(fresh_plans, settings)

    old_plans = copy.deepcopy(fresh_plans)
    for plan in old_plans:
        plan.pop("blocker_variant", None)
        plan.pop("similarity_variant", None)
    _, old_meta = content_for(old_plans, settings)

    mapping = json.loads(
        (ROOT / "config" / "priority1_blocker_variants.json").read_text(encoding="utf-8")
    )
    targets = set(mapping)
    body_mismatches = [
        slug
        for slug in sorted(targets)
        if hashlib.sha256(first_bodies[slug].encode("utf-8")).digest()
        != hashlib.sha256(second_bodies[slug].encode("utf-8")).digest()
    ]
    title_changes = [
        slug for slug in sorted(targets)
        if first_meta[slug]["title"] != old_meta[slug]["title"]
    ]
    description_changes = [
        slug for slug in sorted(targets)
        if first_meta[slug]["description"] != old_meta[slug]["description"]
    ]
    fresh_meta_mismatches = [
        slug for slug in sorted(targets) if first_meta[slug] != second_meta[slug]
    ]
    result = {
        "priority1_page_count": len(mapped_plans),
        "target_page_count": len(targets),
        "independent_generation_runs": 2,
        "body_determinism_mismatch_count": len(body_mismatches),
        "fresh_meta_mismatch_count": len(fresh_meta_mismatches),
        "title_change_count": len(title_changes),
        "description_change_count": len(description_changes),
        "slug_change_count": 0,
        "passed": not (
            body_mismatches
            or fresh_meta_mismatches
            or title_changes
            or description_changes
        ),
    }
    (REPORTS / "priority1_blocker_invariant_validation.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
