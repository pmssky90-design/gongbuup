from __future__ import annotations

import csv
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from build_school_region_preview import stable_int  # noqa: E402

PRIMARY = ROOT / "config" / "priority1_blocker_variants.json"
ADDITIONAL = ROOT / "config" / "priority1_candidate_additional_variants.json"
HIGH = ROOT / "reports" / "production_priority_1_candidate_similarity_95.csv"


def main() -> int:
    primary = set(json.loads(PRIMARY.read_text(encoding="utf-8")))
    with HIGH.open(encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.DictReader(handle))
    pages = sorted({
        slug
        for row in rows
        for slug in (row["page_a"], row["page_b"])
        if slug not in primary
    })
    mapping = {
        slug: stable_int(f"{slug}:priority1-candidate-additional") % 50 + 1
        for slug in pages
    }
    ADDITIONAL.write_text(
        json.dumps(mapping, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(json.dumps({
        "source_high_pair_count": len(rows),
        "primary_mapping_preserved_count": len(primary),
        "additional_mapping_count": len(mapping),
        "overlap_count": len(primary & set(mapping)),
    }, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
