from __future__ import annotations

import csv
import json
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SOURCE = (
    ROOT / "reports" / "production_full_repair" / "repair_content_targets.csv"
)
TARGET = ROOT / "config" / "production_full_repair_variants.json"


def main() -> int:
    with SOURCE.open("r", encoding="utf-8-sig", newline="") as source:
        variants = {
            str(row["slug"]): int(row["proposed_variant_id"])
            for row in csv.DictReader(source)
        }
    TARGET.write_text(
        json.dumps(dict(sorted(variants.items())), ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(json.dumps({
        "source": str(SOURCE),
        "target": str(TARGET),
        "variant_count": len(variants),
    }, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
