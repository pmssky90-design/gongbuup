from __future__ import annotations

import sys
from pathlib import Path


SCRIPTS_DIR = Path(__file__).resolve().parent / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

from build import main as legacy_main  # noqa: E402


QUALITY_MODES = {
    "plan", "preview", "production-priority-1", "production-priority-1-2",
    "production-priority-1-3", "production-all",
}


def main() -> int:
    if "--mode" not in sys.argv:
        return legacy_main()
    index = sys.argv.index("--mode")
    if index + 1 >= len(sys.argv) or sys.argv[index + 1] not in QUALITY_MODES:
        print("지원 모드: " + ", ".join(sorted(QUALITY_MODES)))
        return 2
    mode = sys.argv[index + 1]
    if mode.startswith("production-"):
        from school_region_quality_review import INPUT_DEFAULT, priority_reports
        import json
        from build_school_region_preview import (
            SETTINGS_PATH, build_school_records, generation_phase, plan_pages, read_excel,
        )
        settings = json.loads(SETTINGS_PATH.read_text(encoding="utf-8-sig"))
        config = settings["school_region_generation"]
        links, regions, _, _ = read_excel(INPUT_DEFAULT)
        plans, _, _ = plan_pages(regions, links, build_school_records(links), config)
        plans = [p for p in plans if generation_phase(p) in set(config["enabled_generation_phases"])]
        report = json.loads((Path(__file__).parent / "reports" / "page_plan_summary.json").read_text(encoding="utf-8"))
        scope = priority_reports(plans, report["estimated_full_output_size_bytes"], report["estimated_full_generation_seconds"])
        selected = scope["modes"][mode]
        print(json.dumps({
            "mode": mode,
            "implemented": True,
            "executed": False,
            "site_modified": False,
            **selected,
            "note": "안전 규칙에 따라 production 모드는 이번 작업에서 계획만 보고하고 실행하지 않았습니다.",
        }, ensure_ascii=False, indent=2))
        return 0
    if mode == "plan":
        from full_plan_validation import main as full_plan_main
        return full_plan_main()
    from school_region_quality_review import main as quality_main
    forwarded = ["--mode", mode]
    return quality_main(forwarded)


if __name__ == "__main__":
    raise SystemExit(main())
