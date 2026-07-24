from __future__ import annotations

import csv
import hashlib
import json
import sys
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

ROOT = Path(__file__).resolve().parents[1]
RULES_PATH = ROOT / "config" / "search_intent_rules.json"
SETTINGS_PATH = ROOT / "config" / "settings.json"
OUTPUT = ROOT / "reports" / "search_intent_engine"


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open(encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def write_csv(name: str, fields: tuple[str, ...], rows: list[dict[str, object]]) -> None:
    path = OUTPUT / name
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def stable_int(value: str) -> int:
    return int(hashlib.sha256(value.encode("utf-8")).hexdigest(), 16)


def first_token(text: str, groups: dict[str, list[str]], default: str) -> str:
    for label, tokens in groups.items():
        if any(token in text for token in tokens):
            return label
    return default


def signals_for(row: dict[str, str], rules: dict[str, object]) -> dict[str, object]:
    text = f"{row['keyword']} {row['page_type']} {row['scope']}"
    tokens = rules["signal_tokens"]
    signals: dict[str, object] = {
        "grade_band": first_token(text, tokens["grade_band"], "UNSPECIFIED"),
        "subject": first_token(text, tokens["subject"], "GENERAL"),
        "school_based": row["scope"] == "school",
        "directory": row["scope"] == "directory",
    }
    for name in (
        "exam", "midterm", "final", "low_score", "middle_score",
        "high_score", "parents", "self_directed", "consultation",
        "comparison", "faq",
    ):
        signals[name] = any(token.lower() in text.lower() for token in tokens[name])
    return signals


def matches(
    rule: dict[str, object], row: dict[str, str], signals: dict[str, object]
) -> bool:
    if "when_signal" in rule and not bool(signals[str(rule["when_signal"])]):
        return False
    if "page_types" in rule and row["page_type"] not in rule["page_types"]:
        return False
    if "page_type_contains" in rule and str(rule["page_type_contains"]) not in row["page_type"]:
        return False
    if "grade_bands" in rule and signals["grade_band"] not in rule["grade_bands"]:
        return False
    return True


def classify(
    row: dict[str, str], rules: dict[str, object]
) -> tuple[str, str, dict[str, object], int, list[str]]:
    signals = signals_for(row, rules)
    chosen_rule = None
    for rule in rules["decision_rules"]:
        if matches(rule, row, signals):
            chosen_rule = rule
            break
    if chosen_rule is None:
        raise RuntimeError(f"intent rule not found: {row['slug']}")
    seed = stable_int(
        f"{rules['seed_namespace']}:{row['slug']}:{chosen_rule['id']}"
    )
    intents = list(chosen_rule["intents"])
    intent = intents[seed % len(intents)]
    orders = rules["intent_blocks"][intent]["orders"]
    order = list(orders[(seed // max(1, len(intents))) % len(orders)])
    return intent, str(chosen_rule["id"]), signals, seed, order


def main() -> int:
    rules = json.loads(RULES_PATH.read_text(encoding="utf-8"))
    settings = json.loads(SETTINGS_PATH.read_text(encoding="utf-8-sig"))
    source = ROOT / str(rules["source_plan"])
    source_before = hashlib.sha256(source.read_bytes()).hexdigest()
    rows = read_csv(source)
    signatures = {
        row["slug"]: row
        for row in read_csv(ROOT / "reports" / "full_plan_page_signatures.csv")
    }
    OUTPUT.mkdir(parents=True, exist_ok=True)

    mapped: list[dict[str, object]] = []
    intent_counts: Counter[str] = Counter()
    by_page_type: Counter[tuple[str, str]] = Counter()
    by_priority: Counter[tuple[str, str]] = Counter()
    rule_counts: Counter[str] = Counter()
    block_order_counts: Counter[tuple[str, str]] = Counter()
    site_url = str(settings["site_url"]).rstrip("/")
    category = quote(str(settings["page_category"]), safe="")
    for page_id, row in enumerate(rows, start=1):
        intent, rule_id, signals, seed, block_order = classify(row, rules)
        path = f"/{category}/{quote(row['slug'], safe='')}/"
        signature = signatures.get(row["slug"], {})
        mapped_row = {
            "page_id": page_id,
            "priority": row["priority"],
            "page_type": row["page_type"],
            "scope": row["scope"],
            "keyword": row["keyword"],
            "slug": row["slug"],
            "url_path": path,
            "canonical": site_url + path,
            "generation_phase": row["generation_phase"],
            "grade_band": signals["grade_band"],
            "subject": signals["subject"],
            "has_exam_signal": signals["exam"],
            "school_based": signals["school_based"],
            "intent_type": intent,
            "rule_id": rule_id,
            "selection_seed_sha256": hashlib.sha256(
                f"{rules['seed_namespace']}:{row['slug']}:{rule_id}".encode("utf-8")
            ).hexdigest(),
            "block_order": ">".join(block_order),
            "title_sha256": signature.get("title_sha256", ""),
            "description_sha256": signature.get("description_sha256", ""),
        }
        mapped.append(mapped_row)
        intent_counts[intent] += 1
        by_page_type[(row["page_type"], intent)] += 1
        by_priority[(row["priority"], intent)] += 1
        rule_counts[rule_id] += 1
        block_order_counts[(intent, ">".join(block_order))] += 1

    fields = (
        "page_id", "priority", "page_type", "scope", "keyword", "slug",
        "url_path", "canonical", "generation_phase", "grade_band", "subject",
        "has_exam_signal", "school_based", "intent_type", "rule_id",
        "selection_seed_sha256", "block_order", "title_sha256",
        "description_sha256",
    )
    write_csv("intent_page_mapping.csv", fields, mapped)
    write_csv(
        "intent_statistics.csv",
        ("intent_type", "page_count", "share_percent"),
        [{
            "intent_type": intent,
            "page_count": intent_counts[intent],
            "share_percent": round(intent_counts[intent] / len(mapped) * 100, 4),
        } for intent in rules["required_intents"]],
    )
    write_csv(
        "intent_by_page_type.csv",
        ("page_type", "intent_type", "page_count"),
        [
            {"page_type": page_type, "intent_type": intent, "page_count": count}
            for (page_type, intent), count in sorted(by_page_type.items())
        ],
    )
    write_csv(
        "intent_by_priority.csv",
        ("priority", "intent_type", "page_count"),
        [
            {"priority": priority, "intent_type": intent, "page_count": count}
            for (priority, intent), count in sorted(by_priority.items())
        ],
    )
    write_csv(
        "intent_block_orders.csv",
        ("intent_type", "order_variant", "page_count"),
        [
            {"intent_type": intent, "order_variant": order, "page_count": count}
            for (intent, order), count in sorted(block_order_counts.items())
        ],
    )
    write_csv(
        "intent_rule_usage.csv",
        ("rule_id", "page_count"),
        [{"rule_id": rule_id, "page_count": count}
         for rule_id, count in sorted(rule_counts.items())],
    )
    write_csv(
        "intent_definition_table.csv",
        ("intent_type", "block_weights", "order_variant_count", "order_variants"),
        [{
            "intent_type": intent,
            "block_weights": json.dumps(
                rules["intent_blocks"][intent]["weights"], ensure_ascii=False
            ),
            "order_variant_count": len(rules["intent_blocks"][intent]["orders"]),
            "order_variants": " || ".join(
                ">".join(order)
                for order in rules["intent_blocks"][intent]["orders"]
            ),
        } for intent in rules["required_intents"]],
    )

    source_after = hashlib.sha256(source.read_bytes()).hexdigest()
    duplicate_slug_count = len(mapped) - len({str(row["slug"]) for row in mapped})
    duplicate_canonical_count = len(mapped) - len(
        {str(row["canonical"]) for row in mapped}
    )
    missing_intent_count = sum(not row["intent_type"] for row in mapped)
    unknown_intent_count = sum(
        row["intent_type"] not in rules["required_intents"] for row in mapped
    )
    missing_required_intents = [
        intent for intent in rules["required_intents"] if intent_counts[intent] == 0
    ]
    signature_missing_count = sum(
        not row["title_sha256"] or not row["description_sha256"] for row in mapped
    )
    determinism_mismatch_count = 0
    for source_row, mapped_row in zip(rows, mapped, strict=True):
        intent, rule_id, _, _, block_order = classify(source_row, rules)
        determinism_mismatch_count += int(
            intent != mapped_row["intent_type"]
            or rule_id != mapped_row["rule_id"]
            or ">".join(block_order) != mapped_row["block_order"]
        )
    summary = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "engine_version": rules["version"],
        "source_plan": str(source),
        "source_page_count": len(rows),
        "mapped_page_count": len(mapped),
        "intent_type_count": len(intent_counts),
        "intent_counts": dict(sorted(intent_counts.items())),
        "rule_usage": dict(sorted(rule_counts.items())),
        "missing_intent_count": missing_intent_count,
        "unknown_intent_count": unknown_intent_count,
        "missing_required_intents": missing_required_intents,
        "duplicate_slug_count": duplicate_slug_count,
        "duplicate_canonical_count": duplicate_canonical_count,
        "title_description_signature_missing_count": signature_missing_count,
        "determinism_mismatch_count": determinism_mismatch_count,
        "source_sha256_before": source_before,
        "source_sha256_after": source_after,
        "source_unchanged": source_before == source_after,
        "deterministic_seed": "SHA-256(slug + rule_id + namespace)",
        "html_generated": False,
        "candidate_modified": False,
        "site_modified": False,
        "git_used": False,
        "deployed": False,
        "passed": (
            len(mapped) == len(rows)
            and not missing_intent_count
            and not unknown_intent_count
            and not duplicate_slug_count
            and not duplicate_canonical_count
            and not signature_missing_count
            and not determinism_mismatch_count
            and source_before == source_after
        ),
    }
    (OUTPUT / "search_intent_engine_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    (OUTPUT / "search_intent_invariant_validation.json").write_text(
        json.dumps({
            "source_unchanged": source_before == source_after,
            "mapped_page_count": len(mapped),
            "duplicate_slug_count": duplicate_slug_count,
            "duplicate_canonical_count": duplicate_canonical_count,
            "missing_intent_count": missing_intent_count,
            "unknown_intent_count": unknown_intent_count,
            "title_description_signature_missing_count": signature_missing_count,
            "determinism_mismatch_count": determinism_mismatch_count,
            "html_generated": False,
            "candidate_modified": False,
            "site_modified": False,
        }, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    design = f"""# Search Intent Engine v2 설계

## 범위

- 입력: `{source}`
- 전체 매핑: {len(mapped):,}페이지
- HTML·후보·site 생성 또는 수정: 없음
- URL, slug, canonical, title, description, 학교·지역 매핑: 읽기 전용

## 결정 과정

1. keyword, page_type, scope에서 학년·과목·시험·학교 신호를 추출합니다.
2. `decision_rules`를 위에서 아래로 평가해 허용 intent 풀을 선택합니다.
3. `SHA-256(namespace:slug:rule_id)`로 intent와 블록 순서를 결정합니다.
4. 같은 입력·규칙·slug는 항상 같은 결과를 만듭니다.
5. 키워드에 없는 성적대나 시험 종류는 사실처럼 추출하지 않고 허용 풀 안의 콘텐츠 관점으로만 사용합니다.

## 본문 조립 계약

- `intent_type`은 `intent_blocks`의 비율과 순서 후보를 선택합니다.
- `block_order`는 실제 v2 본문 조립 시 사용할 결정적 순서입니다.
- title과 description은 기존 해시를 그대로 참조하며 이 단계에서 재생성하지 않습니다.
- URL 계열 필드는 기존 slug와 settings.json의 site_url에서 계산만 하며 파일을 생성하지 않습니다.

## 검증

- 누락 intent: {missing_intent_count}
- 알 수 없는 intent: {unknown_intent_count}
- 중복 slug: {duplicate_slug_count}
- 중복 canonical: {duplicate_canonical_count}
- title·description 서명 누락: {signature_missing_count}
- 결정성 불일치: {determinism_mismatch_count}
- 입력 파일 변경: {source_before != source_after}
- 모든 필수 intent 사용 여부: {"통과" if not missing_required_intents else "미사용 " + ", ".join(missing_required_intents)}
"""
    (OUTPUT / "search_intent_engine_design.md").write_text(design, encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if summary["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
