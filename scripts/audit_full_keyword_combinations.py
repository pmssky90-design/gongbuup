from __future__ import annotations

import csv
import json
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
FULL = DATA / "full_keyword_combinations.csv"
SAMPLE = DATA / "full_validation_sample_1000.csv"


def main() -> None:
    slugs: set[str] = set()
    titles: set[str] = set()
    audit = Counter()
    page_types = Counter()
    title_lengths: list[int] = []

    with FULL.open("r", encoding="utf-8-sig", newline="") as handle:
        for row in csv.DictReader(handle):
            audit["rows"] += 1
            page_types[row["page_type"]] += 1
            slug = row["slug"]
            title = row["title"]
            audit["duplicate_slugs"] += int(slug in slugs)
            audit["duplicate_titles"] += int(title in titles)
            slugs.add(slug)
            titles.add(title)
            title_lengths.append(len(title))
            audit["below_40"] += int(len(title) < 40)
            audit["above_68"] += int(len(title) > 68)
            terms = [term for term in row["title_core_terms"].split("|") if term]
            audit["missing_core_terms"] += int(any(term not in title for term in terms))
            elementary = row["학년표현"] in ("초등", "초등학생")
            exam = row["내신사용"].lower() == "true"
            audit["elementary_exam"] += int(elementary and exam)
            if row["page_type"] == "school":
                grade = row["학교급"]
                expression = row["학년표현"]
                mismatch = (
                    (grade == "중학교" and expression not in ("중등", "중학생"))
                    or (grade == "고등학교" and expression not in ("고등", "고등학생"))
                    or grade not in ("중학교", "고등학교")
                )
                audit["school_grade_mismatch"] += int(mismatch)

    with SAMPLE.open("r", encoding="utf-8-sig", newline="") as handle:
        sample = list(csv.DictReader(handle))
    sample_page_types = Counter(row["page_type"] for row in sample)
    sample_grades = Counter(row["학교급"] for row in sample if row["page_type"] == "school")
    sample_subjects = Counter(row["과목표현"] for row in sample if row["과목표현"])
    sample_structures = Counter(
        (
            bool(row["학년표현"]),
            bool(row["과목표현"]),
            row["내신사용"].lower() == "true",
        )
        for row in sample
    )

    namespace = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    xlsx_row_count = 0
    freeze_panes = ""
    auto_filter = ""
    adjusted_columns = 0
    with zipfile.ZipFile(DATA / "full_keyword_combinations.xlsx") as archive:
        with archive.open("xl/worksheets/sheet1.xml") as xml_file:
            for _, element in ET.iterparse(xml_file, events=("end",)):
                if element.tag == namespace + "row":
                    xlsx_row_count += 1
                elif element.tag == namespace + "pane":
                    freeze_panes = element.attrib.get("topLeftCell", "")
                elif element.tag == namespace + "autoFilter":
                    auto_filter = element.attrib.get("ref", "")
                elif element.tag == namespace + "col" and element.attrib.get("customWidth") == "1":
                    adjusted_columns += 1
                element.clear()
    xlsx = {
        "rows_without_header": xlsx_row_count - 1,
        "columns": 20,
        "freeze_panes": freeze_panes,
        "auto_filter": auto_filter,
        "adjusted_columns": adjusted_columns,
    }

    bom = {}
    for path in (
        FULL,
        DATA / "full_keyword_conflicts.csv",
        DATA / "full_keyword_invalid.csv",
        SAMPLE,
    ):
        bom[path.name] = path.read_bytes()[:3] == b"\xef\xbb\xbf"

    result = {
        "full": dict(audit),
        "page_types": dict(page_types),
        "unique_slugs": len(slugs),
        "unique_titles": len(titles),
        "title_min": min(title_lengths),
        "title_max": max(title_lengths),
        "sample_rows": len(sample),
        "sample_page_types": dict(sample_page_types),
        "sample_school_grades": dict(sample_grades),
        "sample_provinces": len({row["시도"] for row in sample}),
        "sample_region_provinces": len(
            {row["시도"] for row in sample if row["page_type"] == "region"}
        ),
        "sample_school_provinces": len(
            {row["시도"] for row in sample if row["page_type"] == "school"}
        ),
        "sample_subjects": dict(sample_subjects),
        "sample_internal_exam": sum(
            row["내신사용"].lower() == "true" for row in sample
        ),
        "sample_elementary_exam": sum(
            row["학년표현"] in ("초등", "초등학생")
            and row["내신사용"].lower() == "true"
            for row in sample
        ),
        "sample_structure_count": len(sample_structures),
        "sample_structures": {
            f"grade={grade},subject={subject},exam={exam}": count
            for (grade, subject, exam), count in sample_structures.items()
        },
        "xlsx": xlsx,
        "utf8_bom": bom,
    }
    print(json.dumps(result, ensure_ascii=True, indent=2))


if __name__ == "__main__":
    main()
