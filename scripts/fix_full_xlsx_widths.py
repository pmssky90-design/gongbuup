from __future__ import annotations

import csv
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
CSV_PATH = ROOT / "data" / "full_keyword_combinations.csv"
XLSX_PATH = ROOT / "data" / "full_keyword_combinations.xlsx"


def main() -> None:
    with CSV_PATH.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.reader(handle)
        headers = next(reader)
        widths = [len(value) for value in headers]
        for row in reader:
            for index, value in enumerate(row):
                widths[index] = min(60, max(widths[index], len(value)))

    cols = ["<cols>"]
    for index, width in enumerate(widths, start=1):
        cols.append(
            f'<col min="{index}" max="{index}" width="{min(width + 3, 60)}" '
            'customWidth="1"/>'
        )
    cols.append("</cols>")
    cols_xml = "".join(cols).encode("utf-8")

    temporary = XLSX_PATH.with_suffix(".xlsx.tmp")
    with zipfile.ZipFile(XLSX_PATH, "r") as source, zipfile.ZipFile(
        temporary, "w"
    ) as target:
        for info in source.infolist():
            payload = source.read(info.filename)
            if info.filename == "xl/worksheets/sheet1.xml":
                if b"<cols>" in payload:
                    start = payload.index(b"<cols>")
                    end = payload.index(b"</cols>", start) + len(b"</cols>")
                    payload = payload[:start] + cols_xml + payload[end:]
                else:
                    payload = payload.replace(
                        b"<sheetData>", cols_xml + b"<sheetData>", 1
                    )
            target.writestr(info, payload)
    temporary.replace(XLSX_PATH)
    print(f"adjusted_columns={len(widths)}")


if __name__ == "__main__":
    main()
