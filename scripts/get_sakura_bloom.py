from __future__ import annotations

import csv
import re
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional

import requests


PAGES = [
    (2001, 2010, "https://www.data.jma.go.jp/sakura/data/sakura003_05.html"),
    (2011, 2020, "https://www.data.jma.go.jp/sakura/data/sakura003_06.html"),
    (2021, 2025, "https://www.data.jma.go.jp/sakura/data/sakura003_07.html"),
]

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "data"
OUTPUT_DIR.mkdir(exist_ok=True)
OUTPUT_CSV = OUTPUT_DIR / "sakura_bloom_all.csv"


@dataclass
class BloomRecord:
    location: str
    year: int
    bloom_date: Optional[str]   # YYYY/MM/DD or None
    normal_date: Optional[str]  # M月D日 or None
    substitute_species: str


def to_date_str(year: int, month: str, day: str) -> Optional[str]:
    if month == "-" or day == "-":
        return None
    return f"{year}/{int(month)}/{int(day)}"


def to_jp_date(month: str, day: str) -> Optional[str]:
    if month == "-" or day == "-":
        return None
    return f"{int(month)}月{int(day)}日"


def build_pattern(years: List[int]) -> re.Pattern:
    pattern = r"^(?P<location>\S+)\s+(?:\*\s+)?"

    for year in years:
        pattern += rf"(?P<y{year}_m>-|\d+)\s+(?P<y{year}_d>-|\d+)\s+"

    pattern += r"(?P<normal_m>-|\d+)\s+(?P<normal_d>-|\d+)"
    pattern += r"(?:\s+(?P<substitute>.+))?$"

    return re.compile(pattern)


def parse_sakura_text(text: str, years: List[int]) -> List[BloomRecord]:
    records: List[BloomRecord] = []
    pattern = build_pattern(years)

    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue

        if (
            "地点名" in line
            or "月 日" in line
            or line.startswith("#")
            or line.startswith("「")
            or line.startswith("各種データ・資料")
            or line.startswith("ホーム")
            or line.startswith("このページのトップ")
            or ("（注）" in line and line.count(" ") < 3)
        ):
            continue

        m = pattern.match(line)
        if not m:
            continue

        location = m.group("location")
        substitute = (m.group("substitute") or "").strip()

        for year in years:
            month = m.group(f"y{year}_m")
            day = m.group(f"y{year}_d")

            records.append(
                BloomRecord(
                    location=location,
                    year=year,
                    bloom_date=to_date_str(year, month, day),
                    normal_date=to_jp_date(m.group("normal_m"), m.group("normal_d")),
                    substitute_species=substitute,
                )
            )

    return records


def fetch_page_text(url: str) -> str:
    response = requests.get(url, timeout=30)
    response.raise_for_status()
    response.encoding = response.apparent_encoding
    return response.text


def main() -> None:
    all_records: List[BloomRecord] = []

    for start_year, end_year, url in PAGES:
        years = list(range(start_year, end_year + 1))
        text = fetch_page_text(url)
        records = parse_sakura_text(text, years)
        all_records.extend(records)
        print(f"{start_year}-{end_year}: {len(records)} rows")

    with OUTPUT_CSV.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(
            ["location", "year", "bloom_date", "normal_date", "substitute_species"]
        )

        for r in all_records:
            writer.writerow(
                [
                    r.location,
                    r.year,
                    r.bloom_date or "",
                    r.normal_date or "",
                    r.substitute_species,
                ]
            )

    print(f"saved: {OUTPUT_CSV}")
    print(f"total rows: {len(all_records)}")


if __name__ == "__main__":
    main()