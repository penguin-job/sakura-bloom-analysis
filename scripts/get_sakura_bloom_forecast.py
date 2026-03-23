import csv
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional
from datetime import datetime

import requests
from bs4 import BeautifulSoup


FORECAST_URL = "https://tenki.jp/sakura/expectation/"

OUTPUT_DIR = Path(__file__).resolve().parent / "data"
OUTPUT_DIR.mkdir(exist_ok=True)
OUTPUT_CSV = OUTPUT_DIR / "sakura_forecast.csv"


@dataclass
class BloomRecord:
    location: str
    year: int
    bloom_date: Optional[str]
    normal_date: Optional[str]
    substitute_species: str


def fetch_forecast() -> List[BloomRecord]:
    year = datetime.now().year

    res = requests.get(FORECAST_URL, timeout=30)
    res.raise_for_status()

    soup = BeautifulSoup(res.text, "html.parser")

    records: List[BloomRecord] = []

    # ★ ここが重要（クラスは変わる可能性あり）
    rows = soup.select("table tr")

    for row in rows:
        cols = row.find_all("td")
        if len(cols) < 2:
            continue

        location = cols[0].get_text(strip=True)
        raw_date = cols[1].get_text(strip=True)

        # 都市補正
        if location == "さいたま":
            location = "熊谷"

        bloom_date = None
        try:
            dt = datetime.strptime(f"{year}年{raw_date}", "%Y年%m月%d日")
            bloom_date = dt.strftime("%Y/%m/%d")
        except:
            pass

        records.append(
            BloomRecord(
                location=location,
                year=year,
                bloom_date=bloom_date,
                normal_date=None,
                substitute_species=""
            )
        )

    return records


def save_csv(records: List[BloomRecord]) -> None:
    with OUTPUT_CSV.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(
            ["location", "year", "bloom_date", "normal_date", "substitute_species"]
        )

        for r in records:
            writer.writerow(
                [
                    r.location,
                    r.year,
                    r.bloom_date or "",
                    "",
                    "",
                ]
            )


def main():
    records = fetch_forecast()
    save_csv(records)

    print(f"saved: {OUTPUT_CSV}")
    print(f"rows: {len(records)}")


if __name__ == "__main__":
    main()