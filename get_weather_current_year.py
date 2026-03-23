import sys
import requests
import pandas as pd
from pathlib import Path
from datetime import date

# 引数受け取り
location = sys.argv[1]
lat = float(sys.argv[2])
lon = float(sys.argv[3])
year_from = int(sys.argv[4])
year_to = int(sys.argv[5])   # 形式をそろえるため受け取るが、今年用では使わない
output_base_dir = Path(sys.argv[6])

# 出力先
output_dir = output_base_dir / "data"
output_dir.mkdir(parents=True, exist_ok=True)

# 期間（今年の2/1～今日）
current_year = date.today().year
start_date = f"{current_year}-02-01"
end_date = date.today().isoformat()

# API取得
url = "https://archive-api.open-meteo.com/v1/archive"
params = {
    "latitude": lat,
    "longitude": lon,
    "start_date": start_date,
    "end_date": end_date,
    "daily": "temperature_2m_mean,temperature_2m_max,temperature_2m_min",
    "timezone": "Asia/Tokyo"
}

response = requests.get(url, params=params, timeout=30)
response.raise_for_status()

data = response.json()

if "daily" not in data:
    raise ValueError(f"API応答が想定と異なります: {data}")

# DataFrame化
df = pd.DataFrame({
    "date": data["daily"]["time"],
    "location": location,
    "avg_temp": data["daily"]["temperature_2m_mean"],
    "max_temp": data["daily"]["temperature_2m_max"],
    "min_temp": data["daily"]["temperature_2m_min"]
})

# CSV保存
safe_location = str(location).replace(" ", "_")
output_path = output_dir / f"weather_{safe_location}_{current_year}.csv"

df.to_csv(output_path, index=False, encoding="utf-8-sig")

print("取得完了")
print(f"地点: {location}")
print(f"緯度: {lat}")
print(f"経度: {lon}")
print(f"期間: {start_date} ～ {end_date}")
print(f"保存先: {output_path}")