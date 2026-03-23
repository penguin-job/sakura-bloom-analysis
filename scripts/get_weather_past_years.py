import sys
import requests
import pandas as pd
from pathlib import Path

# 引数受け取り
location = sys.argv[1]
lat = float(sys.argv[2])
lon = float(sys.argv[3])
year_from = int(sys.argv[4])
year_to = int(sys.argv[5])
output_base_dir = Path(sys.argv[6])

# 出力先
output_dir = output_base_dir / "data"
output_dir.mkdir(parents=True, exist_ok=True)

if year_from > year_to:
    raise ValueError("From年がTo年より大きくなっています。")

start_date = f"{year_from}-02-01"
end_date = f"{year_to}-05-31"

print(f"location = {location}")
print(f"lat = {lat}")
print(f"lon = {lon}")
print(f"start_date = {start_date}")
print(f"end_date = {end_date}")

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

print(data["daily"]["time"][:5])
print(data["daily"]["temperature_2m_mean"][:5])
print(data["daily"]["temperature_2m_max"][:5])
print(data["daily"]["temperature_2m_min"][:5])

# DataFrame化
df = pd.DataFrame({
    "date": data["daily"]["time"],
    "location": location,
    "avg_temp": data["daily"]["temperature_2m_mean"],
    "max_temp": data["daily"]["temperature_2m_max"],
    "min_temp": data["daily"]["temperature_2m_min"]
})

# 日付をdatetime型に変換
df["date"] = pd.to_datetime(df["date"])

# 2〜5月だけ抽出
df = df[df["date"].dt.month.isin([2,3,4,5])]

# CSV保存
safe_location = str(location).replace(" ", "_")
output_path = output_dir / f"weather_{safe_location}_{year_from}_{year_to}.csv"

df.to_csv(output_path, index=False, encoding="utf-8-sig")
print("取得完了")
print(f"保存先: {output_path}")