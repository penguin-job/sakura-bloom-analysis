import sys
import requests
import pandas as pd
from pathlib import Path

# 引数
location = sys.argv[1]
lat = float(sys.argv[2])
lon = float(sys.argv[3])
output_base_dir = Path(sys.argv[4])

# 出力先
output_dir = output_base_dir / "data"
output_dir.mkdir(parents=True, exist_ok=True)

# API
url = "https://api.open-meteo.com/v1/forecast"

params = {
    "latitude": lat,
    "longitude": lon,
    "daily": "temperature_2m_mean,temperature_2m_max,temperature_2m_min",
    "forecast_days": 16,
    "timezone": "Asia/Tokyo"
}

response = requests.get(url, params=params, timeout=30)
response.raise_for_status()

data = response.json()

# DataFrame
df = pd.DataFrame({
    "date": data["daily"]["time"],
    "location": location,
    "avg_temp_forecast": data["daily"]["temperature_2m_mean"],
    "max_temp_forecast": data["daily"]["temperature_2m_max"],
    "min_temp_forecast": data["daily"]["temperature_2m_min"]
})

# CSV
safe_location = str(location).replace(" ", "_")
output_path = output_dir / f"weather_forecast_{safe_location}.csv"

df.to_csv(output_path, index=False, encoding="utf-8-sig")

print("予報取得完了")
print(output_path)