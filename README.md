# Sakura Bloom Analysis

桜の開花と気温の関係を分析したプロジェクトです。
気象庁のサイトから気温データを取得し、2月1日からの最高気温の累積を算出して開花日との関係性を可視化しました。

---

## 概要

桜の開花は毎年少しずつ異なります。「開花600度の法則」がどれくらい当てはまるのか、
また、その違いがどのような要因で生まれるのかを知るために、気温データを用いて分析を行いました。

本プロジェクトでは、過去の気温データを取得し、累積気温と開花日の関係を比較しています。
また、今年に関しては実績と公開されている予想データを組み合わせて、日別の推移を可視化しました。

---

## 構成図（全体像）

ポートフォリオ全体の構成については、以下のPDFをご覧ください。

▶ [システム構成図（PDF）]([./sys-01_portfolio.pdf](https://github.com/penguin-job/sakura-bloom-analysis/blob/main/sys-03_portfolio.pdf))

---

## 処理の流れ

1. 気象データの取得
2. 日別気温の整理
3. 累積気温の計算
4. 開花日データとの統合
5. グラフによる可視化

---

## 使用技術

* Excel
* VBA
* Python（データ取得・補助処理）

---

## 分析内容

* 累積気温と開花日の関係性の確認
* 年ごとの変動の比較
* 傾向の可視化（グラフ化）

---

## リポジトリ構成

```plaintext
.
├─ excel-vba/                # Excel VBAモジュール
│  ├─ mod00_Config.bas       # 設定（パス・定数など）
│  ├─ mod01_Main.bas         # メイン処理
│  ├─ mod02_Input.bas        # データ取得処理
│  ├─ mod03_Process.bas      # データ加工処理
│  ├─ mod04_Calc.bas         # 開花計算ロジック
│  └─ mod09_Utility.bas      # 共通処理
│
├─ sample-data/              # サンプルデータ（東京）
│  ├─ sakura_bloom_all.csv
│  ├─ weather_forecast_東京.csv
│  ├─ weather_東京_2020_2025.csv
│  └─ weather_東京_2026.csv
│
├─ scripts/                  # Pythonスクリプト
│  ├─ get_sakura_bloom.py
│  ├─ get_sakura_bloom_forecast.py
│  ├─ get_weather_current_year.py
│  ├─ get_weather_forcast.py
│  └─ get_weather_past_years.py
│
├─ .gitignore
├─ README.md
└ sys-03_portfolio.pdf          # システム構成図
```
