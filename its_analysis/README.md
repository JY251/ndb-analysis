# NDB Biosimilar ITS Analysis

## 概要

NDBオープンデータ（注射薬・診療月別）を用いて、**2024年度診療報酬改定**（令和6年4月）の
バイオシミラー（BS）処方量への影響をITS（中断時系列）解析で評価。

## データ

| ファイル | 内容 | 出典 |
|---------|------|------|
| `rd10_injection_monthly.xlsx` | FY2023注射薬月別（第10回NDB） | 厚生労働省NDBオープンデータ |
| `biosimilar/rd10_injection_monthly_BS_only_v2.xlsx` | FY2023 BSのみ抽出 | 上記から生成 |
| `biosimilar/rd11_injection_monthly_BS_only_v2.xlsx` | FY2024 BSのみ抽出 | No.11ローカルデータから生成 |

## 解析設定

- **介入点**: 2024年4月1日（令和6年度診療報酬改定）
  - 入院患者へのBS使用促進評価新設（バイオ後続品使用体制加算100点）
- **期間**: 2023年4月〜2025年3月（24ヶ月）
- **設定**: 外来（院内+院外合算）/ 入院
- **薬効分類**: 9区分（131, 239, 241, 243, 249, 339, 395, 399, 429）
- **アウトカム**: 月次BS処方数量（人口100,000人あたり）
- **人口**: 総務省 人口推計（e-Stat参考表）

## モデル

セグメント回帰（Segmented Regression）:

```
Y_t = β0 + β1·t + β2·D + β3·(t−T₀)·D + ε
```

- `β2`: 介入直後の水準変化（level change）
- `β3`: 介入後の傾き変化（slope change）

## スクリプト

| ファイル | 内容 |
|---------|------|
| `biosimilar/make_bs_files.py` | BSフィルター・ハイライト・抽出 |
| `biosimilar/run_its.py` | ITS解析・グラフ生成・Excel出力 |
| `biosimilar/README.md` | データ処理の詳細・注意事項 |

## 結果

`biosimilar/results/` 以下:
- `its_results.xlsx`: 全18モデルの係数・95%CI・p値
- `its_*.png`: 薬効分類別グラフ（15枚）

## 注意事項

薬効分類341（人工腎臓透析用剤）のサブラッドＢＳＧは、
製品名に「ＢＳ」を含むがバイオシミラーではないため除外。
詳細は `biosimilar/README.md` 参照。
