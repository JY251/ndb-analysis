# extract-ab-monthly

NDB注射薬（月別薬効分類別）データから抗菌薬行を抽出し、2種のExcelを生成するスキル。

## 対象薬効分類コード

| コード | 分類 | 備考 |
|---|---|---|
| 611–615 | 抗生物質製剤各種 | 含む |
| 616 | 抗酸菌 | 含む |
| 617 | 抗真菌剤 | 含む（スライドから変更） |
| 619 | その他抗生物質製剤 | 含む |
| 621–624 | サルファ剤・抗結核・抗ハンセン・合成抗菌 | 含む |
| 625 | 抗ウイルス剤 | **除外** |
| 629 | その他化学療法剤 | 含む（ただしサムチレール＝Saquinavir、コード622136101のみ除外） |

コードを変えたい場合は `its_no10_11/scripts/extract_antibiotics.py` の `ANTIBIOTIC_CODES` および `EXCLUDE_629_DRUG_CODES` を編集する。

## 出力ファイル

| ファイル名 | 内容 |
|---|---|
| `{label}_injection_monthly_highlighted.xlsx` | 元データの抗菌薬行を黄色ハイライト |
| `{label}_injection_monthly_AB_only.xlsx` | 抗菌薬行のみ抽出（列A・B前方補完） |

## 実行手順

```bash
cd /mnt/c/Users/e231b/WorkingDir/ndb_analysis/its_no10_11

# No.10 (FY2023)
uv run scripts/extract_antibiotics.py \
  data/no10_raw/injection_monthly.xlsx \
  output \
  --label rd10

# No.11 (FY2024)
uv run scripts/extract_antibiotics.py \
  "/mnt/c/Users/e231b/Downloads/001711931/05_処方薬_シート統合後(公開対象)/01_処方薬（内服／外用／注射）全/【注射】診療月別薬効分類別数量(公費含む).xlsx" \
  output \
  --label rd11
```

## データソース

| ラベル | 年度 | 期間 | 入手先 |
|---|---|---|---|
| rd10 | FY2023 | 2023年4月〜2024年3月 | `data/no10_raw/injection_monthly.xlsx`（MHLW DL済） |
| rd11 | FY2024 | 2024年4月〜2025年3月 | `C:\Users\e231b\Downloads\001711931\...` ローカル |

## 注意

- 629 injection では全薬剤がフルコナゾール（抗真菌）のため全件含まれる
- 629 oral ではサムチレール（Saquinavir）のみ除外（残りは抗菌薬または抗真菌薬）
- スキプトは `uv` で依存パッケージを自動管理（Docker不要で単体実行可能）
