#!/usr/bin/env python3
"""
ITS (Interrupted Time Series) analysis for antibiotics injection data.

Input : output/rd10_injection_monthly_AB_only.xlsx  (FY2023: Apr 2023 – Mar 2024)
        output/rd11_injection_monthly_AB_only.xlsx  (FY2024: Apr 2024 – Mar 2025)
Output: output/its/its_ab_{code}.png  per category
        output/its/its_ab_summary.xlsx

Model (segmented regression):
  Y_t = β0 + β1·t + β2·D + β3·(t − T₀)·D + ε
  t   : 1 .. 24 (t=1 → 2023-04, t=24 → 2025-03)
  D   : 0 pre-intervention, 1 post-intervention
  T₀  : intervention month index (default 13 → 2024-04, 診療報酬改定)

Usage:
  uv run scripts/run_its.py
  uv run scripts/run_its.py --rd10 output/rd10_injection_monthly_AB_only.xlsx \\
                             --rd11 output/rd11_injection_monthly_AB_only.xlsx \\
                             --intervention 13 --output output/its
"""
# /// script
# requires-python = ">=3.11"
# dependencies = [
#   "openpyxl>=3.1",
#   "statsmodels>=0.14",
#   "matplotlib>=3.8",
#   "numpy>=1.26",
#   "pandas>=2.1",
# ]
# ///

from __future__ import annotations

import argparse
import math
from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import openpyxl
import pandas as pd
import statsmodels.api as sm

HEADER_ROWS = 4  # rows 0-3 are title / blank / header / blank
MONTH_COL_START = 9  # col index 9 = 4月, ..., col 20 = 3月 (0-indexed)
MONTHS_PER_FILE = 12

# FY2023 starts at April 2023 → t=1; FY2024 starts at April 2024 → t=13
RD10_LABEL = "rd10"
RD11_LABEL = "rd11"

CATEGORY_NAMES = {
    611: "グラム陽性菌",
    612: "グラム陰性菌",
    613: "グラム陽性・陰性菌",
    614: "グラム陽性菌・マイコプラズマ",
    615: "嫌気性菌",
    616: "抗酸菌",
    617: "真菌（抗真菌）",
    619: "その他抗生物質",
    621: "サルファ剤",
    622: "抗結核剤",
    623: "抗ハンセン病剤",
    624: "合成抗菌剤",
    629: "その他化学療法剤",
}


# ── データ読み込み ─────────────────────────────────────────────────────────────

def load_monthly_by_category(path: Path) -> dict[int, list[float | None]]:
    """
    Read an AB_only xlsx and return {category_code: [12 monthly totals]}.
    Masked '-' cells → None.
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    monthly: dict[int, list[float]] = {}
    current_code: int | None = None

    for row_idx, row in enumerate(ws.iter_rows(values_only=True)):
        if row_idx < HEADER_ROWS:
            continue

        raw_code = row[0]
        if raw_code is not None and isinstance(raw_code, (int, float)):
            current_code = int(raw_code)

        if current_code is None:
            continue

        if current_code not in monthly:
            monthly[current_code] = [0.0] * MONTHS_PER_FILE

        for m, val in enumerate(row[MONTH_COL_START: MONTH_COL_START + MONTHS_PER_FILE]):
            if isinstance(val, (int, float)):
                monthly[current_code][m] += float(val)
            # '-' or None → treated as 0 at the drug level; category total may still be valid

    return monthly


def build_timeseries(rd10_path: Path, rd11_path: Path) -> pd.DataFrame:
    """
    Combine rd10 + rd11 into a 24-row DataFrame indexed by t=1..24.
    Columns: category codes (611, 612, ...).
    """
    rd10 = load_monthly_by_category(rd10_path)
    rd11 = load_monthly_by_category(rd11_path)

    all_codes = sorted(set(rd10) | set(rd11))
    records = []
    for t in range(1, 25):
        row: dict[str, object] = {"t": t}
        if t <= 12:
            src, m = rd10, t - 1
        else:
            src, m = rd11, t - 13
        for code in all_codes:
            vals = src.get(code, [0.0] * MONTHS_PER_FILE)
            row[code] = vals[m]
        records.append(row)

    df = pd.DataFrame(records).set_index("t")
    return df


# ── ITS モデル ────────────────────────────────────────────────────────────────

def build_design_matrix(t_index: np.ndarray, T0: int) -> np.ndarray:
    """
    X = [1, t, D, (t-T0)*D]
    D = 0 if t < T0, else 1
    """
    D = (t_index >= T0).astype(float)
    slope_change = (t_index - T0) * D
    return np.column_stack([np.ones_like(t_index), t_index, D, slope_change])


def run_its(y: np.ndarray, t_index: np.ndarray, T0: int) -> dict:
    """
    Fit segmented regression. Returns coefficient dict or None if insufficient data.
    """
    mask = ~np.isnan(y)
    if mask.sum() < 6:
        return None

    X = build_design_matrix(t_index[mask], T0)
    model = sm.OLS(y[mask], X)
    result = model.fit()

    params = result.params
    ci = result.conf_int()
    pvals = result.pvalues

    return {
        "beta0": params[0], "beta1": params[1], "beta2": params[2], "beta3": params[3],
        "ci_beta0": (ci[0, 0], ci[0, 1]),
        "ci_beta1": (ci[1, 0], ci[1, 1]),
        "ci_beta2": (ci[2, 0], ci[2, 1]),
        "ci_beta3": (ci[3, 0], ci[3, 1]),
        "p_beta0": pvals[0], "p_beta1": pvals[1],
        "p_beta2": pvals[2], "p_beta3": pvals[3],
        "r2": result.rsquared,
        "n": int(mask.sum()),
        "result": result,
    }


def predict_its(fit: dict, t_index: np.ndarray, T0: int) -> np.ndarray:
    b0, b1, b2, b3 = fit["beta0"], fit["beta1"], fit["beta2"], fit["beta3"]
    D = (t_index >= T0).astype(float)
    return b0 + b1 * t_index + b2 * D + b3 * (t_index - T0) * D


# ── 月ラベル生成 ──────────────────────────────────────────────────────────────

def month_labels(n: int = 24) -> list[str]:
    """April 2023 → March 2025 (24 months)."""
    labels = []
    year, month = 2023, 4
    for _ in range(n):
        labels.append(f"{year}-{month:02d}")
        month += 1
        if month > 12:
            month = 1
            year += 1
    return labels


# ── プロット ──────────────────────────────────────────────────────────────────

def plot_its(
    code: int,
    y: np.ndarray,
    t_index: np.ndarray,
    fit: dict | None,
    T0: int,
    output_dir: Path,
    labels: list[str],
) -> None:
    fig, ax = plt.subplots(figsize=(10, 5))

    # 観測値
    ax.scatter(t_index, y, color="steelblue", s=30, zorder=3, label="観測値")

    if fit is not None:
        # 介入前フィット
        t_pre = t_index[t_index < T0]
        ax.plot(t_pre, predict_its(fit, t_pre, T0), color="steelblue", lw=2, label="フィット（介入前）")
        # 介入後フィット
        t_post = t_index[t_index >= T0]
        ax.plot(t_post, predict_its(fit, t_post, T0), color="tomato", lw=2, label="フィット（介入後）")
        # 仮想カウンターファクチュアル（介入前トレンド延長）
        b0, b1 = fit["beta0"], fit["beta1"]
        ax.plot(t_post, b0 + b1 * t_post, color="steelblue", lw=1.5,
                linestyle="--", alpha=0.6, label="介入前トレンド延長")

        # 係数注釈
        p2, p3 = fit["p_beta2"], fit["p_beta3"]
        sig2 = "***" if p2 < 0.001 else ("**" if p2 < 0.01 else ("*" if p2 < 0.05 else "n.s."))
        sig3 = "***" if p3 < 0.001 else ("**" if p3 < 0.01 else ("*" if p3 < 0.05 else "n.s."))
        annot = (
            f"β₂(水準変化)={fit['beta2']:+.0f} {sig2}\n"
            f"β₃(傾き変化)={fit['beta3']:+.1f}/月 {sig3}\n"
            f"R²={fit['r2']:.3f}"
        )
        ax.text(0.02, 0.97, annot, transform=ax.transAxes, va="top", fontsize=9,
                bbox=dict(boxstyle="round,pad=0.3", fc="white", alpha=0.8))

    # 介入線
    ax.axvline(T0 - 0.5, color="gray", lw=1.5, linestyle=":", label=f"介入点（2024年4月）")

    # 軸設定
    tick_indices = t_index[::2]  # 2ヶ月ごと
    ax.set_xticks(tick_indices)
    ax.set_xticklabels([labels[int(i) - 1] for i in tick_indices], rotation=45, ha="right", fontsize=8)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:,.0f}"))

    name = CATEGORY_NAMES.get(code, str(code))
    ax.set_title(f"注射薬 外来院内・抗菌薬 ITS解析\n薬効{code}：{name}", fontsize=12)
    ax.set_xlabel("診療年月")
    ax.set_ylabel("月次処方数量（管/瓶）")
    ax.legend(fontsize=8, loc="upper left", bbox_to_anchor=(0.0, 0.88))
    ax.grid(axis="y", alpha=0.3)

    plt.tight_layout()
    out_path = output_dir / f"its_ab_{code}.png"
    fig.savefig(out_path, dpi=150)
    plt.close(fig)
    print(f"  Saved: {out_path}")


# ── Excel 出力 ────────────────────────────────────────────────────────────────

def export_summary(results: list[dict], output_dir: Path) -> None:
    rows = []
    for r in results:
        code = r["code"]
        name = CATEGORY_NAMES.get(code, "")
        fit = r["fit"]
        if fit is None:
            rows.append({"薬効分類": code, "薬効分類名称": name, "備考": "データ不足"})
            continue
        rows.append({
            "薬効分類": code,
            "薬効分類名称": name,
            "β0（切片）": fit["beta0"],
            "β1（介入前傾き/月）": fit["beta1"],
            "β2（水準変化）": fit["beta2"],
            "β2_95CI_下": fit["ci_beta2"][0],
            "β2_95CI_上": fit["ci_beta2"][1],
            "β2_p値": fit["p_beta2"],
            "β3（傾き変化/月）": fit["beta3"],
            "β3_95CI_下": fit["ci_beta3"][0],
            "β3_95CI_上": fit["ci_beta3"][1],
            "β3_p値": fit["p_beta3"],
            "R²": fit["r2"],
            "n": fit["n"],
        })

    df = pd.DataFrame(rows)
    out_path = output_dir / "its_ab_summary.xlsx"
    df.to_excel(out_path, index=False)
    print(f"  Saved: {out_path}")


# ── メイン ────────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(description="ITS analysis for antibiotics injection data")
    parser.add_argument("--rd10", default="output/rd10_injection_monthly_AB_only.xlsx")
    parser.add_argument("--rd11", default="output/rd11_injection_monthly_AB_only.xlsx")
    parser.add_argument("--intervention", type=int, default=13,
                        help="Intervention month index t (1-indexed; default 13 = April 2024)")
    parser.add_argument("--output", default="output/its")
    args = parser.parse_args()

    rd10_path = Path(args.rd10)
    rd11_path = Path(args.rd11)
    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)
    T0 = args.intervention

    print(f"Loading data...")
    df = build_timeseries(rd10_path, rd11_path)

    t_index = np.array(df.index, dtype=float)
    labels = month_labels()

    intervention_label = month_labels()[T0 - 1]
    print(f"Running ITS (T₀={T0}, intervention={intervention_label})...")
    all_results = []

    for code in sorted(df.columns):
        y = df[code].to_numpy(dtype=float)
        # 全ゼロ列はスキップ
        if np.nansum(y) == 0:
            print(f"  Skip {code} (all zero)")
            continue

        fit = run_its(y, t_index, T0)
        if fit is None:
            print(f"  Skip {code} (insufficient data)")
        else:
            name = CATEGORY_NAMES.get(code, str(code))
            sig2 = "p<0.05" if fit["p_beta2"] < 0.05 else "n.s."
            sig3 = "p<0.05" if fit["p_beta3"] < 0.05 else "n.s."
            print(f"  {code} {name}: β2={fit['beta2']:+.0f}({sig2}), β3={fit['beta3']:+.1f}/月({sig3}), R²={fit['r2']:.3f}")

        plot_its(code, y, t_index, fit, T0, output_dir, labels)
        all_results.append({"code": code, "fit": fit})

    print(f"\nExporting summary...")
    export_summary(all_results, output_dir)
    print("Done.")


if __name__ == "__main__":
    main()
