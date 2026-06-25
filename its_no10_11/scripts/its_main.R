#!/usr/bin/env Rscript
# ITS (Interrupted Time Series) 解析
# 対象: NDB注射薬月別データ 抗菌薬（薬効分類611-629）
# 介入点: 2024年4月（令和6年度診療報酬改定）
# データ: 第10回NDB (FY2023, t=1-12) + 第11回NDB (FY2024, t=13-24)

suppressPackageStartupMessages({
  library(readxl)
  library(dplyr)
  library(tidyr)
  library(ggplot2)
  library(nlme)
  library(writexl)
})

# ── 設定 ──────────────────────────────────────────────────────────────────────
ANTIBIOTIC_CODES      <- c(611L, 612L, 613L, 614L, 615L, 616L, 617L, 619L, 624L, 629L)
EXCLUDE_629_DRUG_CODE <- 622136101L
T0                    <- 12L  # t=12 が介入直前最終月（2024年3月）

DRUG_CLASS_NAMES <- c(
  "611" = "グラム陽性菌",
  "612" = "グラム陰性菌",
  "613" = "グラム陽性・陰性菌",
  "614" = "グラム陽性菌・マイコプラズマ",
  "615" = "グラム陽性・陰性菌・クラミジア等",
  "616" = "抗酸菌",
  "617" = "抗真菌薬",
  "619" = "その他抗生物質製剤",
  "624" = "合成抗菌剤",
  "629" = "その他化学療法剤"
)

CATEGORIES <- list(
  "外来院内" = "注射薬 外来 (院内)",
  "外来院外" = "注射薬 外来 (院外)",
  "入院"             = "注射薬 入院"
)

RD10_PATH  <- "/project/data/no10_raw/injection_monthly.xlsx"
RD11_PATH  <- "/project/data/no11_raw/【注射】診療月別薬効分類別数量(公費含む).xlsx"
OUTPUT_DIR <- "/project/output"

# 月列名 → 年度内時刻オフセット（4月=1, ..., 3月=12）
MONTH_COLS   <- c("m4","m5","m6","m7","m8","m9","m10","m11","m12","m1","m2","m3")
MONTH_OFFSET <- setNames(1:12, MONTH_COLS)

# t=1..24 に対応するカレンダー月
DATES_24 <- as.Date(c(
  "2023-04-01","2023-05-01","2023-06-01","2023-07-01","2023-08-01","2023-09-01",
  "2023-10-01","2023-11-01","2023-12-01","2024-01-01","2024-02-01","2024-03-01",
  "2024-04-01","2024-05-01","2024-06-01","2024-07-01","2024-08-01","2024-09-01",
  "2024-10-01","2024-11-01","2024-12-01","2025-01-01","2025-02-01","2025-03-01"
))

dir.create(file.path(OUTPUT_DIR, "its_plots"), showWarnings = FALSE, recursive = TRUE)

# ── Excelシートの読み込み ──────────────────────────────────────────────────────
read_injection_sheet <- function(path, sheet_name) {
  col_names <- c(
    "class_code", "class_name", "drug_code", "drug_name",
    "unit", "price_code", "price", "generic_flag", "total",
    "m4","m5","m6","m7","m8","m9","m10","m11","m12","m1","m2","m3"
  )

  # 先頭4行（タイトル・空・ヘッダー・空）をスキップ
  df <- tryCatch(
    read_excel(path, sheet = sheet_name, col_names = col_names, skip = 4),
    error = function(e) {
      message("  ERROR reading ", path, " sheet=", sheet_name, ": ", e$message)
      return(NULL)
    }
  )
  if (is.null(df)) return(NULL)

  # 列数が不足する場合は補完
  if (ncol(df) < 21) {
    for (col in col_names[!col_names %in% names(df)]) df[[col]] <- NA_real_
  }

  # 薬品コードが NA の行（ヘッダー残渣・区切り行）を除去
  df <- df %>% filter(!is.na(drug_code))

  # 薬効分類コードを前方補完
  df$class_code <- suppressWarnings(as.integer(df$class_code))
  df$class_name <- as.character(df$class_name)
  df <- df %>% fill(class_code, class_name, .direction = "down")

  df$drug_code <- suppressWarnings(as.integer(df$drug_code))

  # 月別数量: 抑制値（'‐'等）および非数値 → 0
  df <- df %>%
    mutate(across(all_of(MONTH_COLS), ~ suppressWarnings(as.numeric(.x)))) %>%
    mutate(across(all_of(MONTH_COLS), ~ replace_na(.x, 0)))

  df
}

# ── 抗菌薬フィルタリング & 薬効分類別月次集計 ────────────────────────────────
aggregate_monthly <- function(df, t_offset) {
  df %>%
    filter(class_code %in% ANTIBIOTIC_CODES) %>%
    filter(!(class_code == 629L & drug_code == EXCLUDE_629_DRUG_CODE)) %>%
    group_by(class_code, class_name) %>%
    summarise(across(all_of(MONTH_COLS), sum, na.rm = TRUE), .groups = "drop") %>%
    pivot_longer(cols = all_of(MONTH_COLS), names_to = "month_col", values_to = "Y") %>%
    mutate(t = MONTH_OFFSET[month_col] + t_offset) %>%
    select(class_code, class_name, t, Y)
}

# ── セグメント回帰モデル ───────────────────────────────────────────────────────
# Y_t = beta0 + beta1*t + beta2*D + beta3*(t-T0)*D + epsilon
# beta2: 介入直後の水準変化（level change）
# beta3: 介入後の傾き変化（slope change）
fit_its <- function(ts_df, t0 = T0) {
  ts_df <- ts_df %>%
    mutate(
      D      = as.integer(t > t0),
      post_t = pmax(t - t0, 0L) * D
    )

  m_ols <- lm(Y ~ t + D + post_t, data = ts_df)

  # GLS + AR(1) 誤差（系列相関の頑健性チェック）
  m_gls <- tryCatch(
    gls(Y ~ t + D + post_t, data = ts_df,
        correlation = corAR1(form = ~t), method = "ML"),
    error = function(e) NULL
  )

  # Durbin-Watson 統計量（手動計算）
  res <- residuals(m_ols)
  dw  <- sum(diff(res)^2) / sum(res^2)

  list(ols = m_ols, gls = m_gls, data = ts_df, dw = dw)
}

# ── 係数テーブル抽出 ──────────────────────────────────────────────────────────
coef_table <- function(fit, model_type = "OLS") {
  if (model_type == "OLS") {
    s  <- summary(fit$ols)
    ct <- coef(s)
    ci <- confint(fit$ols)
    data.frame(
      model     = "OLS",
      term      = rownames(ct),
      estimate  = ct[, "Estimate"],
      ci_lower  = ci[, 1],
      ci_upper  = ci[, 2],
      std_error = ct[, "Std. Error"],
      t_value   = ct[, "t value"],
      p_value   = ct[, "Pr(>|t|)"],
      dw_stat   = fit$dw,
      row.names = NULL
    )
  } else {
    if (is.null(fit$gls)) return(NULL)
    s  <- summary(fit$gls)
    ct <- s$tTable
    ci <- tryCatch(
      intervals(fit$gls, which = "coef")$coef,
      error = function(e) NULL
    )
    if (is.null(ci)) return(NULL)
    data.frame(
      model     = "GLS(AR1)",
      term      = rownames(ct),
      estimate  = ct[, "Value"],
      ci_lower  = ci[, "lower"],
      ci_upper  = ci[, "upper"],
      std_error = ct[, "Std.Error"],
      t_value   = ct[, "t-value"],
      p_value   = ct[, "p-value"],
      dw_stat   = NA_real_,
      row.names = NULL
    )
  }
}

# ── ITS プロット ──────────────────────────────────────────────────────────────
plot_its <- function(fit, title_str) {
  df <- fit$data %>%
    mutate(
      date           = DATES_24[t],
      fitted         = predict(fit$ols),
      counterfactual = coef(fit$ols)["(Intercept)"] + coef(fit$ols)["t"] * t
    )

  y_max <- max(df$Y, na.rm = TRUE)

  ggplot(df, aes(x = date)) +
    geom_point(aes(y = Y), color = "steelblue", size = 2.2, alpha = 0.85) +
    geom_line(aes(y = fitted),         color = "steelblue", linewidth = 1.0) +
    geom_line(aes(y = counterfactual), color = "grey55",    linewidth = 0.8, linetype = "dashed") +
    geom_vline(xintercept = as.Date("2024-04-01"), color = "firebrick",
               linetype = "dotted", linewidth = 0.9) +
    annotate("text", x = as.Date("2024-04-15"), y = y_max * 0.97,
             label = "介入 2024-04", color = "firebrick", size = 3, hjust = 0) +
    scale_x_date(date_breaks = "3 months", date_labels = "%Y-%m") +
    labs(
      title    = title_str,
      subtitle = "実線=フィット値  破線=反事実（介入なし想定）",
      x        = "診療月",
      y        = "処方数量合計"
    ) +
    theme_bw(base_size = 11) +
    theme(
      axis.text.x   = element_text(angle = 45, hjust = 1),
      plot.title    = element_text(size = 11, face = "bold"),
      plot.subtitle = element_text(size = 9, color = "grey40")
    )
}

# ── メイン ────────────────────────────────────────────────────────────────────
all_results <- list()

for (cat_label in names(CATEGORIES)) {
  sheet_name <- CATEGORIES[[cat_label]]
  cat(sprintf("\n=== カテゴリ: %s ===\n", cat_label))

  rd10 <- read_injection_sheet(RD10_PATH, sheet_name)
  rd11 <- read_injection_sheet(RD11_PATH, sheet_name)

  if (is.null(rd10) || is.null(rd11)) {
    cat("  -> データ読み込み失敗。スキップ\n")
    next
  }

  ts10 <- aggregate_monthly(rd10, t_offset = 0L)
  ts11 <- aggregate_monthly(rd11, t_offset = 12L)
  ts   <- bind_rows(ts10, ts11) %>% arrange(class_code, t)

  cat(sprintf("  rd10: %d行, rd11: %d行, 薬効分類: %s\n",
              nrow(rd10), nrow(rd11),
              paste(sort(unique(ts$class_code)), collapse = ",")))

  for (code in sort(unique(ts$class_code))) {
    cn   <- DRUG_CLASS_NAMES[as.character(code)]
    ts_c <- ts %>% filter(class_code == code)

    y_max <- max(ts_c$Y, na.rm = TRUE)
    cat(sprintf("  Class %d (%s): n=%d, Y_max=%.0f\n", code, cn, nrow(ts_c), y_max))

    if (y_max == 0 || nrow(ts_c) < 10) {
      cat("    -> 数量ゼロまたはデータ不足。スキップ\n")
      next
    }

    fit <- fit_its(ts_c)

    ct_ols <- coef_table(fit, "OLS") %>%
      mutate(category = cat_label, class_code = code, class_name = cn)
    ct_gls <- coef_table(fit, "GLS")
    if (!is.null(ct_gls)) {
      ct_gls <- ct_gls %>%
        mutate(category = cat_label, class_code = code, class_name = cn)
    }

    all_results <- c(all_results, list(ct_ols), list(ct_gls))

    cat(sprintf("    DW=%.3f  beta2(level)=%+.1f (p=%.3f)  beta3(slope)=%+.1f (p=%.3f)\n",
                fit$dw,
                coef(fit$ols)["D"],      coef(summary(fit$ols))["D",      "Pr(>|t|)"],
                coef(fit$ols)["post_t"], coef(summary(fit$ols))["post_t", "Pr(>|t|)"]))

    p_title <- sprintf("%s / %d %s", cat_label, code, cn)
    p <- plot_its(fit, p_title)
    ggsave(
      filename = file.path(OUTPUT_DIR, "its_plots",
                           sprintf("its_%s_%d.png", cat_label, code)),
      plot  = p, width = 10, height = 5, dpi = 150
    )
  }
}

# ── 結果保存 ──────────────────────────────────────────────────────────────────
results_df <- bind_rows(all_results[!sapply(all_results, is.null)]) %>%
  select(category, class_code, class_name, model, term,
         estimate, ci_lower, ci_upper, std_error, t_value, p_value, dw_stat) %>%
  arrange(category, class_code, model, term)

guide_df <- data.frame(
  term   = c("(Intercept)", "t", "D", "post_t"),
  meaning = c(
    "t=0時点の切片（ベースライン推定値）",
    "介入前の月次トレンド（傍き）",
    "介入直後の水準変化（Level change）",
    "介入後の傍き変化（Slope change）"
  ),
  note = c(
    "基準となる処方量の推定値",
    "正→介入前に増加傾向、負→減少傾向",
    "正→介入後に処方量が急増、負→急減",
    "正→介入後に増加傾向が強まった、負→弱まった"
  )
)

write_xlsx(
  list("イッツ結果" = results_df, "変数の解釈" = guide_df),
  file.path(OUTPUT_DIR, "its_antibiotics_results.xlsx")
)

n_plots <- length(list.files(file.path(OUTPUT_DIR, "its_plots"), pattern = "\\.png$"))
cat(sprintf("\n[Done] its_antibiotics_results.xlsx + its_plots/ (%d PNG)\n", n_plots))
