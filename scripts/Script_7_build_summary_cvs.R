# ============================================================
# Script 06 (CLEAN / GitHub-ready):
# 1) Build MASTER_ALL from per-tag time_enrichment_MASTER_summary.csv
# 2) Add discipline (math vs lit) via tag prefix (m## vs l##)
# 3) Group summary (mean + t-based 95% CI)
# 4) math-lit diff (bootstrap CI + permutation p)
#
# Expected repo layout:
#   results/
#     m01/time_enrichment_MASTER_summary.csv
#     ...
#     l10/time_enrichment_MASTER_summary.csv
#
# Usage:
#   Rscript scripts/06_master_all_enrichment_math_vs_lit.R
#   Rscript scripts/06_master_all_enrichment_math_vs_lit.R 5000 5000 123
#
# Env (optional):
#   GESTURE_PROJECT_ROOT=/path/to/repo_root
# 
# Note:
#   FPS is read from video metadata; user-specified fps is ignored to prevent mismatch.
# ============================================================

suppressPackageStartupMessages({
  library(readr)
  library(dplyr)
  library(tibble)
  library(purrr)
  library(tidyr)
  library(stringr)
})

# ----------------------------
# (0) SETTINGS / CLI
# ----------------------------
args <- commandArgs(trailingOnly = TRUE)

B    <- if (length(args) >= 1) as.integer(args[[1]]) else 5000L  # bootstrap reps
P    <- if (length(args) >= 2) as.integer(args[[2]]) else 5000L  # perm reps
seed <- if (length(args) >= 3) as.integer(args[[3]]) else 123L

stopifnot(B >= 200, P >= 200)
set.seed(seed)

# project root (portable)
project_root <- Sys.getenv("GESTURE_PROJECT_ROOT", unset = getwd())
results_root <- file.path(project_root, "results")
dir.create(results_root, showWarnings = FALSE, recursive = TRUE)

# outputs
master_all_csv <- file.path(results_root, "time_enrichment_MASTER_ALL.csv")
out_with_disc  <- file.path(results_root, "time_enrichment_MASTER_ALL_with_discipline.csv")
out_group_sum  <- file.path(results_root, "time_enrichment_group_summary_math_vs_lit.csv")
out_diff       <- file.path(results_root, "time_enrichment_diff_math_minus_lit_bootstrap_perm.csv")

cat("[INFO] project_root:", normalizePath(project_root, winslash = "/", mustWork = FALSE), "\n")
cat("[INFO] results_root:", normalizePath(results_root, winslash = "/", mustWork = FALSE), "\n")
cat("[INFO] B=", B, " P=", P, " seed=", seed, "\n\n")

# ----------------------------
# (1) AUTO-DISCOVER TAGS from results/<tag>/time_enrichment_MASTER_summary.csv
# ----------------------------
discover_tag_summaries <- function(results_root) {
  # find all files matching: results/<tag>/time_enrichment_MASTER_summary.csv
  files <- list.files(results_root,
                      pattern = "^time_enrichment_MASTER_summary\\.csv$",
                      recursive = TRUE,
                      full.names = TRUE)
  if (length(files) == 0) return(tibble(tag = character(0), csv_path = character(0)))
  
  # tag = immediate parent folder name
  folder <- basename(dirname(files))
  tibble(tag = folder, csv_path = files) %>%
    distinct(tag, .keep_all = TRUE) %>%
    arrange(tag)
}

per_tag_paths <- discover_tag_summaries(results_root)

if (nrow(per_tag_paths) == 0) {
  stop("No per-tag summaries found.\nExpected: results/<TAG>/time_enrichment_MASTER_summary.csv", call. = FALSE)
}

cat("[OK] Found per-tag summary files:", nrow(per_tag_paths), "\n")
cat("     Tags:", paste(per_tag_paths$tag, collapse = ", "), "\n\n")

# ----------------------------
# (2) BUILD MASTER_ALL
# ----------------------------
df_master <- purrr::map_dfr(seq_len(nrow(per_tag_paths)), function(i) {
  tag <- per_tag_paths$tag[[i]]
  p   <- per_tag_paths$csv_path[[i]]
  
  readr::read_csv(p, show_col_types = FALSE) %>%
    mutate(tag = tag, .before = 1)
})

if (nrow(df_master) == 0) stop("Read 0 rows from per-tag summaries.", call. = FALSE)

names(df_master) <- tolower(names(df_master))
readr::write_csv(df_master, master_all_csv)
cat("[OK] Wrote MASTER_ALL:\n  ", master_all_csv, "\n")
cat("[OK] Rows:", nrow(df_master), " | Videos:", n_distinct(df_master$tag), "\n\n")

# ----------------------------
# (3) ADD DISCIPLINE from tag prefix
#   m## -> math
#   l## -> lit
# ----------------------------
df_master2 <- df_master %>%
  mutate(
    discipline = case_when(
      str_detect(tag, "^m\\d+") ~ "math",
      str_detect(tag, "^l\\d+") ~ "lit",
      TRUE ~ NA_character_
    )
  )

bad <- df_master2 %>% filter(is.na(discipline)) %>% distinct(tag)
if (nrow(bad) > 0) {
  stop("Unrecognized tag(s) (cannot infer discipline): ", paste(bad$tag, collapse = ", "), call. = FALSE)
}

readr::write_csv(df_master2, out_with_disc)
cat("[OK] Wrote:\n  ", out_with_disc, "\n\n")

# ----------------------------
# (4) GROUP SUMMARY (math vs lit) + t-based 95% CI
# ----------------------------
stopifnot("lift" %in% names(df_master2))
has_delta <- "delta_rate" %in% names(df_master2)
has_p     <- "p_value_plus1" %in% names(df_master2)

group_cols <- c("discipline", "pointset", "null_mode", "window_s")
group_cols <- intersect(group_cols, names(df_master2))

safe_mean <- function(x) if (all(is.na(x))) NA_real_ else mean(x, na.rm = TRUE)
safe_sd   <- function(x) if (sum(is.finite(x)) < 2) NA_real_ else sd(x, na.rm = TRUE)
safe_med  <- function(x) if (all(is.na(x))) NA_real_ else median(x, na.rm = TRUE)

df_group <- df_master2 %>%
  group_by(across(all_of(group_cols))) %>%
  summarise(
    # lift
    lift_n    = sum(is.finite(lift)),
    lift_mean = safe_mean(lift),
    lift_sd   = safe_sd(lift),
    lift_se   = ifelse(lift_n >= 2, lift_sd / sqrt(lift_n), NA_real_),
    lift_t    = ifelse(lift_n >= 2, qt(0.975, df = lift_n - 1), NA_real_),
    lift_ci95_lo = ifelse(lift_n >= 2, lift_mean - lift_t * lift_se, NA_real_),
    lift_ci95_hi = ifelse(lift_n >= 2, lift_mean + lift_t * lift_se, NA_real_),
    lift_median  = safe_med(lift),
    
    # delta_rate optional
    delta_n    = if (has_delta) sum(is.finite(delta_rate)) else 0L,
    delta_mean = if (has_delta) safe_mean(delta_rate) else NA_real_,
    delta_sd   = if (has_delta) safe_sd(delta_rate) else NA_real_,
    delta_se   = ifelse(delta_n >= 2, delta_sd / sqrt(delta_n), NA_real_),
    delta_t    = ifelse(delta_n >= 2, qt(0.975, df = delta_n - 1), NA_real_),
    delta_ci95_lo = ifelse(delta_n >= 2, delta_mean - delta_t * delta_se, NA_real_),
    delta_ci95_hi = ifelse(delta_n >= 2, delta_mean + delta_t * delta_se, NA_real_),
    delta_median  = if (has_delta) safe_med(delta_rate) else NA_real_,
    
    sig_rate_p005 = if (has_p) mean(p_value_plus1 < 0.05, na.rm = TRUE) else NA_real_,
    .groups = "drop"
  ) %>%
  arrange(across(all_of(group_cols)))

readr::write_csv(df_group, out_group_sum)
cat("[OK] Wrote:\n  ", out_group_sum, "\n\n")

# ----------------------------
# (5) DIFF TABLE: math - lit (bootstrap CI + permutation p)
# ----------------------------
strata_cols <- c("pointset", "null_mode", "window_s")
strata_cols <- intersect(strata_cols, names(df_master2))

boot_diff_ci <- function(x_math, x_lit, B = 5000L) {
  x_math <- x_math[is.finite(x_math)]
  x_lit  <- x_lit[is.finite(x_lit)]
  n_m <- length(x_math)
  n_l <- length(x_lit)
  
  mean_m <- if (n_m > 0) mean(x_math) else NA_real_
  mean_l <- if (n_l > 0) mean(x_lit)  else NA_real_
  obs <- mean_m - mean_l
  
  if (n_m < 2 || n_l < 2) {
    return(tibble(
      n_math = n_m, n_lit = n_l,
      mean_math = mean_m, mean_lit = mean_l,
      diff_mean = obs,
      diff_ci95_lo = NA_real_, diff_ci95_hi = NA_real_,
      boot_sd = NA_real_
    ))
  }
  
  diffs <- replicate(B, {
    mean(sample(x_math, size = n_m, replace = TRUE)) -
      mean(sample(x_lit,  size = n_l, replace = TRUE))
  })
  
  tibble(
    n_math = n_m, n_lit = n_l,
    mean_math = mean_m, mean_lit = mean_l,
    diff_mean = obs,
    diff_ci95_lo = unname(quantile(diffs, 0.025, na.rm = TRUE, type = 7)),
    diff_ci95_hi = unname(quantile(diffs, 0.975, na.rm = TRUE, type = 7)),
    boot_sd = sd(diffs, na.rm = TRUE)
  )
}

perm_p_value <- function(x_math, x_lit, P = 5000L) {
  x_math <- x_math[is.finite(x_math)]
  x_lit  <- x_lit[is.finite(x_lit)]
  n_m <- length(x_math)
  n_l <- length(x_lit)
  
  if (n_m < 2 || n_l < 2) {
    return(tibble(
      perm_p = NA_real_,
      perm_P = P,
      perm_method = "label_permutation_two_sided"
    ))
  }
  
  obs <- mean(x_math) - mean(x_lit)
  allx <- c(x_math, x_lit)
  n <- length(allx)
  
  perm_stats <- replicate(P, {
    idx <- sample.int(n, size = n_m, replace = FALSE)
    mean(allx[idx]) - mean(allx[-idx])
  })
  
  p <- (sum(abs(perm_stats) >= abs(obs)) + 1) / (P + 1)
  
  tibble(
    perm_p = p,
    perm_P = P,
    perm_method = "label_permutation_two_sided"
  )
}

make_diff_table <- function(df, metric_col, B = 5000L, P = 5000L) {
  stopifnot(metric_col %in% names(df))
  
  df %>%
    group_by(across(all_of(strata_cols))) %>%
    summarise(
      math_vals = list(.data[[metric_col]][discipline == "math"]),
      lit_vals  = list(.data[[metric_col]][discipline == "lit"]),
      .groups = "drop"
    ) %>%
    rowwise() %>%
    mutate(
      boot = list(boot_diff_ci(unlist(math_vals), unlist(lit_vals), B = B)),
      perm = list(perm_p_value(unlist(math_vals), unlist(lit_vals), P = P))
    ) %>%
    ungroup() %>%
    select(-math_vals, -lit_vals) %>%
    unnest(boot) %>%
    unnest(perm) %>%
    mutate(
      metric = metric_col,
      sig_perm_p005 = ifelse(is.na(perm_p), NA, perm_p < 0.05)
    ) %>%
    relocate(metric, .before = 1)
}

df_diff_lift <- make_diff_table(df_master2, "lift", B = B, P = P)
df_diff_delta <- if (has_delta) make_diff_table(df_master2, "delta_rate", B = B, P = P) else tibble()

df_diff <- bind_rows(df_diff_lift, df_diff_delta) %>%
  arrange(metric, across(all_of(strata_cols)))

readr::write_csv(df_diff, out_diff)
cat("[OK] Wrote:\n  ", out_diff, "\n\n")

cat("DONE.\n",
    "MASTER_ALL      : ", master_all_csv, "\n",
    "MASTER_ALL+disc : ", out_with_disc,  "\n",
    "Group summary   : ", out_group_sum,  "\n",
    "Diff (boot+perm): ", out_diff,       "\n",
    sep = "")
