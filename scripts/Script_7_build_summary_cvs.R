# ============================================================
# Script 07 (GitHub-ready / AUTO-DISCOVER TAGS)
# 1) Auto-discover per-tag summaries under results/<TAG>/time_enrichment_MASTER_summary.csv
# 2) Build MASTER_ALL
# 3) Add discipline (math vs lit) based on tag prefix (m/l)
# 4) Group summary (mean + t-based 95% CI + bootstrap 95% CI)
# 5) math-lit diff (bootstrap CI + permutation p)
#
# Repo layout expected:
#   results/<TAG>/time_enrichment_MASTER_summary.csv
#
# CLI (optional):
#   Rscript scripts/07_master_aggregate.R
#   Rscript scripts/07_master_aggregate.R 5000 5000 5000 123
#     args: B_GROUP, B_DIFF, P_DIFF, seed
#
# Env (optional):
#   GESTURE_PROJECT_ROOT=/path/to/repo_root
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

B_GROUP <- if (length(args) >= 1) as.integer(args[[1]]) else 5000L
B_DIFF  <- if (length(args) >= 2) as.integer(args[[2]]) else 5000L
P_DIFF  <- if (length(args) >= 3) as.integer(args[[3]]) else 5000L
seed    <- if (length(args) >= 4) as.integer(args[[4]]) else 123L

stopifnot(B_GROUP >= 200, B_DIFF >= 200, P_DIFF >= 200)
set.seed(seed)

# switches
REBUILD_MASTER_ALL  <- TRUE
STRICT_TAGS         <- TRUE
KEEP_ONLY_STATUS_OK <- TRUE

# Optional: restrict tags (NULL = all)
TAG_WHITELIST <- NULL
# TAG_WHITELIST <- c("m04_2","m04_3","l08")

# Portable project root
project_root <- Sys.getenv("GESTURE_PROJECT_ROOT", unset = getwd())
results_root <- file.path(project_root, "results")
if (!dir.exists(results_root)) {
  stop("results/ folder not found. Run from repo root or set GESTURE_PROJECT_ROOT.", call. = FALSE)
}

# outputs
master_all_csv <- file.path(results_root, "time_enrichment_MASTER_ALL.csv")
out_with_disc  <- file.path(results_root, "time_enrichment_MASTER_ALL_with_discipline.csv")
out_group_sum  <- file.path(results_root, "time_enrichment_group_summary_math_vs_lit.csv")
out_diff       <- file.path(results_root, "time_enrichment_diff_math_minus_lit_bootstrap_perm.csv")

cat("[INFO] project_root :", normalizePath(project_root, winslash="/", mustWork = FALSE), "\n")
cat("[INFO] results_root :", normalizePath(results_root, winslash="/", mustWork = FALSE), "\n\n")

# ----------------------------
# (1) DISCOVER per-tag summary files
# ----------------------------
paths <- list.files(
  path = results_root,
  pattern = "^time_enrichment_MASTER_summary\\.csv$",
  recursive = TRUE,
  full.names = TRUE
)

paths <- paths[file.exists(paths)]
if (length(paths) == 0) {
  stop(
    "No per-tag summary files found under results/.\n",
    "Expected: results/<TAG>/time_enrichment_MASTER_summary.csv",
    call. = FALSE
  )
}

tags_found <- basename(dirname(paths))

# normalize tags to underscore canonical (avoid m02-2 vs m02_2 mismatch)
tags_found <- str_replace_all(tags_found, "-", "_")

df_paths <- tibble(tag = tags_found, csv_path = paths) %>%
  distinct(tag, .keep_all = TRUE)

if (!is.null(TAG_WHITELIST)) {
  TAG_WHITELIST <- str_replace_all(TAG_WHITELIST, "-", "_")
  df_paths <- df_paths %>% filter(tag %in% TAG_WHITELIST)
  if (nrow(df_paths) == 0) stop("TAG_WHITELIST filtered everything out.", call. = FALSE)
}

cat("[OK] Discovered tags:", nrow(df_paths), "\n")
cat("     Examples:", paste(head(df_paths$tag, 8), collapse = ", "), "\n\n")

# ----------------------------
# (2) BUILD MASTER_ALL
# ----------------------------
need_build <- REBUILD_MASTER_ALL || !file.exists(master_all_csv)

read_one_safe <- function(tag, csv_path) {
  tryCatch({
    x <- readr::read_csv(csv_path, show_col_types = FALSE, progress = FALSE)
    x$tag <- tag
    x
  }, error = function(e) {
    if (STRICT_TAGS) stop("Failed reading: ", csv_path, "\n", conditionMessage(e), call. = FALSE)
    message("[WARN] Failed reading: ", csv_path, " | ", conditionMessage(e))
    NULL
  })
}

if (need_build) {
  cat("[BUILD] Building MASTER_ALL...\n")
  df_master <- purrr::pmap_dfr(df_paths, read_one_safe)
  
  if (nrow(df_master) == 0) stop("MASTER_ALL is empty (all reads failed).", call. = FALSE)
  
  names(df_master) <- tolower(names(df_master))
  
  # normalize legacy variants
  if (!("pointset" %in% names(df_master)) && "point_set" %in% names(df_master)) {
    df_master <- df_master %>% rename(pointset = point_set)
  }
  if (!("null_mode" %in% names(df_master)) && "null" %in% names(df_master)) {
    df_master <- df_master %>% rename(null_mode = null)
  }
  
  # standardize pointset naming
  if ("pointset" %in% names(df_master)) {
    df_master <- df_master %>%
      mutate(pointset = as.character(pointset)) %>%
      mutate(pointset = case_when(
        pointset %in% c("all", "allpoints", "all_points") ~ "all_points",
        pointset %in% c("conf1", "conf1_only", "c1_only") ~ "conf1_only",
        pointset %in% c("conf2", "conf2_only", "c2_only") ~ "conf2_only",
        pointset %in% c("conf3", "conf3_only", "c3_only") ~ "conf3_only",
        TRUE ~ pointset
      ))
  }
  
  # optionally keep only ok rows
  if (KEEP_ONLY_STATUS_OK && "status" %in% names(df_master)) {
    df_master <- df_master %>% filter(status == "ok")
  }
  
  # ensure tag is character + canonical
  df_master <- df_master %>%
    mutate(tag = str_replace_all(as.character(tag), "-", "_"))
  
  writer::write_csv(df_master, master_all_csv)
  cat("[OK] Wrote MASTER_ALL:\n  ", master_all_csv, "\n", sep = "")
  cat("[OK] Rows:", nrow(df_master), " | Videos:", dplyr::n_distinct(df_master$tag), "\n\n")
  
} else {
  cat("[OK] MASTER_ALL exists and REBUILD_MASTER_ALL=FALSE. Using existing file.\n\n")
}

# ----------------------------
# (3) READ MASTER_ALL + ADD DISCIPLINE
# ----------------------------
df_master <- readr::read_csv(master_all_csv, show_col_types = FALSE, progress = FALSE)
names(df_master) <- tolower(names(df_master))

stopifnot("tag" %in% names(df_master))
stopifnot("lift" %in% names(df_master))

df_master2 <- df_master %>%
  mutate(
    tag = str_replace_all(as.character(tag), "-", "_"),
    discipline = case_when(
      str_detect(tag, "^m") ~ "math",
      str_detect(tag, "^l") ~ "lit",
      TRUE ~ NA_character_
    )
  )

bad <- df_master2 %>% filter(is.na(discipline)) %>% distinct(tag)
if (nrow(bad) > 0) {
  stop("Some tags do not start with m/l, cannot map discipline:\n",
       paste(bad$tag, collapse = ", "), call. = FALSE)
}

writer::write_csv(df_master2, out_with_disc)
cat("[OK] Wrote:\n  ", out_with_disc, "\n\n", sep = "")

# ----------------------------
# (4) GROUP SUMMARY: mean + t CI + bootstrap CI
# ----------------------------
safe_mean <- function(x) if (all(!is.finite(x))) NA_real_ else mean(x, na.rm = TRUE)
safe_sd   <- function(x) if (sum(is.finite(x)) < 2) NA_real_ else sd(x, na.rm = TRUE)
safe_med  <- function(x) if (all(!is.finite(x))) NA_real_ else median(x, na.rm = TRUE)

boot_mean_ci <- function(x, B = 5000L) {
  x <- x[is.finite(x)]
  n <- length(x)
  if (n < 2) return(tibble(boot_lo = NA_real_, boot_hi = NA_real_, boot_sd = NA_real_))
  boots <- replicate(B, mean(sample(x, size = n, replace = TRUE)))
  tibble(
    boot_lo = unname(quantile(boots, 0.025, na.rm = TRUE, type = 7)),
    boot_hi = unname(quantile(boots, 0.975, na.rm = TRUE, type = 7)),
    boot_sd = sd(boots, na.rm = TRUE)
  )
}

group_cols <- intersect(c("discipline", "pointset", "null_mode", "window_s"), names(df_master2))
has_p <- "p_value_plus1" %in% names(df_master2)

df_group <- df_master2 %>%
  group_by(across(all_of(group_cols))) %>%
  summarise(
    lift_vals  = list(lift),
    lift_n     = sum(is.finite(lift)),
    lift_mean  = safe_mean(lift),
    lift_mean_centered = lift_mean - 1,   # <-- NEW
    lift_sd    = safe_sd(lift),
    lift_se    = ifelse(lift_n >= 2, lift_sd / sqrt(lift_n), NA_real_),
    lift_tcrit = ifelse(lift_n >= 2, qt(0.975, df = lift_n - 1), NA_real_),
    lift_t_lo  = ifelse(lift_n >= 2, lift_mean - lift_tcrit * lift_se, NA_real_),
    lift_t_hi  = ifelse(lift_n >= 2, lift_mean + lift_tcrit * lift_se, NA_real_),
    lift_median = safe_med(lift),
    sig_rate_p005 = if (has_p) mean(p_value_plus1 < 0.05, na.rm = TRUE) else NA_real_,
    .groups = "drop"
  ) %>%
  rowwise() %>%
  mutate(boot = list(boot_mean_ci(unlist(lift_vals), B = B_GROUP))) %>%
  ungroup() %>%
  select(-lift_vals) %>%
  unnest(boot)


if ("pointset" %in% names(df_group)) {
  df_group <- df_group %>%
    mutate(pointset = factor(pointset, levels = c("conf3_only", "conf2_only", "conf1_only", "all_points")))
}
if ("null_mode" %in% names(df_group)) {
  df_group <- df_group %>%
    mutate(null_mode = factor(null_mode, levels = c("teacher_only", "global")))
}

writer::write_csv(df_group, out_group_sum)
cat("[OK] Wrote:\n  ", out_group_sum, "\n\n", sep = "")

# ----------------------------
# (5) DIFF: math - lit (bootstrap CI + permutation p)
# ----------------------------
strata_cols <- intersect(c("pointset", "null_mode", "window_s"), names(df_master2))

boot_diff_ci <- function(x_math, x_lit, B = 5000L) {
  x_math <- x_math[is.finite(x_math)]
  x_lit  <- x_lit[is.finite(x_lit)]
  n_m <- length(x_math); n_l <- length(x_lit)
  
  mean_m <- if (n_m > 0) mean(x_math) else NA_real_
  mean_l <- if (n_l > 0) mean(x_lit)  else NA_real_
  obs <- mean_m - mean_l
  
  if (n_m < 2 || n_l < 2) {
    return(tibble(
      n_math = n_m, n_lit = n_l,
      mean_math = mean_m, mean_lit = mean_l,
      diff_mean = obs,
      diff_ci95_lo = NA_real_, diff_ci95_hi = NA_real_
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
    diff_ci95_hi = unname(quantile(diffs, 0.975, na.rm = TRUE, type = 7))
  )
}

perm_p_value <- function(x_math, x_lit, P = 5000L) {
  x_math <- x_math[is.finite(x_math)]
  x_lit  <- x_lit[is.finite(x_lit)]
  n_m <- length(x_math); n_l <- length(x_lit)
  
  if (n_m < 2 || n_l < 2) return(tibble(perm_p = NA_real_, perm_P = P))
  
  obs <- mean(x_math) - mean(x_lit)
  allx <- c(x_math, x_lit)
  n <- length(allx)
  
  perm_stats <- replicate(P, {
    idx <- sample.int(n, size = n_m, replace = FALSE)
    mean(allx[idx]) - mean(allx[-idx])
  })
  
  p <- (sum(abs(perm_stats) >= abs(obs)) + 1) / (P + 1)
  tibble(perm_p = p, perm_P = P)
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

df_diff <- make_diff_table(df_master2, "lift", B = B_DIFF, P = P_DIFF)
writer::write_csv(df_diff, out_diff)
cat("[OK] Wrote:\n  ", out_diff, "\n\n", sep = "")

cat("DONE.\n",
    "MASTER_ALL      : ", master_all_csv, "\n",
    "MASTER_ALL+disc : ", out_with_disc,  "\n",
    "Group summary   : ", out_group_sum,  "\n",
    "Diff (boot+perm): ", out_diff,       "\n",
    sep = "")
