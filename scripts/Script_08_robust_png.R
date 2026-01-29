# ============================================================
# Script 08 (ROBUST PLOTTER):
#   - Read group summary + diff tables produced by Script 07
#   - Auto-detect CI columns (bootstrap vs t-based vs legacy)
#   - Robust to duplicated column names & make.names()
#   - NEW: optional centered metric plotting
#
# Expected repo layout (run from repo root):
#   results/
#     time_enrichment_group_summary_math_vs_lit.csv
#     time_enrichment_diff_math_minus_lit_bootstrap_perm.csv
#
# Output:
#   results/fig_group_mean_lift_math_vs_lit.png
#   results/fig_diff_lift_math_minus_lit.png
#
# Usage:
#   Rscript scripts/08_plotter.R
#   Rscript scripts/08_plotter.R auto
#   Rscript scripts/08_plotter.R boot
#   Rscript scripts/08_plotter.R t
#
# Optional env var:
#   GESTURE_PROJECT_ROOT=/path/to/repo_root
# ============================================================

suppressPackageStartupMessages({
  library(readr)
  library(dplyr)
  library(ggplot2)
  library(stringr)
  library(tidyr)
})

# ----------------------------
# (0) SETTINGS / CLI
# ----------------------------
args <- commandArgs(trailingOnly = TRUE)

CI_MODE <- if (length(args) >= 1) tolower(args[[1]]) else "auto"
if (!(CI_MODE %in% c("auto","boot","t"))) {
  stop("CI_MODE must be one of: auto | boot | t", call. = FALSE)
}

SAVE_PNG <- TRUE
DPI <- 320

# NEW: centered plotting switch
#   FALSE -> y = lift_mean, baseline at 1
#   TRUE  -> y = lift_mean_centered (if exists), baseline at 0
USE_CENTERED <- TRUE

project_root <- Sys.getenv("GESTURE_PROJECT_ROOT", unset = getwd())
results_root <- file.path(project_root, "results")
dir.create(results_root, recursive = TRUE, showWarnings = FALSE)

in_group <- file.path(results_root, "time_enrichment_group_summary_math_vs_lit.csv")
in_diff  <- file.path(results_root, "time_enrichment_diff_math_minus_lit_bootstrap_perm.csv")

out_png_group <- file.path(results_root, "fig_group_mean_lift_math_vs_lit.png")
out_png_diff  <- file.path(results_root, "fig_diff_lift_math_minus_lit.png")

if (!file.exists(in_group)) stop("Missing input group file:\n  ", in_group, call. = FALSE)
if (!file.exists(in_diff))  stop("Missing input diff file:\n  ", in_diff,  call. = FALSE)

message("[INFO] project_root: ", normalizePath(project_root, winslash="/", mustWork = FALSE))
message("[INFO] results_root: ", normalizePath(results_root, winslash="/", mustWork = FALSE))
message("[INFO] CI_MODE     : ", CI_MODE)
message("[INFO] USE_CENTERED: ", USE_CENTERED)

# ----------------------------
# Helpers
# ----------------------------
clean_cols <- function(nms) {
  nms %>%
    tolower() %>%
    str_replace_all("\\s+", "_") %>%
    str_replace_all("[^a-z0-9_\\.]", "_") %>%
    str_replace_all("_+", "_") %>%
    str_trim()
}

force_unique_cols <- function(df) {
  names(df) <- make.names(names(df), unique = TRUE)
  df
}

has_cols <- function(df, cols) all(cols %in% names(df))

pick_ci_cols_group <- function(df, mode = "auto") {
  boot   <- c("boot_lo", "boot_hi")
  legacy <- c("lift_ci95_lo", "lift_ci95_hi")
  tci    <- c("lift_t_lo", "lift_t_hi")
  
  if (mode == "boot") {
    if (!has_cols(df, boot)) stop("CI_MODE='boot' but boot_lo/boot_hi not found.", call. = FALSE)
    return(boot)
  }
  if (mode == "t") {
    if (has_cols(df, tci)) return(tci)
    if (has_cols(df, legacy)) return(legacy)
    stop("CI_MODE='t' but no t-based CI columns found.", call. = FALSE)
  }
  
  if (has_cols(df, boot)) return(boot)
  if (has_cols(df, tci)) return(tci)
  if (has_cols(df, legacy)) return(legacy)
  
  stop("No CI columns found in group file.", call. = FALSE)
}

pick_ci_cols_diff <- function(df) {
  lo <- grep("^diff_ci95_lo(\\.|$)", names(df), value = TRUE)
  hi <- grep("^diff_ci95_hi(\\.|$)", names(df), value = TRUE)
  if (length(lo) >= 1 && length(hi) >= 1) return(c(lo[[1]], hi[[1]]))
  
  lo2 <- grep("^(ci_lo|boot_lo|lo)(\\.|$)", names(df), value = TRUE)
  hi2 <- grep("^(ci_hi|boot_hi|hi)(\\.|$)", names(df), value = TRUE)
  if (length(lo2) >= 1 && length(hi2) >= 1) return(c(lo2[[1]], hi2[[1]]))
  
  stop("No CI columns found in diff file.", call. = FALSE)
}

# NEW: pick y column (centered preferred if requested)
pick_y_col <- function(df, centered_name, raw_name, use_centered = TRUE) {
  if (use_centered && centered_name %in% names(df)) return(centered_name)
  if (raw_name %in% names(df)) return(raw_name)
  stop("Missing required metric column. Tried: ", centered_name, " and ", raw_name, call. = FALSE)
}

theme_mainlike <- theme_minimal(base_size = 12) +
  theme(
    plot.title    = element_text(face = "bold", size = 14, color = "grey10"),
    plot.subtitle = element_text(size = 10, color = "grey35"),
    axis.title    = element_text(size = 11, color = "grey15"),
    axis.text     = element_text(size = 10, color = "grey20"),
    panel.grid.minor = element_blank(),
    panel.grid.major.x = element_blank(),
    panel.grid.major.y = element_line(linewidth = 0.3, color = "grey85"),
    strip.text = element_text(face = "bold", color = "grey15"),
    legend.position = "none",
    plot.margin = margin(10, 12, 10, 12)
  )

COL_DARK <- "grey15"
COL_MID  <- "grey35"

# ============================================================
# (1) GROUP SUMMARY PLOT
# ============================================================
df_g <- read_csv(in_group, show_col_types = FALSE)
names(df_g) <- clean_cols(names(df_g))
df_g <- force_unique_cols(df_g)

stopifnot("discipline" %in% names(df_g))

# NEW: choose y
y_col_g <- pick_y_col(df_g, "lift_mean_centered", "lift_mean", use_centered = USE_CENTERED)

ci_cols <- pick_ci_cols_group(df_g, mode = CI_MODE)
ci_lo_name <- ci_cols[[1]]
ci_hi_name <- ci_cols[[2]]

if (!("pointset" %in% names(df_g)))  df_g$pointset  <- "all_points"
if (!("null_mode" %in% names(df_g))) df_g$null_mode <- "global"
if (!("window_s" %in% names(df_g)))  df_g$window_s  <- NA_character_

baseline_g <- if (USE_CENTERED && y_col_g == "lift_mean_centered") 0 else 1
ylabel_g <- if (USE_CENTERED && y_col_g == "lift_mean_centered") "Mean lift (centered)" else "Mean lift"

df_g <- df_g %>%
  mutate(
    discipline = factor(discipline, levels = c("math", "lit")),
    pointset   = factor(pointset),
    null_mode  = factor(null_mode),
    y = suppressWarnings(as.numeric(.data[[y_col_g]])),
    ci_lo = suppressWarnings(as.numeric(.data[[ci_lo_name]])),
    ci_hi = suppressWarnings(as.numeric(.data[[ci_hi_name]])),
    ci_kind = case_when(
      ci_lo_name == "boot_lo" ~ "Bootstrap 95% CI",
      ci_lo_name %in% c("lift_t_lo", "lift_ci95_lo") ~ "t-based 95% CI",
      TRUE ~ "95% CI"
    )
  )

p1 <- ggplot(df_g, aes(x = discipline, y = y)) +
  geom_hline(yintercept = baseline_g, linewidth = 0.3, color = "grey75") +
  geom_errorbar(aes(ymin = ci_lo, ymax = ci_hi), width = 0.15, linewidth = 0.55, color = COL_DARK) +
  geom_point(size = 2.8, color = COL_DARK) +
  facet_grid(null_mode ~ pointset) +
  labs(
    title = "Group mean lift (video-level) by discipline",
    subtitle = paste0("Metric: ", y_col_g, " | Error bars: ", unique(df_g$ci_kind)[1]),
    x = NULL,
    y = ylabel_g
  ) +
  theme_mainlike +
  theme(axis.text.x = element_text(face = "bold"))

print(p1)

# ============================================================
# (2) DIFF PLOT (math − lit)
# ============================================================
df_d <- read_csv(in_diff, show_col_types = FALSE)
names(df_d) <- clean_cols(names(df_d))
df_d <- force_unique_cols(df_d)

# metric filter if exists
metric_candidates <- c("metric", "measure", "outcome", "target")
metric_col <- metric_candidates[metric_candidates %in% names(df_d)]
metric_col <- if (length(metric_col) >= 1) metric_col[[1]] else NA_character_

if (!is.na(metric_col)) {
  df_d <- df_d %>% mutate(.metric_tmp = tolower(as.character(.data[[metric_col]])))
  if (any(df_d$.metric_tmp == "lift", na.rm = TRUE)) {
    df_d <- df_d %>% filter(.metric_tmp == "lift")
  } else if (any(str_detect(df_d$.metric_tmp, "lift"), na.rm = TRUE)) {
    df_d <- df_d %>% filter(str_detect(.metric_tmp, "lift"))
  } else {
    message("[WARN] metric column exists but no lift rows found. Keeping all rows.")
  }
  df_d <- df_d %>% select(-.metric_tmp)
}

stopifnot("perm_p" %in% names(df_d))

# NEW: choose y
y_col_d <- pick_y_col(df_d, "diff_mean_centered", "diff_mean", use_centered = USE_CENTERED)

if (!("pointset" %in% names(df_d)))  df_d$pointset  <- "all_points"
if (!("null_mode" %in% names(df_d))) df_d$null_mode <- "global"
if (!("window_s" %in% names(df_d)))  df_d$window_s  <- NA_character_

ci_diff_cols <- pick_ci_cols_diff(df_d)
d_lo <- ci_diff_cols[[1]]
d_hi <- ci_diff_cols[[2]]

baseline_d <- 0
ylabel_d <- if (USE_CENTERED && y_col_d == "diff_mean_centered") {
  "Diff in mean lift (math − lit, centered)"
} else {
  "Diff in mean lift (math − lit)"
}

df_d <- df_d %>%
  mutate(
    pointset  = factor(pointset),
    null_mode = factor(null_mode),
    window_s  = as.character(window_s),
    strata_label = ifelse(is.na(window_s) | window_s == "NA", "w=NA", paste0("w=", window_s)),
    y = suppressWarnings(as.numeric(.data[[y_col_d]])),
    ci_lo = suppressWarnings(as.numeric(.data[[d_lo]])),
    ci_hi = suppressWarnings(as.numeric(.data[[d_hi]])),
    sig = case_when(
      is.na(perm_p) ~ "NA",
      perm_p < 0.001 ~ "p<.001",
      perm_p < 0.01  ~ "p<.01",
      perm_p < 0.05  ~ "p<.05",
      TRUE ~ "n.s."
    )
  )

p2 <- ggplot(df_d, aes(x = strata_label, y = y)) +
  geom_hline(yintercept = baseline_d, linewidth = 0.35, color = "grey70") +
  geom_errorbar(aes(ymin = ci_lo, ymax = ci_hi), width = 0.15, linewidth = 0.55, color = COL_DARK) +
  geom_point(size = 2.8, color = COL_DARK) +
  geom_text(aes(label = sig), vjust = -0.8, size = 3.2, color = COL_MID) +
  facet_grid(null_mode ~ pointset) +
  labs(
    title = "Difference in mean lift (math − lit)",
    subtitle = paste0("Metric: ", y_col_d, " | Error bars: bootstrap 95% CI; label: permutation p"),
    x = "Stratum (window)",
    y = ylabel_d
  ) +
  theme_mainlike +
  theme(axis.text.x = element_text(face = "bold"))

print(p2)

# ============================================================
# (3) SAVE PNG
# ============================================================
if (SAVE_PNG) {
  ggsave(out_png_group, p1, width = 8, height = 6, dpi = DPI, bg = "white")
  message("[OK] Saved PNG: ", normalizePath(out_png_group, winslash="/", mustWork = FALSE))
  
  ggsave(out_png_diff,  p2, width = 8, height = 6, dpi = DPI, bg = "white")
  message("[OK] Saved PNG: ", normalizePath(out_png_diff, winslash="/", mustWork = FALSE))
}

cat("\nDONE ✅\n")
cat("Group file:", in_group, "\n")
cat("Diff  file:", in_diff, "\n")
