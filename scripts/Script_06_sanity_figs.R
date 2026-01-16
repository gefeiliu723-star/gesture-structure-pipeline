# ============================================================
# Script 6: Sanity Figures (Methodological QC) -- FINAL CLEAN
#
# PURPOSE:
#   - Figure A: Confidence distribution of structural points
#   - Figure B: Gesture event density near structural points vs random times
#
# DATA SOURCE (ONLY):
#   exports/<TAG>/<TAG>_alignment.xlsx
#     - Sheet: Alignment (preferred; has Confidence)
#       fallback: StructuralPoints
#     - Sheet: GestureEvents
#
# CLEAN RULES:
#   - TAG is REQUIRED (no default)
#   - No absolute paths
#   - Root from env GESTURE_PROJECT_ROOT (fallback: getwd())
#   - No template usage
#
# Usage:
#   Rscript scripts/06_sanity_figs.R <TAG> [w] [n_null] [seed]
# Example:
#   Rscript scripts/06_sanity_figs.R m02 1.5 500 123
# ============================================================

suppressPackageStartupMessages({
  library(openxlsx)
  library(dplyr)
  library(ggplot2)
  library(tibble)
  library(scales)
})

# ----------------------------
# (0) SETTINGS / CLI  (TAG REQUIRED)
# ----------------------------
args <- commandArgs(trailingOnly = TRUE)

if (length(args) < 1 || !nzchar(args[[1]])) {
  stop(
    "Missing TAG.\n",
    "Usage:\n  Rscript scripts/06_sanity_figs.R <TAG> [w] [n_null] [seed]\n",
    call. = FALSE
  )
}

TAG    <- args[[1]]
w      <- if (length(args) >= 2 && nzchar(args[[2]])) as.numeric(args[[2]]) else 1.5
n_null <- if (length(args) >= 3 && nzchar(args[[3]])) as.integer(args[[3]]) else 500L
seed   <- if (length(args) >= 4 && nzchar(args[[4]])) as.integer(args[[4]]) else 123L

stopifnot(is.finite(w), w > 0)
stopifnot(is.finite(n_null), n_null >= 100)

set.seed(seed)

project_root <- Sys.getenv("GESTURE_PROJECT_ROOT", unset = getwd())
exports_root <- file.path(project_root, "exports")
results_root <- file.path(project_root, "results")

xlsx_path <- file.path(exports_root, TAG, paste0(TAG, "_alignment.xlsx"))
if (!file.exists(xlsx_path)) stop("Missing alignment file: ", xlsx_path, call. = FALSE)

out_dir <- file.path(results_root, TAG, "sanity_figs")
dir.create(out_dir, recursive = TRUE, showWarnings = FALSE)

# ----------------------------
# (1) HELPERS
# ----------------------------
pick_col <- function(df, candidates) {
  nms <- names(df)
  for (c in candidates) {
    hit <- nms[tolower(nms) == tolower(c)]
    if (length(hit) > 0) return(hit[[1]])
  }
  NA_character_
}

as_num <- function(x) suppressWarnings(as.numeric(as.character(x)))

count_events_in_windows <- function(event_t, centers, w) {
  vapply(
    centers,
    function(c0) sum(abs(event_t - c0) <= w, na.rm = TRUE),
    integer(1)
  )
}

# ----------------------------
# (2) READ DATA
# ----------------------------
sheets <- getSheetNames(xlsx_path)

# Structural points
if ("Alignment" %in% sheets) {
  pts <- read.xlsx(xlsx_path, sheet = "Alignment")
} else if ("StructuralPoints" %in% sheets) {
  pts <- read.xlsx(xlsx_path, sheet = "StructuralPoints")
} else {
  stop("No Alignment or StructuralPoints sheet found in: ", xlsx_path, call. = FALSE)
}

# Gesture events
if (!("GestureEvents" %in% sheets)) {
  stop("GestureEvents sheet not found in: ", xlsx_path, call. = FALSE)
}
ev <- read.xlsx(xlsx_path, sheet = "GestureEvents")

# ----------------------------
# (3) EXTRACT STRUCTURAL TIMES + CONFIDENCE
# ----------------------------
tcol_pts <- pick_col(pts, c("t_struct_sec", "time_sec", "time_s", "t_sec", "time"))
ccol_pts <- pick_col(pts, c("confidence_1to3", "confidence", "conf"))

if (is.na(tcol_pts)) stop("No structural time column found in Alignment/StructuralPoints.", call. = FALSE)

pts <- pts %>%
  mutate(
    t_struct = as_num(.data[[tcol_pts]]),
    conf     = if (!is.na(ccol_pts)) as_num(.data[[ccol_pts]]) else NA_real_
  ) %>%
  filter(is.finite(t_struct))

pts_conf <- pts %>% filter(is.finite(conf))

# ----------------------------
# (4) EXTRACT EVENT TIMES
# ----------------------------
tcol_ev <- pick_col(ev, c("peak_sec", "t_peak", "peak", "t_peak_sec", "start_sec"))
if (is.na(tcol_ev)) stop("No event peak/start time column found in GestureEvents.", call. = FALSE)

ev <- ev %>%
  mutate(t_event = as_num(.data[[tcol_ev]])) %>%
  filter(is.finite(t_event))

if (nrow(ev) == 0) stop("No valid event times found in GestureEvents.", call. = FALSE)
if (nrow(pts) == 0) stop("No valid structural times found in Alignment/StructuralPoints.", call. = FALSE)

T_total <- max(c(ev$t_event, pts$t_struct), na.rm = TRUE)

# ============================================================
# FIGURE A: CONFIDENCE DISTRIBUTION
# ============================================================
if (nrow(pts_conf) > 0) {
  conf_tab <- pts_conf %>%
    mutate(conf = factor(conf, levels = c(1, 2, 3))) %>%
    count(conf) %>%
    mutate(prop = n / sum(n))
  
  pA <- ggplot(conf_tab, aes(x = conf, y = prop)) +
    geom_col() +
    scale_y_continuous(labels = percent_format(accuracy = 1)) +
    labs(
      title = paste0("Figure A. Confidence distribution (", TAG, ")"),
      x = "Confidence level",
      y = "Proportion"
    )
  
  ggsave(
    filename = file.path(out_dir, paste0(TAG, "_FigA_conf_distribution.png")),
    plot = pA, width = 6.5, height = 4.2, dpi = 160
  )
}

# ============================================================
# FIGURE B: STRUCTURAL vs RANDOM EVENT DENSITY
# ============================================================
struct_counts <- count_events_in_windows(ev$t_event, pts$t_struct, w)

# random centers avoid borders so window stays inside [0, T_total]
hi <- max(w, T_total - w)
null_centers <- runif(n_null, min = w, max = hi)
null_counts  <- count_events_in_windows(ev$t_event, null_centers, w)

dfB <- tibble(
  type = c(
    rep("Structural points", length(struct_counts)),
    rep("Random points", length(null_counts))
  ),
  count_in_window = c(struct_counts, null_counts)
)

pB <- ggplot(dfB, aes(x = type, y = count_in_window)) +
  geom_boxplot(outlier.alpha = 0.25) +
  labs(
    title = paste0("Figure B. Event density within ±", w, " s (", TAG, ")"),
    x = NULL,
    y = "Number of gesture events"
  )

ggsave(
  filename = file.path(out_dir, paste0(TAG, "_FigB_struct_vs_random.png")),
  plot = pB, width = 7.2, height = 4.2, dpi = 160
)

# ----------------------------
# (5) CONSOLE SUMMARY
# ----------------------------
cat("\n--- SANITY CHECK SUMMARY ---\n")
cat("TAG:", TAG, "\n")
cat("Alignment file:", normalizePath(xlsx_path, winslash = "/"), "\n")
cat("Structural points (usable):", nrow(pts), "\n")
cat("Gesture events (usable):", nrow(ev), "\n")

if (nrow(pts_conf) > 0) {
  cat("\nConfidence distribution:\n")
  print(conf_tab)
} else {
  cat("\n(No confidence column found / no finite confidence values)\n")
}

cat("\nMedian events in ±w window:\n")
cat("  Structural:", median(struct_counts), "\n")
cat("  Random    :", median(null_counts), "\n")

cat("\nSaved to:", normalizePath(out_dir, winslash = "/"), "\n")
cat("DONE:", TAG, "\n")
