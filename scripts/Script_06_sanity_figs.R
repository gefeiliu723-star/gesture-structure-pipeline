# ============================================================
# Script 05a: Sanity Figures (Methodological QC)
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
# NOTE:
#   This script is for diagnostic sanity checks only.
#   Outputs are NOT used as main statistical results.
# ============================================================

suppressPackageStartupMessages({
  library(openxlsx)
  library(dplyr)
  library(ggplot2)
  library(stringr)
  library(tibble)
  library(scales)
})

# ----------------------------
# (0) SETTINGS / CLI
# ----------------------------
args <- commandArgs(trailingOnly = TRUE)

TAG   <- if (length(args) >= 1) args[[1]] else "m01"
w     <- if (length(args) >= 2) as.numeric(args[[2]]) else 1.5   # half-window (sec)
n_null <- if (length(args) >= 3) as.integer(args[[3]]) else 500L
seed  <- if (length(args) >= 4) as.integer(args[[4]]) else 123L

stopifnot(is.finite(w), w > 0)
stopifnot(n_null >= 100)

set.seed(seed)

project_root <- Sys.getenv("GESTURE_PROJECT_ROOT", unset = getwd())
exports_root <- file.path(project_root, "exports")
results_root <- file.path(project_root, "results")

xlsx_path <- file.path(exports_root, TAG, paste0(TAG, "_alignment.xlsx"))
if (!file.exists(xlsx_path)) stop("Missing alignment file: ", xlsx_path)

out_dir <- file.path(results_root, TAG, "sanity_figs")
dir.create(out_dir, recursive = TRUE, showWarnings = FALSE)

# ----------------------------
# (1) HELPERS
# ----------------------------
pick_col <- function(df, candidates) {
  nms <- names(df)
  for (c in candidates) {
    hit <- nms[tolower(nms) == tolower(c)]
    if (length(hit) > 0) return(hit[1])
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

# Structural points (Alignment preferred)
if ("Alignment" %in% sheets) {
  pts <- read.xlsx(xlsx_path, sheet = "Alignment")
} else if ("StructuralPoints" %in% sheets) {
  pts <- read.xlsx(xlsx_path, sheet = "StructuralPoints")
} else {
  stop("No Alignment or StructuralPoints sheet found.")
}

# Gesture events
if (!("GestureEvents" %in% sheets)) {
  stop("GestureEvents sheet not found.")
}
ev <- read.xlsx(xlsx_path, sheet = "GestureEvents")

# ----------------------------
# (3) EXTRACT STRUCTURAL TIMES + CONFIDENCE
# ----------------------------
tcol_pts <- pick_col(pts, c("t_struct_sec", "time_sec", "time_s", "t_sec", "time"))
ccol_pts <- pick_col(pts, c("confidence", "conf"))

if (is.na(tcol_pts)) stop("No structural time column found.")

pts <- pts %>%
  mutate(
    t_struct = as_num(.data[[tcol_pts]]),
    conf     = if (!is.na(ccol_pts)) as_num(.data[[ccol_pts]]) else NA_real_
  ) %>%
  filter(is.finite(t_struct))

# Only keep rows with confidence for Fig A
pts_conf <- pts %>% filter(is.finite(conf))

# ----------------------------
# (4) EXTRACT EVENT TIMES
# ----------------------------
tcol_ev <- pick_col(ev, c("peak_sec", "t_peak", "peak", "t_peak_sec", "start_sec"))
if (is.na(tcol_ev)) stop("No event peak time column found.")

ev <- ev %>%
  mutate(t_event = as_num(.data[[tcol_ev]])) %>%
  filter(is.finite(t_event))

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
    file.path(out_dir, paste0(TAG, "_FigA_conf_distribution.png")),
    pA, width = 6.5, height = 4.2, dpi = 160
  )
}

# ============================================================
# FIGURE B: STRUCTURAL vs RANDOM EVENT DENSITY
# ============================================================
struct_counts <- count_events_in_windows(ev$t_event, pts$t_struct, w)

null_centers <- runif(n_null, min = w, max = max(w, T_total - w))
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
  file.path(out_dir, paste0(TAG, "_FigB_struct_vs_random.png")),
  pB, width = 7.2, height = 4.2, dpi = 160
)

# ----------------------------
# (5) CONSOLE SUMMARY
# ----------------------------
cat("\n--- SANITY CHECK SUMMARY ---\n")
cat("TAG:", TAG, "\n")
cat("Structural points:", nrow(pts), "\n")
if (nrow(pts_conf) > 0) print(conf_tab)
cat("Median events (struct):", median(struct_counts), "\n")
cat("Median events (random):", median(null_counts), "\n")
cat("Saved to:", normalizePath(out_dir, winslash = "/"), "\n")
