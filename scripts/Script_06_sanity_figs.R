# ============================================================
# Script 06 (SANITY FIGURES / GitHub-ready)
#   Fig A: Confidence distribution (if available)
#   Fig B: Event density near structural points vs random times
#
# INPUT (repo layout):
#   exports/**/<TAG>_alignment.xlsx
#     - sheet: Alignment (preferred) or StructuralPoints
#     - sheet: GestureEvents
#
# OUTPUT:
#   results/<TAG>/sanity_figs/
#
# CLI:
#   Rscript scripts/06_sanity_figs.R l01 1.5 500 123
#     args: TAG, w, n_null, seed
#
# Env (optional):
#   GESTURE_PROJECT_ROOT=/path/to/repo_root
# ============================================================

suppressPackageStartupMessages({
  library(openxlsx)
  library(dplyr)
  library(tibble)
  library(stringr)
  library(ggplot2)
  library(scales)
})

# ----------------------------
# (0) SETTINGS / CLI
# ----------------------------
args <- commandArgs(trailingOnly = TRUE)

TAG_in  <- if (length(args) >= 1) args[[1]] else "l01"
w       <- if (length(args) >= 2) as.numeric(args[[2]]) else 1.5
n_null  <- if (length(args) >= 3) as.integer(args[[3]]) else 500L
seed    <- if (length(args) >= 4) as.integer(args[[4]]) else 123L

stopifnot(is.finite(w), w > 0)
stopifnot(is.finite(n_null), n_null >= 10)

set.seed(seed)

# canonical TAG: underscores
TAG_raw <- str_trim(TAG_in)
TAG <- str_replace_all(TAG_raw, "-", "_")
TAG_dash <- str_replace_all(TAG, "_", "-")
if (TAG_raw != TAG) message("TAG normalized: ", TAG_raw, " -> ", TAG)

# project root (portable)
project_root <- Sys.getenv("GESTURE_PROJECT_ROOT", unset = getwd())
exports_root <- file.path(project_root, "exports")
results_root <- file.path(project_root, "results")

if (!dir.exists(exports_root)) {
  stop("exports/ folder not found. Run from repo root or set GESTURE_PROJECT_ROOT.", call. = FALSE)
}

out_dir <- file.path(results_root, TAG, "sanity_figs")
dir.create(out_dir, recursive = TRUE, showWarnings = FALSE)

# ----------------------------
# (1) HELPERS
# ----------------------------
clean_names <- function(nms) nms %>% str_replace_all("\\s+", " ") %>% str_trim()

pick_col <- function(df, candidates) {
  nm <- names(df)
  if (is.null(nm) || length(nm) == 0) return(NA_character_)
  low <- tolower(nm)
  for (cand in candidates) {
    j <- which(low == tolower(cand))
    if (length(j) >= 1) return(nm[j[[1]]])
  }
  NA_character_
}

need_col <- function(df, candidates, where = "data") {
  col <- pick_col(df, candidates)
  if (is.na(col)) {
    stop(
      "Cannot find required column in ", where, ". Tried:\n  - ",
      paste(candidates, collapse = ", "),
      "\nAvailable columns:\n  - ",
      paste(names(df), collapse = ", "),
      call. = FALSE
    )
  }
  col
}

as_num <- function(x) suppressWarnings(as.numeric(as.character(x)))

count_events_in_windows <- function(event_t, centers, w) {
  vapply(centers, function(c0) sum(abs(event_t - c0) <= w, na.rm = TRUE), integer(1))
}

# Find alignment xlsx robustly under exports/
find_alignment_xlsx <- function(exports_root, TAG, TAG_dash) {
  all_align <- list.files(
    path = exports_root,
    pattern = "_alignment\\.xlsx$",
    recursive = TRUE,
    full.names = TRUE
  )
  if (length(all_align) == 0) stop("No *_alignment.xlsx found under exports/.", call. = FALSE)
  
  targets <- c(paste0(TAG, "_alignment.xlsx"), paste0(TAG_dash, "_alignment.xlsx"))
  hit <- all_align[basename(all_align) %in% targets]
  
  if (length(hit) == 0) {
    # folder match fallback
    hit <- all_align[basename(dirname(all_align)) %in% c(TAG, TAG_dash)]
  }
  
  if (length(hit) == 0) {
    stop(
      "Cannot find alignment xlsx for TAG=", TAG, "\n",
      "Tried: ", paste(targets, collapse = ", "), "\n",
      "Under: ", exports_root,
      call. = FALSE
    )
  }
  if (length(hit) > 1) {
    warning("Multiple alignment files matched. Using first:\n", hit[[1]])
  }
  hit[[1]]
}

# Paper-ish theme (clean, compact, consistent)
theme_paper <- theme_classic(base_size = 12) +
  theme(
    plot.title    = element_text(face = "bold", size = 13),
    plot.subtitle = element_text(size = 10, color = "grey30"),
    plot.caption  = element_text(size = 9, color = "grey35"),
    axis.title    = element_text(size = 11),
    axis.text     = element_text(size = 10, color = "grey15"),
    axis.line     = element_line(linewidth = 0.4),
    axis.ticks    = element_line(linewidth = 0.4),
    legend.position = "none"
  )

# ----------------------------
# (2) LOCATE + READ XLSX
# ----------------------------
xlsx_path <- find_alignment_xlsx(exports_root, TAG, TAG_dash)
message("Using xlsx: ", normalizePath(xlsx_path, winslash = "/", mustWork = FALSE))

sheets <- openxlsx::getSheetNames(xlsx_path)

pts_sheet <- if ("Alignment" %in% sheets) "Alignment" else if ("StructuralPoints" %in% sheets) "StructuralPoints" else NA_character_
if (is.na(pts_sheet)) stop("No Alignment or StructuralPoints sheet found in xlsx.", call. = FALSE)

if (!("GestureEvents" %in% sheets)) stop("No GestureEvents sheet found in xlsx.", call. = FALSE)

pts <- openxlsx::read.xlsx(xlsx_path, sheet = pts_sheet)
names(pts) <- clean_names(names(pts))

ev <- openxlsx::read.xlsx(xlsx_path, sheet = "GestureEvents")
names(ev) <- clean_names(names(ev))

# ----------------------------
# (3) STRUCTURAL TIMES (+ optional confidence)
# ----------------------------
tcol_pts <- need_col(
  pts,
  c("t_struct_sec", "t_struct_s", "time_sec", "time_s", "t_sec", "time", "t"),
  where = paste0(pts_sheet, " sheet")
)

ccol_pts <- pick_col(pts, c("Confidence", "confidence", "conf", "Conf", "CONF"))

pts2 <- pts %>%
  mutate(
    t_struct = as_num(.data[[tcol_pts]]),
    conf     = if (!is.na(ccol_pts)) as_num(.data[[ccol_pts]]) else NA_real_
  ) %>%
  filter(is.finite(t_struct))

if (nrow(pts2) == 0) stop("No valid structural times after parsing.", call. = FALSE)

# ----------------------------
# (4) EVENT PEAK TIMES
# ----------------------------
tcol_ev <- need_col(
  ev,
  c(
    "peak_sec", "t_sec",
    "t_peak", "tpeak", "t_peak_sec",
    "peak_sec_proxy",
    "peak", "peak_s", "peak_time",
    "start_sec"
  ),
  where = "GestureEvents sheet"
)

ev2 <- ev %>%
  mutate(t_event = as_num(.data[[tcol_ev]])) %>%
  filter(is.finite(t_event))

if (nrow(ev2) == 0) stop("No valid event times after parsing GestureEvents.", call. = FALSE)

T_total <- max(c(ev2$t_event, pts2$t_struct), na.rm = TRUE)
if (!is.finite(T_total) || T_total <= 2*w) stop("T_total invalid/too small. T_total=", T_total, " w=", w, call. = FALSE)

# ============================================================
# FIGURE A: confidence distribution
# ============================================================
if (!all(is.na(pts2$conf))) {
  
  conf_tab <- pts2 %>%
    filter(is.finite(conf)) %>%
    mutate(conf = factor(conf, levels = c(1, 2, 3), labels = c("1", "2", "3"))) %>%
    count(conf) %>%
    mutate(prop = n / sum(n),
           lab  = percent(prop, accuracy = 1))
  
  pA <- ggplot(conf_tab, aes(x = conf, y = prop)) +
    geom_col(fill = "#2C7FB8", width = 0.68) +
    geom_text(aes(label = lab), vjust = -0.4, size = 3.4, color = "grey15") +
    scale_y_continuous(
      labels = percent_format(accuracy = 1),
      limits = c(0, max(conf_tab$prop) * 1.18),
      expand = expansion(mult = c(0, 0.02))
    ) +
    labs(
      title = paste0("Figure A. Confidence distribution (", TAG, ")"),
      subtitle = paste0("Points from: ", pts_sheet, " | N = ", sum(conf_tab$n)),
      x = "Confidence",
      y = "Proportion"
    ) +
    theme_paper
  
  ggsave(
    filename = file.path(out_dir, paste0(TAG, "_FigA_conf_distribution_paper.png")),
    plot = pA, width = 6.2, height = 3.8, dpi = 300
  )
} else {
  message("No Confidence column found. Skipping Figure A.")
}

# ============================================================
# FIGURE B: struct vs random event counts within ±w windows
# ============================================================
struct_counts <- count_events_in_windows(ev2$t_event, pts2$t_struct, w)
null_centers  <- runif(n_null, min = w, max = T_total - w)
null_counts   <- count_events_in_windows(ev2$t_event, null_centers, w)

dfB <- tibble(
  group = factor(
    c(rep("Structural points", length(struct_counts)),
      rep("Random times",      length(null_counts))),
    levels = c("Random times", "Structural points")
  ),
  count = c(struct_counts, null_counts)
)

sumB <- dfB %>%
  group_by(group) %>%
  summarize(
    med = median(count),
    mean = mean(count),
    .groups = "drop"
  )

fill_map  <- c("Random times" = "#C9C9C9", "Structural points" = "#F39C6B")
line_map  <- c("Random times" = "#7A7A7A", "Structural points" = "#B24B1E")

pB <- ggplot(dfB, aes(x = group, y = count)) +
  geom_boxplot(aes(fill = group),
               width = 0.55,
               outlier.shape = NA,
               alpha = 0.95,
               color = "grey25",
               linewidth = 0.35) +
  geom_jitter(aes(color = group),
              width = 0.14,
              height = 0,
              size = 1.3,
              alpha = 0.22) +
  geom_point(data = sumB, aes(x = group, y = med),
             inherit.aes = FALSE,
             shape = 16, size = 2.4, color = "grey10") +
  geom_point(data = sumB, aes(x = group, y = mean),
             inherit.aes = FALSE,
             shape = 23, size = 2.6, fill = "white", color = "grey10", stroke = 0.45) +
  scale_fill_manual(values = fill_map) +
  scale_color_manual(values = line_map) +
  labs(
    title = paste0("Figure B. Gesture-event density near structural points (", TAG, ")"),
    subtitle = paste0("Count of events within ±", w, "s windows | Null N = ", n_null,
                      " | Median = ●, Mean = ◇"),
    x = NULL,
    y = "Events per window"
  ) +
  theme_paper +
  theme(axis.text.x = element_text(face = "bold"))

ggsave(
  filename = file.path(out_dir, paste0(TAG, "_FigB_struct_vs_random_paper.png")),
  plot = pB, width = 7.0, height = 4.0, dpi = 300
)

# ----------------------------
# QUICK PRINT (minimal; avoids local paths)
# ----------------------------
cat("\n--- SANITY FIGURES DONE ---\n")
cat("TAG:", TAG, "\n")
cat("points sheet:", pts_sheet, "\n")
cat("Structural points:", nrow(pts2), "\n")
cat("Events:", nrow(ev2), "\n")
cat("T_total:", signif(T_total, 6), "\n")
cat("Median events (struct):", median(struct_counts), "\n")
cat("Median events (random):", median(null_counts), "\n")
cat("Saved to: results/", TAG, "/sanity_figs/\n", sep = "")
