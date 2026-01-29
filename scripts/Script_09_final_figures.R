# ============================================================
# Script 09 (FINAL / PAPER-LEVEL FIGURES + ROBUST):
#   Within-teacher stability + teacher-level domain comparison
#
# Inputs (from Script 07):
#   results/time_enrichment_MASTER_ALL.csv
#
# Outputs (results/teacher_level/):
#   - teacher_level_summary.csv
#   - domain_diff_teacher_level.csv
#   - fig_signature_within_teacher.png
#   - fig_teacher_level_domain_means.png
#
# Usage:
#   Rscript scripts/09_teacher_level.R
#
# Optional env var:
#   GESTURE_PROJECT_ROOT=/path/to/repo_root
# ============================================================

suppressPackageStartupMessages({
  library(readr)
  library(dplyr)
  library(stringr)
  library(tidyr)
  library(ggplot2)
})

# ----------------------------
# (0) USER SETTINGS (EDIT ONLY HERE)
# ----------------------------

# Which subset to visualize / analyze
# Use NULL to keep all values present in MASTER_ALL.
teachers_keep   <- NULL              # e.g., c("m01","m02","m03","l01","l02","l03")
pointsets_keep  <- c("conf3_only")   # e.g., c("all_points","conf1_only","conf2_only","conf3_only")
null_modes_keep <- c("teacher_only") # e.g., c("global","teacher_only")
window_keep     <- 1.5               # numeric; set NULL to keep all windows

# Plot mode:
#   "single" = one clean panel for MAIN text
#   "facet"  = facet by pointset/null_mode (supp figure style)
PLOT_MODE <- "single"

# Bootstrap / permutation replicates
B <- 5000L
P <- 5000L
set.seed(123)

# Use centered lift if available (auto fallback to raw lift)
USE_CENTERED <- TRUE   # TRUE = prefer *_centered columns

# Save figs
DPI <- 300
BG  <- "white"

# ----------------------------
# (0.1) PATHS (repo-root portable)
# ----------------------------
project_root <- Sys.getenv("GESTURE_PROJECT_ROOT", unset = getwd())
results_root <- file.path(project_root, "results")
out_dir      <- file.path(results_root, "teacher_level")
dir.create(out_dir, recursive = TRUE, showWarnings = FALSE)

master_csv <- file.path(results_root, "time_enrichment_MASTER_ALL.csv")
if (!file.exists(master_csv)) {
  stop(
    "Missing input:\n  ", master_csv, "\n",
    "Expected: results/time_enrichment_MASTER_ALL.csv (from Script 07).",
    call. = FALSE
  )
}

message("[INFO] project_root: ", normalizePath(project_root, winslash="/", mustWork = FALSE))
message("[INFO] master_csv   : ", normalizePath(master_csv, winslash="/", mustWork = FALSE))
message("[INFO] out_dir      : ", normalizePath(out_dir, winslash="/", mustWork = FALSE))

# ----------------------------
# (1) Load MASTER_ALL
# ----------------------------
df <- read_csv(master_csv, show_col_types = FALSE)
names(df) <- tolower(names(df))

stopifnot("tag" %in% names(df))
stopifnot("lift" %in% names(df))

# ----------------------------
# (1.1) Choose lift column (centered vs raw)
# ----------------------------
lift_col <- if (USE_CENTERED && "lift_centered" %in% names(df)) {
  message("[INFO] Using lift_centered")
  "lift_centered"
} else {
  if (USE_CENTERED) message("[INFO] lift_centered not found. Falling back to lift.")
  "lift"
}

y_lab <- if (lift_col == "lift") "Lift" else "Centered lift (Δ from baseline)"
baseline_y <- if (lift_col == "lift") 1 else 0

# ----------------------------
# (2) Add teacher_id + discipline (ROBUST)
#     m04-2 / m04_2 / m04-3 / m04_3  -> m04
# ----------------------------
strip_suffix <- function(x) {
  str_replace(as.character(x), "([-_][0-9]+)$", "")
}

df2 <- df %>%
  mutate(
    tag = as.character(tag),
    teacher_id = strip_suffix(tag),
    discipline = case_when(
      str_starts(teacher_id, "m") ~ "math",
      str_starts(teacher_id, "l") ~ "lit",
      TRUE ~ NA_character_
    )
  )

if (any(is.na(df2$discipline))) {
  bad <- df2 %>% filter(is.na(discipline)) %>% distinct(tag, teacher_id)
  stop(
    "Some tags do not start with m/l after suffix stripping:\n",
    paste0(" - ", bad$tag, " (teacher_id=", bad$teacher_id, ")", collapse = "\n"),
    call. = FALSE
  )
}

# ----------------------------
# (3) Filter to target subset
# ----------------------------
df_sub <- df2

if (!is.null(teachers_keep)) df_sub <- df_sub %>% filter(teacher_id %in% teachers_keep)

if (!is.null(pointsets_keep) && "pointset" %in% names(df_sub)) {
  df_sub <- df_sub %>% filter(pointset %in% pointsets_keep)
}

if (!is.null(null_modes_keep) && "null_mode" %in% names(df_sub)) {
  df_sub <- df_sub %>% filter(null_mode %in% null_modes_keep)
}

# robust numeric filter for window_s
if (!is.null(window_keep) && "window_s" %in% names(df_sub)) {
  df_sub <- df_sub %>%
    mutate(window_s_num = suppressWarnings(as.numeric(window_s))) %>%
    filter(is.finite(window_s_num), abs(window_s_num - as.numeric(window_keep)) < 1e-9) %>%
    select(-window_s_num)
}

# Keep only ok rows if status exists
if ("status" %in% names(df_sub)) {
  df_sub <- df_sub %>% filter(status == "ok")
}

if (nrow(df_sub) == 0) stop("Filtered dataset is empty. Relax filters.", call. = FALSE)

# ----------------------------
# (4) Teacher-level summary (unit = teacher; aggregate across segments)
# ----------------------------
grp_cols <- intersect(c("teacher_id","discipline","pointset","window_s","null_mode"), names(df_sub))

teacher_summary <- df_sub %>%
  group_by(across(all_of(grp_cols))) %>%
  summarise(
    n_segments = n_distinct(tag),
    mean_lift  = mean(.data[[lift_col]], na.rm = TRUE),
    sd_lift    = ifelse(
      sum(is.finite(.data[[lift_col]])) >= 2,
      sd(.data[[lift_col]], na.rm = TRUE),
      NA_real_
    ),
    .groups = "drop"
  ) %>%
  arrange(across(all_of(intersect(c("pointset","null_mode","discipline","teacher_id"), names(.)))))

write_csv(teacher_summary, file.path(out_dir, "teacher_level_summary.csv"))
message("[OK] wrote teacher_level_summary.csv")

# ----------------------------
# (5) Domain comparison at TEACHER LEVEL (math vs lit; unit = teacher)
# ----------------------------
domain_test_one <- function(dat, B = 5000L, P = 5000L) {
  stopifnot(all(c("teacher_id","discipline","mean_lift") %in% names(dat)))
  
  d_math <- dat %>% filter(discipline == "math") %>% pull(mean_lift)
  d_lit  <- dat %>% filter(discipline == "lit")  %>% pull(mean_lift)
  
  if (length(d_math) < 2 || length(d_lit) < 2) {
    return(tibble(
      n_math = length(d_math), n_lit = length(d_lit),
      diff_obs = NA_real_, boot_lo = NA_real_, boot_hi = NA_real_, perm_p = NA_real_
    ))
  }
  
  diff_obs <- mean(d_math) - mean(d_lit)
  
  boot <- replicate(B, mean(sample(d_math, replace = TRUE)) - mean(sample(d_lit, replace = TRUE)))
  boot_ci <- quantile(boot, probs = c(0.025, 0.975), na.rm = TRUE, type = 7)
  
  vals <- dat$mean_lift
  disc <- dat$discipline
  perm <- replicate(P, {
    perm_disc <- sample(disc, replace = FALSE)
    mean(vals[perm_disc == "math"]) - mean(vals[perm_disc == "lit"])
  })
  perm_p <- (sum(abs(perm) >= abs(diff_obs)) + 1) / (P + 1)
  
  tibble(
    n_math = length(d_math),
    n_lit  = length(d_lit),
    diff_obs = diff_obs,
    boot_lo = unname(boot_ci[[1]]),
    boot_hi = unname(boot_ci[[2]]),
    perm_p  = perm_p
  )
}

domain_group_cols <- intersect(c("pointset","window_s","null_mode"), names(teacher_summary))

domain_results <- teacher_summary %>%
  group_by(across(all_of(domain_group_cols))) %>%
  group_modify(~domain_test_one(.x, B = B, P = P)) %>%
  ungroup()

write_csv(domain_results, file.path(out_dir, "domain_diff_teacher_level.csv"))
message("[OK] wrote domain_diff_teacher_level.csv")

# ============================================================
# (6) PAPER-LEVEL SIGNATURE PLOT (within-teacher stability)
# ============================================================

seg_points <- df_sub %>%
  select(tag, teacher_id, discipline, any_of(c("pointset","window_s","null_mode"))) %>%
  mutate(lift = .data[[lift_col]]) %>%
  filter(is.finite(lift)) %>%
  mutate(type = "Segment (3 per teacher)")

teacher_means <- teacher_summary %>%
  select(teacher_id, discipline, any_of(c("pointset","window_s","null_mode")), mean_lift) %>%
  rename(lift = mean_lift) %>%
  filter(is.finite(lift)) %>%
  mutate(type = "Teacher mean")

# stable ordering: by discipline then teacher_id
teacher_order <- seg_points %>%
  distinct(teacher_id, discipline) %>%
  arrange(discipline, teacher_id) %>%
  pull(teacher_id)

seg_points    <- seg_points    %>% mutate(teacher_id = factor(teacher_id, levels = teacher_order))
teacher_means <- teacher_means %>% mutate(teacher_id = factor(teacher_id, levels = teacher_order))

# paper-ish theme
theme_paper <- theme_minimal(base_size = 12) +
  theme(
    plot.title    = element_text(face="bold", size=16, color="grey10"),
    axis.title    = element_text(size=13, color="grey15"),
    axis.text     = element_text(size=11, color="grey25"),
    panel.grid.minor = element_blank(),
    panel.grid.major.x = element_blank(),
    panel.grid.major.y = element_line(linewidth=0.35, color="grey88"),
    axis.line = element_line(linewidth=0.35, color="grey50"),
    plot.margin = margin(12, 14, 10, 14),
    legend.position = "top",
    legend.title = element_blank()
  )

df_sig <- bind_rows(seg_points, teacher_means) %>%
  mutate(type = factor(type, levels = c("Segment (3 per teacher)", "Teacher mean")))

p_sig <- ggplot(df_sig, aes(x = teacher_id, y = lift)) +
  geom_hline(yintercept = baseline_y, linewidth = 0.35, color = "grey70") +
  # segment points
  geom_point(
    data = df_sig %>% filter(type == "Segment (3 per teacher)"),
    aes(shape = type),
    position = position_jitter(width = 0.12, height = 0),
    size = 2.2, alpha = 0.35, color = "grey40"
  ) +
  # teacher means: hollow circles
  geom_point(
    data = df_sig %>% filter(type == "Teacher mean"),
    aes(shape = type),
    size = 3.8, stroke = 0.9, color = "black", fill = "white"
  ) +
  scale_shape_manual(values = c(
    "Segment (3 per teacher)" = 16,
    "Teacher mean" = 21
  )) +
  labs(
    title = "Within-teacher stability of gestural lift",
    x = "Teacher",
    y = y_lab
  ) +
  theme_paper +
  theme(axis.text.x = element_text(angle = 35, hjust = 1))

if (PLOT_MODE == "facet" && all(c("pointset","null_mode") %in% names(df_sig))) {
  p_sig <- p_sig + facet_grid(null_mode ~ pointset)
}

ggsave(
  filename = file.path(out_dir, "fig_signature_within_teacher.png"),
  plot = p_sig,
  width = 10.5,
  height = ifelse(PLOT_MODE == "facet", 6.5, 5.2),
  dpi = DPI,
  bg = BG
)
message("[OK] saved fig_signature_within_teacher.png")

# ============================================================
# (7) Teacher-level domain means plot (unit=teacher; bootstrap 95% CI)
# ============================================================
boot_ci_mean <- function(x, B = 5000L) {
  x <- x[is.finite(x)]
  if (length(x) == 0) return(c(mean = NA, lo = NA, hi = NA))
  m <- mean(x)
  if (length(x) < 2) return(c(mean = m, lo = NA, hi = NA))
  boot <- replicate(B, mean(sample(x, replace = TRUE)))
  ci <- quantile(boot, c(0.025, 0.975), na.rm = TRUE, type = 7)
  c(mean = m, lo = unname(ci[[1]]), hi = unname(ci[[2]]))
}

plot_group_cols <- intersect(c("pointset","null_mode","discipline"), names(teacher_summary))

domain_means_plot_df <- teacher_summary %>%
  group_by(across(all_of(plot_group_cols))) %>%
  summarise(tmp = list(boot_ci_mean(mean_lift, B = B)), .groups = "drop") %>%
  mutate(
    mean = sapply(tmp, `[[`, "mean"),
    lo   = sapply(tmp, `[[`, "lo"),
    hi   = sapply(tmp, `[[`, "hi")
  ) %>%
  select(-tmp) %>%
  mutate(discipline = factor(discipline, levels = c("math","lit")))

p_dom <- ggplot(domain_means_plot_df, aes(x = discipline, y = mean)) +
  geom_hline(yintercept = baseline_y, linewidth = 0.35, color = "grey70") +
  geom_errorbar(aes(ymin = lo, ymax = hi), width = 0.15, linewidth = 0.65, color = "grey15") +
  geom_point(size = 3.2, shape = 21, fill = "white", color = "black", stroke = 0.9) +
  labs(
    title = "Teacher-level mean lift by discipline",
    x = NULL,
    y = y_lab
  ) +
  theme_paper +
  theme(axis.text.x = element_text(face = "bold"))

if (all(c("pointset","null_mode") %in% names(domain_means_plot_df))) {
  p_dom <- p_dom + facet_grid(null_mode ~ pointset)
}

ggsave(
  filename = file.path(out_dir, "fig_teacher_level_domain_means.png"),
  plot = p_dom,
  width = 9.5,
  height = ifelse(all(c("pointset","null_mode") %in% names(domain_means_plot_df)), 6.0, 4.8),
  dpi = DPI,
  bg = BG
)
message("[OK] saved fig_teacher_level_domain_means.png")

message("DONE.\nOutputs in: ", normalizePath(out_dir, winslash="/", mustWork = FALSE))
