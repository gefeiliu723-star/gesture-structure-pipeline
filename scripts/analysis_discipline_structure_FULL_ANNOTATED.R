############################################################
# FINAL / CLEAN / PAPER-READY ANALYSIS SCRIPT
# Gesture–Structure Alignment: Discipline × Structural Granularity
#
# WHAT THIS SCRIPT PRODUCES (all auto-saved):
# ----------------------------------------------------------
# exports/
#   figures/
#     Fig1_all_points_discipline_distribution.png
#     Fig2a_conf3_only_distribution.png
#     Fig2b_conf3_only_variability.png
#     Fig3_structural_granularity_comparison.png
#     FigS1_null_robustness_conf3.png
#
#   tables/
#     Table1_desc_all_points.csv
#     Table2_desc_conf3_only.csv
#     Table3_variability_conf3_only.csv
#     EffectSizes_cohens_d.csv
#
#   logs/
#     EXPORT_LOG.txt
#     RESULTS_SUMMARY.txt
#
# Key idea:
# - NOT “which discipline gestures more”
# - BUT “when/where do discipline differences emerge,
#        and do they show up as mean shifts or variability?”
############################################################


# ==========================================================
# SECTION 0: Setup
# ==========================================================

suppressPackageStartupMessages({
  library(readr)
  library(dplyr)
  library(ggplot2)
  library(effectsize)
  library(tibble)
})

# NOTE: Change only this if you move folders
project_root <- "D:/gesture_project"

# Input
in_csv <- file.path(project_root, "results", "time_enrichment_MASTER_ALL.csv")

# Output dirs
exports_dir <- file.path(project_root, "exports")
fig_dir     <- file.path(exports_dir, "figures")
tab_dir     <- file.path(exports_dir, "tables")
log_dir     <- file.path(exports_dir, "logs")

dir.create(fig_dir, recursive = TRUE, showWarnings = FALSE)
dir.create(tab_dir, recursive = TRUE, showWarnings = FALSE)
dir.create(log_dir, recursive = TRUE, showWarnings = FALSE)

# Safety check
stopifnot(file.exists(in_csv))

# ==========================================================
# GLOBAL PLOT THEME (publication-ready)
# ==========================================================

theme_pub <- theme_classic(base_size = 13) +
  theme(
    plot.title     = element_text(size = 14, face = "bold", hjust = 0.5),
    axis.title.y   = element_text(size = 12),
    axis.text      = element_text(size = 11),
    strip.background = element_rect(fill = "white", color = "black"),
    strip.text     = element_text(size = 11, face = "bold"),
    legend.position = "none"
  )


# ==========================================================
# SECTION 1: Load MASTER summary table
# ==========================================================

# NOTE:
# Each row = one video under one analysis configuration.
# Columns should include: tag, pointset, window_s, null_mode, lift, delta_rate

df_summary <- read_csv(in_csv, show_col_types = FALSE)

# Minimal required columns check
required_cols <- c("tag", "pointset", "window_s", "null_mode", "lift", "delta_rate")
missing_cols <- setdiff(required_cols, colnames(df_summary))
if (length(missing_cols) > 0) {
  stop(paste("Missing required columns:", paste(missing_cols, collapse = ", ")))
}


# ==========================================================
# SECTION 2: Add discipline labels (math vs lit)
# ==========================================================

# NOTE:
# We map video tags to discipline labels (math / lit).
discipline_map <- tibble(
  tag = c(
    paste0("m", sprintf("%02d", 1:10)),
    paste0("l", sprintf("%02d", 1:10))
  ),
  discipline = c(
    rep("math", 10),
    rep("lit",  10)
  )
)

df_summary2 <- df_summary %>%
  left_join(discipline_map, by = "tag")


# ==========================================================
# SECTION 3: Sanity checks (DO NOT SKIP)
# ==========================================================

# 3.1 Missing discipline?
missing_disc <- df_summary2 %>%
  filter(is.na(discipline)) %>%
  distinct(tag)

if (nrow(missing_disc) > 0) {
  stop(paste("Some tags are missing discipline labels:", paste(missing_disc$tag, collapse = ", ")))
}

# 3.2 Balanced split?
disc_counts <- df_summary2 %>%
  distinct(tag, discipline) %>%
  count(discipline)

# Expect math=10, lit=10
if (!all(sort(disc_counts$n) == c(10, 10))) {
  warning("Discipline split is not 10 vs 10. Check your tags or discipline_map.")
}

# 3.3 Save the joined master (useful for later)
out_master_disc <- file.path(project_root, "results", "time_enrichment_MASTER_ALL_with_discipline.csv")
write_csv(df_summary2, out_master_disc)


# ==========================================================
# SECTION 4: Helper functions (save plots / summarize)
# ==========================================================

save_plot <- function(p, filename, width, height, dpi = 300) {
  ggsave(
    filename = file.path(fig_dir, filename),
    plot = p,
    width = width,
    height = height,
    dpi = dpi
  )
}

desc_by_disc <- function(df) {
  df %>%
    group_by(discipline) %>%
    summarise(
      n = n(),
      mean_lift  = mean(lift),
      sd_lift    = sd(lift),
      mean_delta = mean(delta_rate),
      sd_delta   = sd(delta_rate),
      .groups = "drop"
    )
}

var_by_disc <- function(df) {
  df %>%
    group_by(discipline) %>%
    summarise(
      n = n(),
      var_lift = var(lift),
      sd_lift  = sd(lift),
      iqr_lift = IQR(lift),
      .groups = "drop"
    )
}

cohens_d_row <- function(df, label) {
  d_obj <- effectsize::cohens_d(lift ~ discipline, data = df)
  # effectsize returns a data.frame-like object; coerce safely:
  d_df <- as.data.frame(d_obj)
  d_df$analysis <- label
  d_df
}


# ==========================================================
# SECTION 5: Define analysis datasets
# ==========================================================

# NOTE:
# Main story configuration:
# all_points + window 1.5 + teacher_only null
df_main <- df_summary2 %>%
  filter(
    pointset == "all_points",
    window_s == 1.5,
    null_mode == "teacher_only"
  )

# Key structural points configuration:
df_conf3 <- df_summary2 %>%
  filter(
    pointset == "conf3_only",
    window_s == 1.5,
    null_mode == "teacher_only"
  )

# Granularity comparison (two pointsets)
df_compare <- df_summary2 %>%
  filter(
    window_s == 1.5,
    null_mode == "teacher_only",
    pointset %in% c("all_points", "conf3_only")
  ) %>%
  mutate(
    pointset = factor(
      pointset,
      levels = c("all_points", "conf3_only"),
      labels = c("All structural points", "Key structural points")
    )
  )

# Null robustness (conf3_only; teacher_only vs global)
df_robust <- df_summary2 %>%
  filter(
    pointset == "conf3_only",
    window_s == 1.5,
    null_mode %in% c("teacher_only", "global")
  )

# Sanity checks: expect 20 rows for df_main / df_conf3
if (nrow(df_main) != 20)  warning(paste("df_main row count is", nrow(df_main), "expected 20"))
if (nrow(df_conf3) != 20) warning(paste("df_conf3 row count is", nrow(df_conf3), "expected 20"))


# ==========================================================
# SECTION 6: Tables (auto-export)
# ==========================================================

# Table 1: Descriptives (all_points)
tab1 <- desc_by_disc(df_main)
write_csv(tab1, file.path(tab_dir, "Table1_desc_all_points.csv"))

# Table 2: Descriptives (conf3_only)
tab2 <- desc_by_disc(df_conf3)
write_csv(tab2, file.path(tab_dir, "Table2_desc_conf3_only.csv"))

# Table 3: Variability (conf3_only)
tab3 <- var_by_disc(df_conf3)
write_csv(tab3, file.path(tab_dir, "Table3_variability_conf3_only.csv"))

# Effect sizes
d_main  <- cohens_d_row(df_main,  "all_points_teacher_only")
d_conf3 <- cohens_d_row(df_conf3, "conf3_only_teacher_only")
d_all   <- bind_rows(d_main, d_conf3)
write_csv(d_all, file.path(tab_dir, "EffectSizes_cohens_d.csv"))


# ==========================================================
# SECTION 7: Figures (auto-export)
# ==========================================================

# --------------------------
# Figure 1: all_points distribution
# --------------------------
# NOTE:
# Why this plot?
# - Means hide overlap; this shows full distribution.
# How to read:
# - Overlap high => modest discipline separation
# - Separation clear => stronger discipline effect

p_fig1 <- ggplot(df_main, aes(x = discipline, y = lift)) +
  geom_violin(trim = FALSE, alpha = 0.35) +
  geom_jitter(width = 0.08, height = 0, size = 2, alpha = 0.8) +
  stat_summary(fun = mean, geom = "point", size = 3.5, color = "black") +
  labs(
    x = NULL,
    y = "Gesture–Structure Enrichment (Lift)",
    title = "Discipline-Level Differences in Gesture–Structure Alignment (All structural points)"
  ) +
  theme_pub

save_plot(p_fig1, "Fig1_all_points_discipline_distribution.png", width = 6.5, height = 4.5)


# --------------------------
# Figure 2a: conf3_only distribution
# --------------------------
# NOTE:
# Why this plot?
# - Tests whether discipline differences emerge at key structural moments.
# How to read:
# - Compare to Fig1: larger separation or different spread => structure-conditional effect

p_fig2a <- ggplot(df_conf3, aes(x = discipline, y = lift)) +
  geom_violin(trim = FALSE, alpha = 0.35) +
  geom_jitter(width = 0.08, height = 0, size = 2, alpha = 0.8) +
  stat_summary(fun = mean, geom = "point", size = 3.5, color = "black") +
  labs(
    x = NULL,
    y = "Gesture–Structure Enrichment (Lift)",
    title = "Discipline Differences at Key Structural Points (conf3_only)"
  ) +
  theme_pub

save_plot(p_fig2a, "Fig2a_conf3_only_distribution.png", width = 6.5, height = 4.5)


# --------------------------
# Figure 2b: conf3_only variability (boxplot)
# --------------------------
# NOTE:
# Why this plot?
# - Focuses on variability: who is more stylistically diverse?
# How to read:
# - Larger IQR/whiskers => greater variation across videos within a discipline

p_fig2b <- ggplot(df_conf3, aes(x = discipline, y = lift)) +
  geom_boxplot(width = 0.45, outlier.shape = NA) +
  geom_jitter(width = 0.08, height = 0, alpha = 0.7, size = 2) +
  labs(
    x = NULL,
    y = "Gesture–Structure Enrichment (Lift)",
    title = "Variability by Discipline at Key Structural Points (conf3_only)"
  ) +
  theme_pub

save_plot(p_fig2b, "Fig2b_conf3_only_variability.png", width = 6.0, height = 4.5)


# --------------------------
# Figure 3: structural granularity comparison (facet)
# --------------------------
# NOTE:
# Why this plot?
# - Directly tests whether discipline differences depend on how “structure” is defined.
# How to read:
# - If the discipline pattern changes across panels => interaction:
#   discipline differences are conditional on structural granularity.

p_fig3 <- ggplot(df_compare, aes(x = discipline, y = lift)) +
  geom_violin(trim = FALSE, alpha = 0.35) +
  geom_jitter(width = 0.08, height = 0, size = 2, alpha = 0.8) +
  stat_summary(fun = mean, geom = "point", size = 3, color = "black") +
  facet_wrap(~ pointset) +
  labs(
    x = NULL,
    y = "Gesture–Structure Enrichment (Lift)",
    title = "Discipline Differences Depend on Structural Granularity"
  ) +
  theme_pub

save_plot(p_fig3, "Fig3_structural_granularity_comparison.png", width = 8.0, height = 4.5)


# --------------------------
# Supplement Figure S1: null robustness
# --------------------------
# NOTE:
# Why this plot?
# - Checks whether conclusions depend on null baseline choice.
# How to read:
# - Similar pattern across facets => robust to null model.

p_figS1 <- ggplot(df_robust, aes(x = discipline, y = lift)) +
  geom_boxplot(width = 0.45, outlier.shape = NA) +
  geom_jitter(width = 0.08, height = 0, alpha = 0.7, size = 2) +
  facet_wrap(~ null_mode) +
  labs(
    x = NULL,
    y = "Gesture–Structure Enrichment (Lift)",
    title = "Robustness Across Null Models (conf3_only)"
  ) +
  theme_pub

save_plot(p_figS1, "FigS1_null_robustness_conf3.png", width = 8.0, height = 4.5)


# ==========================================================
# SECTION 8: EXPORT_LOG.txt (what each figure answers)
# ==========================================================

export_log_path <- file.path(log_dir, "EXPORT_LOG.txt")

log_lines <- c(
  "============================================================",
  "EXPORT LOG: Gesture–Structure Alignment Figures",
  paste0("Timestamp: ", format(Sys.time(), "%Y-%m-%d %H:%M:%S")),
  paste0("Project root: ", project_root),
  "============================================================",
  "",
  "GENERAL NOTE:",
  "- Each dot represents one video (N=20; 10 math, 10 lit) under a fixed configuration.",
  "- Lift: higher => gesture is more concentrated around structural points than expected under the null.",
  "- Overlap (math vs lit) => modest discipline separation.",
  "- Differences in spread (SD/IQR) => differences in stylistic variability rather than uniform mean shifts.",
  "",
  "FIGURE 1: Fig1_all_points_discipline_distribution.png",
  "  QUESTION: Do disciplines differ when ALL structural points are included?",
  "  READ: compare overlap + mean; heavy overlap => modest effect.",
  "",
  "FIGURE 2a: Fig2a_conf3_only_distribution.png",
  "  QUESTION: Do discipline differences become clearer at KEY structural points (conf3_only)?",
  "  READ: compare to Fig1; increased separation or changed spread => structure-conditional differences.",
  "",
  "FIGURE 2b: Fig2b_conf3_only_variability.png",
  "  QUESTION: Is within-discipline variability different at key structural points?",
  "  READ: wider box / longer whiskers => greater stylistic diversity.",
  "",
  "FIGURE 3: Fig3_structural_granularity_comparison.png",
  "  QUESTION: Do discipline differences depend on structural granularity (all vs key points)?",
  "  READ: if the pattern changes across panels => interaction with structural definition.",
  "",
  "SUPPLEMENT S1: FigS1_null_robustness_conf3.png",
  "  QUESTION: Are results robust to null choice (teacher_only vs global)?",
  "  READ: similar shapes across facets => robustness; big differences => revisit baseline assumptions.",
  "",
  "============================================================",
  "END OF EXPORT LOG",
  "============================================================"
)

writeLines(log_lines, con = export_log_path)


# ==========================================================
# SECTION 9: RESULTS_SUMMARY.txt (numbers + ready-to-write claims)
# ==========================================================

results_path <- file.path(log_dir, "RESULTS_SUMMARY.txt")

# Pull key numbers (safe indexing by discipline)
get_stat <- function(tab, disc, col) {
  tab %>% filter(discipline == disc) %>% pull({{col}})
}

# All points
m_mean_all <- get_stat(tab1, "math", mean_lift)
m_sd_all   <- get_stat(tab1, "math", sd_lift)
l_mean_all <- get_stat(tab1, "lit",  mean_lift)
l_sd_all   <- get_stat(tab1, "lit",  sd_lift)

# conf3
m_mean_c3 <- get_stat(tab2, "math", mean_lift)
m_sd_c3   <- get_stat(tab2, "math", sd_lift)
l_mean_c3 <- get_stat(tab2, "lit",  mean_lift)
l_sd_c3   <- get_stat(tab2, "lit",  sd_lift)

# variability conf3
m_iqr_c3 <- get_stat(tab3, "math", iqr_lift)
l_iqr_c3 <- get_stat(tab3, "lit",  iqr_lift)

# effect sizes (grab estimate column name robustly)
d_col <- if ("Cohens_d" %in% names(d_all)) "Cohens_d" else if ("Cohens_d" %in% tolower(names(d_all))) names(d_all)[tolower(names(d_all))=="cohens_d"][1] else names(d_all)[1]
# safer: pick the numeric-looking column
num_cols <- names(d_all)[sapply(d_all, is.numeric)]
d_est_col <- if (length(num_cols) > 0) num_cols[1] else names(d_all)[1]

d_allpoints <- d_all %>% filter(analysis == "all_points_teacher_only") %>% pull(!!sym(d_est_col))
d_conf3only <- d_all %>% filter(analysis == "conf3_only_teacher_only") %>% pull(!!sym(d_est_col))

summary_lines <- c(
  "============================================================",
  "RESULTS SUMMARY (auto-generated)",
  paste0("Timestamp: ", format(Sys.time(), "%Y-%m-%d %H:%M:%S")),
  "============================================================",
  "",
  "MAIN CONFIGURATION:",
  "- all_points, window = 1.5s, null = teacher_only",
  paste0("Math: mean lift = ", round(m_mean_all, 3), ", SD = ", round(m_sd_all, 3)),
  paste0("Lit : mean lift = ", round(l_mean_all, 3), ", SD = ", round(l_sd_all, 3)),
  paste0("Cohen's d (all_points) = ", round(d_allpoints, 3)),
  "",
  "KEY STRUCTURAL POINTS CONFIGURATION:",
  "- conf3_only, window = 1.5s, null = teacher_only",
  paste0("Math: mean lift = ", round(m_mean_c3, 3), ", SD = ", round(m_sd_c3, 3)),
  paste0("Lit : mean lift = ", round(l_mean_c3, 3), ", SD = ", round(l_sd_c3, 3)),
  paste0("Cohen's d (conf3_only) = ", round(d_conf3only, 3)),
  "",
  "VARIABILITY (conf3_only):",
  paste0("Math IQR = ", round(m_iqr_c3, 3)),
  paste0("Lit  IQR = ", round(l_iqr_c3, 3)),
  "",
  "READY-TO-WRITE INTERPRETATION (edit wording as needed):",
  "- Compare Fig1 vs Fig2a: if discipline separation increases at conf3_only, differences are structure-conditional.",
  "- If Lit shows larger SD/IQR at conf3_only, interpret as greater stylistic flexibility / structural openness.",
  "- Fig3 visualizes the dependence of discipline differences on structural granularity (interaction-like pattern).",
  "- FigS1 checks robustness: similar patterns across nulls support baseline-independence of conclusions.",
  "",
  "FILES EXPORTED:",
  paste0("Figures: ", fig_dir),
  paste0("Tables : ", tab_dir),
  paste0("Logs   : ", log_dir),
  "============================================================"
)

writeLines(summary_lines, con = results_path)


# ==========================================================
# SECTION 10: Final console messages (so you know it worked)
# ==========================================================

message("DONE ✅")
message("Master w/ discipline saved: ", out_master_disc)
message("Figures saved to: ", fig_dir)
message("Tables saved to : ", tab_dir)
message("Logs saved to   : ", log_dir)
message("EXPORT_LOG: ", export_log_path)
message("RESULTS_SUMMARY: ", results_path)

# ==========================================================
# SECTION 11: FIGURE CAPTIONS (journal-ready)
# ==========================================================

captions_path <- file.path(log_dir, "CAPTIONS.txt")

caption_lines <- c(
  "============================================================",
  "FIGURE CAPTIONS",
  paste0("Generated: ", format(Sys.time(), "%Y-%m-%d %H:%M:%S")),
  "============================================================",
  "",
  "Figure 1.",
  "Discipline-level differences in gesture–structure alignment across all structural points.",
  "Each point represents one video (N = 20; 10 mathematics, 10 literature).",
  "The violin plots depict the distribution of gesture–structure enrichment (lift),",
  "with black dots indicating discipline means.",
  "Substantial overlap between distributions indicates modest discipline-level separation",
  "when alignment is assessed across all structural points.",
  "",
  "Figure 2a.",
  "Discipline differences in gesture–structure alignment at key structural points (conf3_only).",
  "Each point represents one video, aligned to a ±1.5 s window around key structural moments",
  "under a teacher-only null model.",
  "Compared to Figure 1, differences in both central tendency and dispersion suggest that",
  "discipline effects are more pronounced at moments of strong structural emphasis.",
  "",
  "Figure 2b.",
  "Variability in gesture–structure alignment across disciplines at key structural points.",
  "Boxplots summarize the distribution of lift values (conf3_only), with overlaid points",
  "representing individual videos.",
  "Greater spread (IQR and whiskers) indicates higher stylistic variability in how structure",
  "is marked through gesture, particularly in literature instruction.",
  "",
  "Figure 3.",
  "Dependence of discipline differences on structural granularity.",
  "Gesture–structure alignment is shown separately for all structural points (left panel)",
  "and key structural points (right panel), using the same scale and null model.",
  "Changes in distributional patterns across panels indicate that discipline differences",
  "are conditional on how instructional structure is defined.",
  "",
  "Figure S1.",
  "Robustness of discipline differences across null models.",
  "Gesture–structure enrichment at key structural points (conf3_only) is shown under",
  "teacher-only and global null baselines.",
  "Qualitatively similar patterns across null models indicate that observed discipline",
  "differences are not driven by the specific choice of baseline.",
  "",
  "============================================================",
  "END OF CAPTIONS",
  "============================================================"
)

writeLines(caption_lines, con = captions_path)

message("Wrote CAPTIONS file: ", captions_path)
